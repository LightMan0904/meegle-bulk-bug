#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
本地图形界面：飞书项目批量建单器

- 只监听 127.0.0.1，数据不出本机
- 复用 meegle_bulk.py 的字段解析/值转换/CLI 调用逻辑
- 登录态用的是运行者本人的 meegle auth login

启动：python webapp/server.py   （或双击 start_gui.bat）
"""
from __future__ import annotations

import base64
import csv
from concurrent.futures import ThreadPoolExecutor
import io
import json
import mimetypes
import re
import secrets
import socket
import sys
import threading
import time
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, urlparse

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT))

import meegle_bulk as core  # noqa: E402

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    Workbook = None

STATIC = HERE / "static"
TOKEN = secrets.token_urlsafe(18)

CFG = {
    "project_key": "",
    "work_item_type": "",
    "rate_limit_seconds": 0.4,
    "max_retries": 4,
    "template_name": "",
    "host": "project.feishu.cn",
}
UPLOADS: dict[str, dict] = {}
METAS: dict[str, dict] = {}
META_TS: dict[str, float] = {}
JOBS: dict[str, dict] = {}
LOCK = threading.Lock()
LAYOUT_DIR = ROOT / "layouts"
META_TTL = 600.0


def layout_path(pk: str, wt: str) -> Path:
    LAYOUT_DIR.mkdir(exist_ok=True)
    return LAYOUT_DIR / f"{pk}__{wt}.json"


def load_layout(pk: str, wt: str) -> dict | None:
    p = layout_path(pk, wt)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            return None
    # 兼容别名：URL 里是 simple_name（如 your-space-key），接口用的是 project_key（67e209...）。
    # 同一工作项类型只存了一份布局时，直接复用，避免“字段清单又退回 10 个”。
    cands = [f for f in LAYOUT_DIR.glob(f"*__{wt}.json") if f.name != p.name]
    if len(cands) == 1:
        try:
            return json.loads(cands[0].read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            return None
    return None


def save_layout(pk: str, wt: str, items: list) -> dict:
    data = {"project_key": pk, "work_item_type": wt,
            "saved_at": time.strftime("%Y-%m-%d %H:%M:%S"), "items": items}
    layout_path(pk, wt).write_text(json.dumps(data, ensure_ascii=False, indent=2),
                                   encoding="utf-8")
    return data


PROFILE_DIR = ROOT / "profiles"
PROFILE_TTL = 86400.0


def profiles_path(pk: str, wt: str) -> Path:
    PROFILE_DIR.mkdir(exist_ok=True)
    return PROFILE_DIR / f"{pk}__{wt}.json"


def load_profiles(pk: str, wt: str) -> dict | None:
    p = profiles_path(pk, wt)
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return None


def _probe_required(type_oid: str, tpl_oid: str, pk: str, wt: str) -> list[str] | None:
    """发一个「故意缺字段」的创建请求，从 ErrFieldRequired 里解析该类型的必填字段。

    故意不传 role_owners（解决人在三种测试单类型下都是必填），保证不会真的建单。
    """
    cfg = dict(CFG, project_key=pk, work_item_type=wt)
    fields = [
        {"field_key": "name", "field_value": "ZZZ-必填探测-勿用"},
        {"field_key": "template", "field_value": tpl_oid},
        {"field_key": "field_8e9f64", "field_value": type_oid},
    ]
    try:
        core.run_cli(cfg, ["workitem", "create", "--project-key", pk,
                           "--work-item-type", wt,
                           "--params", json.dumps({"fields": fields}, ensure_ascii=False)],
                     timeout=120)
        return None  # 不应该成功；成功说明探测不安全
    except core.CliError as e:
        msg = str(e)
        groups = re.findall(r"\(([A-Za-z0-9_,]+)\)\s*必填", msg)
        if not groups:
            return []
        return max(groups, key=len).split(",")


def discover_profiles(meta: dict, pk: str, wt: str, items: list[dict]) -> dict | None:
    allf = all_fields(meta)
    type_item = next((it for it in items if it["field_key"] == "field_8e9f64"), None)
    tpl_item = next((it for it in items if it["field_key"] == "template"), None)
    if not type_item or not tpl_item or not type_item.get("options"):
        return None
    tpl_oid = None
    default_label = (tpl_item.get("default") or "")
    for o in tpl_item.get("options", []):
        if core.norm_name(o["label"]) == core.norm_name(default_label):
            tpl_oid = o["value"]
            break
    if not tpl_oid and tpl_item.get("options"):
        tpl_oid = tpl_item["options"][0]["value"]
    if not tpl_oid:
        return None

    prof = {}
    for o in type_item["options"][:6]:
        keys = _probe_required(o["value"], tpl_oid, pk, wt)
        if keys is None:
            return None
        prof[o["value"]] = {
            "label": o["label"],
            "required": [k for k in keys if k in allf],
            "role_required": [k for k in keys if k.startswith("role_")],
        }
    return {"checked_at": time.time(), "template": tpl_oid, "profiles": prof}


def get_profiles(meta: dict, pk: str, wt: str, items: list[dict],
                 refresh: bool = False) -> dict:
    cached = load_profiles(pk, wt)
    if (not refresh and cached
            and time.time() - float(cached.get("checked_at") or 0) < PROFILE_TTL):
        return cached
    try:
        prof = discover_profiles(meta, pk, wt, items)
        if prof:
            profiles_path(pk, wt).write_text(json.dumps(prof, ensure_ascii=False, indent=2),
                                             encoding="utf-8")
            return prof
    except Exception:  # noqa: BLE001
        pass
    return cached or {"profiles": {}}


def form_items(meta: dict, pk: str, wt: str):
    layout = load_layout(pk, wt)
    if layout and layout.get("items"):
        return core.resolve_form_items(meta, layout["items"],
                                       layout.get("project_key") or pk,
                                       layout.get("work_item_type") or wt)
    # 没有布局：退回「创建表单可写字段」
    fields = core.order_fields(core.importable_fields(meta))
    items = [{
        "kind": "field", "field_key": f["field_key"], "field_name": f["field_name"],
        "field_type": f.get("field_type", ""), "required": bool(f.get("required")),
        "options": f.get("options", []), "related": f.get("related") or [],
        "writable": True, "note": "", "default": "", "is_new": False,
    } for f in fields]
    return items, []


# ----------------------------------------------------------------- 工具


def ok(data=None):
    return {"ok": True, "data": data}


def err(msg: str):
    return {"ok": False, "error": str(msg)}


def meta_key(pk: str, wt: str) -> str:
    return f"{pk}::{wt}"


def all_fields(meta: dict) -> dict:
    return core.merge_defs(core.extract_field_defs(meta.get("create_fields")),
                           core.extract_field_defs(meta.get("fields")))


def find_user_name(resp) -> str:
    for d in core.walk(resp):
        if isinstance(d, dict):
            for k in ("name", "user_name", "en_name"):
                v = d.get(k)
                if isinstance(v, str) and v and k in d:
                    return v
    return ""


def get_meta(pk: str, wt: str, refresh: bool = False) -> dict:
    key = meta_key(pk, wt)
    if (not refresh and key in METAS
            and time.time() - META_TS.get(key, 0) < META_TTL):
        return METAS[key]
    cfg = dict(CFG, project_key=pk, work_item_type=wt)
    meta = core.fetch_meta(cfg)
    METAS[key] = meta
    META_TS[key] = time.time()
    return meta


def build_cols(meta: dict, mapping: dict[str, str], headers: list[str],
               items: list[dict] | None = None) -> tuple[dict, list[str]]:
    """mapping: {列下标(字符串): field_key 或 role:xxx} -> core 需要的 cols 结构"""
    fields = all_fields(meta)
    by_key = {it["field_key"]: it for it in (items or [])}
    cols, problems = {}, []
    for idx_s, fkey in (mapping or {}).items():
        if not fkey:
            continue
        try:
            idx = int(idx_s)
        except ValueError:
            continue
        header = headers[idx] if idx < len(headers) else f"列{idx + 1}"
        item = by_key.get(fkey)

        if str(fkey).startswith("role:"):
            if not item:
                problems.append(f"角色 {fkey} 不在当前页面布局里")
                continue
            cols[header] = {
                "index": idx, "field_key": fkey,
                "field_name": item["field_name"], "field_type": "role",
                "mode": "multi_user", "options": {},
                "required": bool(item.get("required")),
                "default": item.get("default", ""), "note": item.get("note", ""),
                "kind": "role", "role_id": item.get("role_id"),
                "api_role_key": item.get("api_role_key") or fkey.replace("role:", ""),
                "writable": True,
            }
            continue

        field = fields.get(fkey)
        if not field:
            problems.append(f"字段 {fkey} 不存在")
            continue
        opts = {}
        paths = {}
        for o in field.get("options", []):
            opts[core.norm_name(o["label"])] = o["value"]
            opts[core.norm_name(o["value"])] = o["value"]
            if o.get("path"):
                paths[core.norm_name(o["label"])] = o["path"]
        ftype = core.canon_type(field.get("field_type"))
        cols[header] = {
            "index": idx,
            "field_key": field["field_key"],
            "field_name": field["field_name"],
            "field_type": field.get("field_type", ""),
            "mode": core.field_mode(field),
            "options": opts,
            "related": field.get("related") or [],
            "tree": ftype in ("tree-select", "tree-multi-select"),
            "paths": paths,
            "required": bool(field.get("required")) or bool((item or {}).get("required")),
            "default": (item or {}).get("default", ""),
            "note": (item or {}).get("note", ""),
            "kind": "field",
            "writable": (item or {}).get("writable", True),
        }
    return cols, problems


def parse_table(data: bytes, filename: str, sheet: str | None = None):
    """返回 (sheets, sheet_name, rows[list[list[str]]])"""
    lower = (filename or "").lower()
    if lower.endswith(".csv"):
        text = None
        for enc in ("utf-8-sig", "gb18030", "utf-8"):
            try:
                text = data.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise ValueError("CSV 编码无法识别，请另存为 UTF-8 或 xlsx")
        rows = [[str(c).strip() for c in r] for r in csv.reader(io.StringIO(text))]
        return ["CSV"], "CSV", [r for r in rows if any(x for x in r)]
    if load_workbook is None:
        raise RuntimeError("缺少 openpyxl，请先 pip install openpyxl")
    wb = load_workbook(io.BytesIO(data), data_only=True)
    names = wb.sheetnames
    name = sheet if sheet in names else names[0]
    ws = wb[name]
    rows = []
    for r in ws.iter_rows(values_only=True):
        row = [core.cell_str(c) for c in r]
        if any(row):
            rows.append(row)
    return names, name, rows


def guess_mapping(headers: list[str], meta: dict) -> dict[str, str]:
    fields = all_fields(meta)
    out = {}
    for i, h in enumerate(headers):
        f = core.guess_field(h, fields) if h else None
        if f:
            out[str(i)] = f["field_key"]
    return out


# ----------------------------------------------------------------- 业务 API


def api_status(_payload):
    try:
        st = core.run_cli(CFG, ["auth", "status"], check=False)
        authed = bool(st.get("authenticated")) and not st.get("__error__")
        user = ""
        if authed:
            me = core.run_cli(CFG, ["user", "me"], check=False)
            user = find_user_name(me)
        return ok({"authenticated": authed, "user": user, "raw": st})
    except core.CliError as e:
        return ok({"authenticated": False, "user": "", "error": str(e)})


def api_login_start(payload):
    """设备码登录第 1 步：拿 user_code / 二维码链接（非交互，不会弹终端菜单）"""
    host = (payload.get("host") or CFG.get("host") or "project.feishu.cn").strip()
    if not re.fullmatch(r"[A-Za-z0-9.:-]{4,80}", host):
        return err("登录环境不合法")
    CFG["host"] = host
    resp = core.run_cli(CFG, ["auth", "login", "--device-code", "--host", host, "--phase", "init"],
                        check=False, timeout=90)
    if resp.get("__error__"):
        return err(resp["__error__"])
    if not resp.get("device_code"):
        return err("生成二维码失败：" + json.dumps(resp, ensure_ascii=False)[:300])
    return ok({
        "device_code": resp.get("device_code"),
        "user_code": resp.get("user_code"),
        "client_id": resp.get("client_id"),
        "interval": resp.get("interval") or 5,
        "expires_in": resp.get("expires_in") or 1800,
        "verification_uri": resp.get("verification_uri"),
        "verification_uri_complete": resp.get("verification_uri_complete"),
    })


def api_login_poll(payload):
    """设备码登录第 2 步：单次轮询，前端按 interval 调"""
    dc = payload.get("device_code") or ""
    cid = payload.get("client_id") or ""
    if not dc or not cid:
        return err("缺少 device_code / client_id")
    resp = core.run_cli(CFG, ["auth", "login", "--device-code", "--phase", "poll",
                              "--device-code-value", dc, "--client-id", cid, "--once"],
                        check=False, timeout=60)
    if resp.get("__error__"):
        return ok({"status": "error", "error": _cli_error_message(resp["__error__"])})
    status = str(resp.get("status") or "").lower()
    if status in ("success", "authorized", "ok", "done"):
        st = core.run_cli(CFG, ["auth", "status"], check=False)
        if st.get("authenticated") and not st.get("__error__"):
            user = find_user_name(core.run_cli(CFG, ["user", "me"], check=False))
            return ok({"status": "success", "user": user})
        return ok({"status": "error", "error": "授权成功但本地凭证没保存下来，请点「刷新二维码」重试"})
    return ok({"status": status or "unknown"})


def _cli_error_message(raw: str) -> str:
    """把 CLI 的 JSON 错误压成一句中文，方便界面展示"""
    text = str(raw)
    try:
        start, end = text.find("{"), text.rfind("}")
        j = json.loads(text[start:end + 1])
        e = j.get("error") or {}
        return str(e.get("message") or j.get("status") or text)[:300]
    except Exception:  # noqa: BLE001
        return text[:300]


def api_projects(_payload):
    resp = core.run_cli(CFG, ["project", "search", "--page-num", "1"], check=False)
    if resp.get("__error__"):
        return ok({"projects": [], "error": resp["__error__"]})
    seen, items = set(), []
    for d in core.walk(resp):
        pk = core.pick(d, ("project_key", "key", "simple_name"))
        name = core.pick(d, ("name", "project_name", "display_name"))
        if pk and name and str(pk) not in seen:
            seen.add(str(pk))
            items.append({"key": str(pk), "name": str(name)})
    return ok({"projects": items})


def api_types(payload):
    pk = (payload.get("project_key") or "").strip()
    if not pk:
        return err("请先填空间 key")
    resp = core.run_cli(CFG, ["workitem", "meta-types", "--project-key", pk])
    return ok({"types": core.extract_types(resp)})


def api_meta(payload):
    pk = (payload.get("project_key") or "").strip()
    wt = (payload.get("work_item_type") or "").strip()
    if not pk or not wt:
        return err("请选择空间和工作项类型")
    meta = get_meta(pk, wt, refresh=bool(payload.get("refresh")))
    global CFG
    CFG = dict(CFG, project_key=pk, work_item_type=wt)
    items, missing = form_items(meta, pk, wt)
    # 后台预热关联字段的目标列表（项目/研发工单/…），下拉即开即用
    try:
        threading.Thread(target=_preload_relations, args=(items,), daemon=True).start()
    except Exception:  # noqa: BLE001
        pass
    everything = all_fields(meta)
    tpl = everything.get("template", {})
    prof = get_profiles(meta, pk, wt, items, refresh=bool(payload.get("refresh")))
    return ok({
        "profiles": prof.get("profiles") or {},
        "profile_checked_at": prof.get("checked_at"),
        "project_key": pk,
        "work_item_type": wt,
        "items": items,
        "missing": missing,
        "new_fields": [i["field_name"] for i in items if i.get("is_new")],
        "layout_saved": bool(load_layout(pk, wt)),
        "all_field_count": len(everything),
        "create_field_count": len(core.create_field_defs(meta)),
        "template_options": [o["label"] for o in tpl.get("options", [])],
        "fetched_at": meta.get("fetched_at"),
    })


def api_template(payload):
    pk = (payload.get("project_key") or "").strip()
    wt = (payload.get("work_item_type") or "").strip()
    meta = get_meta(pk, wt)
    items, _missing = form_items(meta, pk, wt)
    order = payload.get("order")
    if isinstance(order, list) and order:
        rank = {str(k): i for i, k in enumerate(order)}
        original = {it["field_key"]: i for i, it in enumerate(items)}
        items.sort(key=lambda it: rank.get(it["field_key"], len(order) + original.get(it["field_key"], 0)))
        save_layout(pk, wt, [{
            "kind": it["kind"],
            "key": it.get("role_id") or it["field_key"],
            "name": it["field_name"], "required": bool(it.get("required")),
            "note": it.get("note", ""), "default": it.get("default", ""),
        } for it in items])
    prof = get_profiles(meta, pk, wt, items)
    profiles = prof.get("profiles") or {}
    if profiles:
        req_sets = [set(p.get("required") or []) for p in profiles.values()]
        common = set.intersection(*req_sets) if req_sets else set()
        common |= {"name", "template", "field_8e9f64"}
        everyone = set.union(*req_sets) if req_sets else set()
        role_sets = [set(p.get("role_required") or []) for p in profiles.values()]
        role_common = set.intersection(*role_sets) if role_sets else set()
        for it in items:
            k = it["field_key"]
            if it.get("kind") == "role":
                it["required"] = (it.get("api_role_key_full") in role_common)
            else:
                it["required"] = k in common
                if k in everyone - common:
                    it["note"] = (it.get("note") or "") + "（按测试单类型必填）"
    wb, _cols = core.build_template_items(items, meta, wt or "数据", profiles)
    buf = io.BytesIO()
    wb.save(buf)
    return {"__file__": buf.getvalue(),
            "filename": f"批量建单模板_{wt}.xlsx",
            "content_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}


def _parse_query_items(resp) -> list[dict]:
    out = []
    data = resp.get("data") or {}
    groups = data.values() if isinstance(data, dict) else []
    for group in groups:
        for rec in (group or []):
            vals = {}
            for f in rec.get("moql_field_list") or []:
                v = f.get("value") or {}
                if "long_value" in v:
                    vals[f.get("key")] = v["long_value"]
                elif "string_value" in v:
                    vals[f.get("key")] = v["string_value"]
            if vals.get("work_item_id") and vals.get("name"):
                out.append({"id": str(vals["work_item_id"]), "name": str(vals["name"])})
    return out


REL_CACHE: dict[tuple, dict] = {}
REL_TTL = 900.0
REL_LIMIT = 300
REL_LOCK = threading.Lock()
REL_LOCKS: dict[tuple, threading.Lock] = {}


def _relation_fetch(project: str, target: str, limit: int = REL_LIMIT):
    mql = f"SELECT `work_item_id`, `name` FROM `{project}`.`{target}` LIMIT {limit}"
    resp = core.run_cli(CFG, ["workitem", "query", "--project-key", project, "--mql", mql],
                        timeout=90)
    items = _parse_query_items(resp)
    total = None
    for g in resp.get("list") or []:
        if isinstance(g, dict) and isinstance(g.get("count"), int):
            total = g["count"]
    # 接口单页最多 50 条；用 count 判断列表是否完整
    complete = (len(items) >= total) if total is not None else (len(items) < limit)
    return items, complete


def _relation_cache(project: str, target: str) -> dict:
    key = (project, target)
    with REL_LOCK:
        c = REL_CACHE.get(key)
        if c and time.time() - c["ts"] < REL_TTL:
            return c
        lock = REL_LOCKS.setdefault(key, threading.Lock())
    with lock:
        with REL_LOCK:
            c = REL_CACHE.get(key)
            if c and time.time() - c["ts"] < REL_TTL:
                return c
        items, full = _relation_fetch(project, target)
        c = {"ts": time.time(), "items": items, "full": full}
        with REL_LOCK:
            REL_CACHE[key] = c
        return c


def _preload_relations(items: list[dict]):
    """后台把关联字段的目标类型列表缓存起来，打开下拉时就不用再等 CLI。"""
    seen = []
    for it in items or []:
        for rel in it.get("related") or []:
            key = (rel.get("project_key"), rel.get("work_item_type"))
            if not all(key) or key in seen:
                continue
            seen.append(key)
    if not seen:
        return
    def one(key):
        try:
            _relation_cache(*key)
        except Exception:  # noqa: BLE001
            pass
    with ThreadPoolExecutor(max_workers=3) as ex:
        list(ex.map(one, seen))


def api_relation_options(payload):
    """按关键字搜索关联字段可选的工作项（优先走本地缓存，毫秒级）"""
    project = (payload.get("project_key") or CFG["project_key"] or "").strip()
    target = (payload.get("target_type") or "").strip()
    q = (payload.get("q") or "").strip()
    if not project or not target:
        return err("缺少 project_key / target_type")
    try:
        c = _relation_cache(project, target)
    except core.CliError as e:
        return err(str(e)[:400])

    if not q:
        return ok({"items": c["items"][:200], "cached": True, "complete": c["full"]})

    ql = q.casefold()
    hits = [i for i in c["items"] if ql in i["name"].casefold()][:20]
    # full=True 表示列表已经完整，本地没有就是没有，不用再回源
    if hits or c["full"]:
        return ok({"items": hits, "cached": True, "complete": c["full"]})

    # 列表被截断且缓存里没有：回源做一次 LIKE 搜索
    safe = q.replace("'", "''")
    mql = (f"SELECT `work_item_id`, `name` FROM `{project}`.`{target}` "
           f"WHERE `name` LIKE '%{safe}%' LIMIT 20")
    try:
        resp = core.run_cli(CFG, ["workitem", "query", "--project-key", project,
                                  "--mql", mql], timeout=90)
    except core.CliError as e:
        return err(str(e)[:400])
    return ok({"items": _parse_query_items(resp), "cached": False, "complete": False})


def _apply_relations(rows, cols, relations):
    """relations: {行下标: {field_key: 工作项ID}}，覆盖到行数据里"""
    if not relations:
        return rows
    rows = [list(r) for r in rows]
    key_to_col = {c["field_key"]: c for c in (cols or {}).values()}
    for ridx, m in relations.items():
        try:
            ri = int(ridx)
        except (TypeError, ValueError):
            continue
        if ri < 0 or ri >= len(rows):
            continue
        for fkey, val in (m or {}).items():
            col = key_to_col.get(fkey)
            if not col:
                continue
            row = rows[ri]
            while len(row) <= col["index"]:
                row.append("")
            row[col["index"]] = val
    return rows


def api_layout(payload):
    """保存页面字段顺序（含必填/说明/默认值）"""
    pk = (payload.get("project_key") or "").strip()
    wt = (payload.get("work_item_type") or "").strip()
    items = payload.get("items")
    if not pk or not wt or not isinstance(items, list) or not items:
        return err("参数不完整")
    clean = []
    for it in items:
        kind = it.get("kind") or "field"
        key = it.get("key") or it.get("field_key")
        if not key:
            continue
        clean.append({
            "kind": kind, "key": str(key),
            "name": it.get("name") or it.get("field_name") or "",
            "required": bool(it.get("required")),
            "note": it.get("note", ""), "default": it.get("default", ""),
        })
    save_layout(pk, wt, clean)
    return ok({"saved": True, "count": len(clean)})


def api_upload(payload):
    name = payload.get("name") or "upload.xlsx"
    raw = base64.b64decode(payload.get("data_b64") or "")
    sheet = payload.get("sheet")
    sheets, used, rows = parse_table(raw, name, sheet)
    if not rows:
        return err("表格是空的")
    headers = [str(h) for h in rows[0]]
    body = rows[1:]
    uid = secrets.token_hex(6)
    UPLOADS[uid] = {"headers": headers, "rows": body, "name": name, "sheet": used, "sheets": sheets}
    pk = payload.get("project_key") or CFG["project_key"]
    wt = payload.get("work_item_type") or CFG["work_item_type"]
    meta = METAS.get(meta_key(pk, wt)) if pk and wt else None
    mapping = guess_mapping(headers, meta) if meta else {}
    if meta:
        items, _m = form_items(meta, pk, wt)
        role_by_name = {core.norm_name(it["field_name"]): it["field_key"]
                        for it in items if it.get("kind") == "role"}
        for i, h in enumerate(headers):
            if str(i) in mapping or not h:
                continue
            rk = role_by_name.get(core.norm_name(core.norm_header(h)))
            if rk:
                mapping[str(i)] = rk
    return ok({
        "upload_id": uid,
        "filename": name,
        "sheet": used,
        "sheets": sheets,
        "headers": headers,
        "row_count": len(body),
        "rows": body,
        "preview": body[:5],
        "mapping": mapping,
    })


def _cols_for(payload):
    up = UPLOADS.get(payload.get("upload_id") or "")
    if not up:
        raise ValueError("请先上传表格")
    meta = METAS.get(meta_key(CFG["project_key"], CFG["work_item_type"]))
    if not meta:
        raise ValueError("请先加载字段")
    items, _missing = form_items(meta, CFG["project_key"], CFG["work_item_type"])
    cols, problems = build_cols(meta, payload.get("mapping") or {}, up["headers"], items)
    return up, meta, cols, problems


def api_validate(payload):
    up, meta, cols, problems = _cols_for(payload)
    if not cols:
        return err("还没有做列映射")
    users = core.UserResolver(CFG, HERE / ".cache_users.json")
    items, _m = form_items(meta, CFG["project_key"], CFG["work_item_type"])
    profiles = (get_profiles(meta, CFG["project_key"], CFG["work_item_type"], items)
                .get("profiles") or {})
    type_col = next((c for c in cols.values() if c["field_key"] == "field_8e9f64"), None)
    mapped_keys = {c["field_key"] for c in cols.values() if c.get("kind") == "field"}
    rels_in = payload.get("relations") or {}
    rows = _apply_relations(up["rows"], cols, rels_in)
    applied_relations = sum(len(m or {}) for m in rels_in.values())
    issues, valid = [], 0
    for r, row in enumerate(rows):
        row_err = []
        required_keys = None
        if profiles and type_col:
            cell = row[type_col["index"]] if type_col["index"] < len(row) else None
            raw = core.cell_str(cell)
            oid = type_col["options"].get(core.norm_name(raw), raw)
            prof = profiles.get(oid)
            if prof:
                required_keys = set(prof.get("required") or []) | {"name"}
                miss = required_keys - mapped_keys
                if miss:
                    row_err.append({"col": "模板", "msg": "缺少该类型的必填列：" +
                                    "、".join(sorted(miss))})
        for h, c in cols.items():
            if c.get("kind") == "field" and not c.get("writable", True):
                continue  # 附件等不可导入列，跳过
            cell = row[c["index"]] if c["index"] < len(row) else None
            good, val, e = core.convert_value(cell, c, users)
            is_req = c.get("required")
            if required_keys is not None and c.get("kind") == "field":
                is_req = c["field_key"] in required_keys
            if not good:
                row_err.append({"col": h, "msg": e})
            elif is_req and val in (None, "") and not c.get("default"):
                row_err.append({"col": h, "msg": "必填为空"})
        if row_err:
            issues.append({"row": r + 2, "title": row[0] if row else "", "errors": row_err})
        else:
            valid += 1
    return ok({"issues": issues[:300], "issue_rows": len(issues), "valid_rows": valid,
               "mapped_columns": len(cols), "problems": problems,
               "applied_relations": applied_relations})


def api_run(payload):
    up, meta, cols, problems = _cols_for(payload)
    if not cols:
        return err("还没有做列映射")
    if problems:
        return err("；".join(problems))
    # 网页「关联字段」里选的值要覆盖到行数据，否则执行时还是空
    rels = payload.get("relations") or {}
    applied = 0
    if rels:
        up = {**up, "rows": _apply_relations(up["rows"], cols, rels)}
        applied = sum(len(m or {}) for m in rels.values())
    job_id = secrets.token_hex(4)
    job = {
        "id": job_id,
        "status": "running",
        "total": len(up["rows"]),
        "done": 0, "ok": 0, "failed": 0, "planned": 0,
        "dry_run": bool(payload.get("dry_run", True)),
        "started": time.time(),
        "logs": [],
        "results": [],
        "stop": False,
        "current": "",
        "applied_relations": applied,
    }
    if applied:
        job["logs"].append(f"已应用 {applied} 个网页关联字段选择")
    # 问题管理类型（模板）是创单必填项：界面没选就按布局默认值兜底
    tpl_name = (payload.get("template_name") or "").strip()
    items, _m = form_items(meta, CFG["project_key"], CFG["work_item_type"])
    tpl_item = next((it for it in items if it["field_key"] == "template"), None)
    if not tpl_name and tpl_item:
        opts = tpl_item.get("options") or []
        tpl_name = tpl_item.get("default") or (opts[0]["label"] if opts else "")
    if not tpl_name:
        tpl_field = all_fields(meta).get("template") or {}
        if tpl_field.get("options"):
            tpl_name = tpl_field["options"][0]["label"]
    if not tpl_name and (tpl_item or all_fields(meta).get("template")):
        return err("请选择「流程/模板」：问题管理类型是创建 BUG 的必填项")
    CFG["template_name"] = tpl_name

    JOBS[job_id] = job
    options = {
        "delay": float(payload.get("delay") or 0.4),
        "dry_run": job["dry_run"],
        "template_name": tpl_name,
        "limit": int(payload.get("limit") or 0),
    }
    t = threading.Thread(target=run_job, args=(job, up, meta, cols, options), daemon=True)
    t.start()
    return ok({"job_id": job_id})


def run_job(job, up, meta, cols, options):
    cfg = dict(CFG, template_name=options["template_name"] or CFG.get("template_name") or "")
    users = core.UserResolver(cfg, HERE / ".cache_users.json")
    payload_cols = {h: c for h, c in cols.items()
                    if c.get("kind") != "role"
                    and (c.get("kind") != "field" or c.get("writable", True))}
    role_cols = {h: c for h, c in cols.items() if c.get("kind") == "role"}
    skipped = [h for h, c in cols.items()
               if c.get("kind") == "field" and not c.get("writable", True)]
    if skipped:
        job["logs"].append("以下列不可导入，已忽略：" + "、".join(skipped))
    items, _m = form_items(meta, CFG["project_key"], CFG["work_item_type"])
    profiles = (get_profiles(meta, CFG["project_key"], CFG["work_item_type"], items)
                .get("profiles") or {})
    type_col = next((c for c in cols.values() if c["field_key"] == "field_8e9f64"), None)
    rows = up["rows"]
    if options["limit"]:
        rows = rows[: options["limit"]]
        job["total"] = len(rows)
    try:
        for i, row in enumerate(rows):
            if job["stop"]:
                job["status"] = "stopped"
                break
            title = row[0] if row else ""
            job["current"] = str(title)[:60]

            # 先按该行「测试单类型」确定必填集合，再构建 payload，
            # 否则 build_payload 会按静态 required 误报（如升级单要求关联项目）
            tprof = None
            if profiles and type_col:
                tcell = row[type_col["index"]] if type_col["index"] < len(row) else None
                traw = core.cell_str(tcell)
                toid = type_col["options"].get(core.norm_name(traw), traw)
                tprof = profiles.get(toid)
            row_cols = payload_cols
            if tprof:
                req = set(tprof.get("required") or []) | {"name"}
                row_cols = {
                    h: ({**c, "required": (c["field_key"] in req
                                           if c.get("kind") == "field" else c.get("required"))})
                    for h, c in payload_cols.items()
                }
            payload, issues = core.build_payload(row, row_cols, meta, cfg, users)

            role_values = []
            for h, c in role_cols.items():
                cell = row[c["index"]] if c["index"] < len(row) else None
                okv, val, err = core.convert_value(cell, {**c, "mode": "multi_user"}, users)
                if not okv:
                    issues.append(f"{h}: {err}")
                elif val:
                    try:
                        keys = json.loads(val)
                    except (TypeError, ValueError):
                        keys = [val]
                    if isinstance(keys, str):
                        keys = [keys]
                    keys = list(dict.fromkeys(keys))
                    role_values.append((c, keys))
                elif c.get("required"):
                    issues.append(f"{h}: 必填为空")

            # 必填角色必须在创建时就带上 role_owners，否则 create 直接报
            # ErrFieldRequired: 解决人(role_<pk>_<type>_role_xxx) 必填
            if role_values:
                owners = [{"role": c.get("api_role_key") or c.get("role_id"),
                           "owners": list(keys)} for c, keys in role_values]
                payload.setdefault("fields", []).append({
                    "field_key": "role_owners",
                    "field_value": json.dumps(owners, ensure_ascii=False),
                })

            # 创建接口要求的必填字段（meta-create-fields）必须在 payload 里
            req_keys = {k for k, f in core.extract_field_defs(meta.get("create_fields")).items()
                        if f.get("required")}
            if tprof:
                req_keys = set(tprof.get("required") or []) | {"name"}
            have = {f["field_key"] for f in payload.get("fields", [])}
            missing_keys = req_keys - have
            if missing_keys:
                issues.append("缺少创建必填字段：" + "、".join(sorted(missing_keys)))
            if all_fields(meta).get("template") and "template" not in have:
                issues.append("缺少流程模板 template")

            if issues:
                job["failed"] += 1
                job["results"].append({"row": i + 2, "title": title, "status": "失败",
                                       "id": "", "error": "；".join(issues[:3])})
                job["logs"].append(f"行{i + 2} 失败：{'；'.join(issues[:2])}")
                job["done"] += 1
                continue
            if options["dry_run"]:
                job["planned"] += 1
                job["done"] += 1
                job["results"].append({"row": i + 2, "title": title, "status": "预演",
                                       "id": "", "error": ""})
                if len(job["logs"]) < 300:
                    job["logs"].append(f"行{i + 2} 预演：{title}")
                continue

            params = json.dumps(payload, ensure_ascii=False)
            wid, last_err, attempt = None, "", 0
            while attempt <= int(cfg.get("max_retries") or 4):
                attempt += 1
                try:
                    resp = core.run_cli(cfg, ["workitem", "create",
                                              "--project-key", cfg["project_key"],
                                              "--work-item-type", cfg["work_item_type"],
                                              "--params", params], timeout=180)
                    wid = core.find_work_item_id(resp)
                    break
                except core.CliError as e:
                    last_err = str(e)
                    if any(s in last_err for s in ("429", "10429", "频繁", "超时", "timeout", "EOF")):
                        time.sleep(min(2 ** attempt, 20))
                        continue
                    break
            if wid:
                role_err = ""
                for c, keys in role_values:
                    try:
                        core.run_cli(cfg, [
                            "workitem", "update", "--work-item-id", wid,
                            "--project-key", cfg["project_key"],
                            "--role-operate", json.dumps(
                                [{"op": "add",
                                  "role_key": c.get("api_role_key") or c.get("role_id"),
                                  "user_keys": list(keys)}],
                                ensure_ascii=False),
                        ], timeout=120)
                    except core.CliError as e:
                        role_err = f"角色 {c.get('field_name')} 写入失败：{str(e)[:160]}"
                job["ok"] += 1
                job["results"].append({"row": i + 2, "title": title, "status": "成功",
                                       "id": wid, "error": role_err})
                job["logs"].append(f"行{i + 2} 成功：{wid} {str(title)[:30]}"
                                   + (f"（{role_err}）" if role_err else ""))
            else:
                job["failed"] += 1
                job["results"].append({"row": i + 2, "title": title, "status": "失败",
                                       "id": "", "error": last_err})
                job["logs"].append(f"行{i + 2} 失败：{last_err[:1500]}")
                try:  # 完整请求 + 完整错误落盘，便于排查
                    logdir = ROOT / "logs"
                    logdir.mkdir(exist_ok=True)
                    (logdir / "last_create_error.json").write_text(json.dumps({
                        "time": time.strftime("%Y-%m-%d %H:%M:%S"),
                        "excel_row": i + 2, "title": title,
                        "payload": payload, "error": last_err,
                    }, ensure_ascii=False, indent=2), encoding="utf-8")
                    job["error_log"] = str(logdir / "last_create_error.json")
                    job["logs"].append(f"完整请求/报错已写入 {job['error_log']}")
                except Exception:  # noqa: BLE001
                    pass
            job["done"] += 1
            if len(job["logs"]) > 300:
                del job["logs"][:100]
            time.sleep(max(0.0, options["delay"]))
        else:
            job["status"] = "done"
    except Exception as e:  # noqa: BLE001
        job["status"] = "error"
        job["error"] = f"{e}"
        job["logs"].append("异常：" + traceback.format_exc(limit=1)[:400])
    finally:
        job["elapsed"] = round(time.time() - job["started"], 1)
        job["current"] = ""


def api_job(payload):
    job = JOBS.get(payload.get("id") or "")
    if not job:
        return err("任务不存在")
    return ok({k: v for k, v in job.items() if k != "stop"})


def api_stop(payload):
    job = JOBS.get(payload.get("id") or "")
    if job:
        job["stop"] = True
    return ok({"stopped": bool(job)})


def api_results(payload):
    job = JOBS.get(payload.get("id") or "")
    if not job:
        return err("任务不存在")
    wb = Workbook()
    ws = wb.active
    ws.title = "结果"
    headers = ["行号", "标题", "状态", "工作项ID", "错误信息"]
    ws.append(headers)
    for i, c in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=i)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E79")
    for r in job["results"]:
        ws.append([r["row"], r["title"], r["status"], r["id"], r["error"]])
    for i, w in enumerate([8, 50, 10, 14, 60], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return {"__file__": buf.getvalue(), "filename": "批量建单结果.xlsx",
            "content_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}


ROUTES = {
    "/api/status": api_status,
    "/api/login/start": api_login_start,
    "/api/login/poll": api_login_poll,
    "/api/projects": api_projects,
    "/api/types": api_types,
    "/api/meta": api_meta,
    "/api/template": api_template,
    "/api/layout": api_layout,
    "/api/relation_options": api_relation_options,
    "/api/upload": api_upload,
    "/api/validate": api_validate,
    "/api/run": api_run,
    "/api/job": api_job,
    "/api/stop": api_stop,
    "/api/results": api_results,
}


# ----------------------------------------------------------------- HTTP


class Handler(BaseHTTPRequestHandler):
    server_version = "MeegleBulk/1.0"

    def log_message(self, fmt, *args):  # 安静一点
        if "/api/" in (args[0] if args else ""):
            return

    def _token_ok(self, qs) -> bool:
        tok = self.headers.get("X-Token") or (qs.get("t", [""])[0] if qs else "")
        return tok == TOKEN

    def _send_json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, data: bytes, filename: str, content_type: str):
        # HTTP 头只能是 latin-1，中文文件名必须按 RFC 5987 做百分号编码
        quoted = quote(filename, safe="")
        suffix = Path(filename).suffix or ""
        fallback = "download" + suffix
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header(
            "Content-Disposition",
            f"attachment; filename=\"{fallback}\"; filename*=UTF-8''{quoted}",
        )
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        if parsed.path == "/":
            html = (STATIC / "index.html").read_text(encoding="utf-8").replace("%%TOKEN%%", TOKEN)
            body = html.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if parsed.path.startswith("/static/"):
            f = STATIC / parsed.path[len("/static/"):]
            if f.exists() and f.is_file():
                data = f.read_bytes()
                self.send_response(200)
                self.send_header("Content-Type", mimetypes.guess_type(str(f))[0] or "application/octet-stream")
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(data)
                return
            self.send_error(404)
            return
        if parsed.path in ("/api/status", "/api/job", "/api/projects"):
            if not self._token_ok(qs):
                self._send_json(err("token 无效"), 403)
                return
            payload = {k: v[0] for k, v in qs.items()}
            try:
                self._send_json(ROUTES[parsed.path](payload))
            except Exception as e:  # noqa: BLE001
                self._send_json(err(f"{e}"), 500)
            return
        self.send_error(404)

    def do_POST(self):
        parsed = urlparse(self.path)
        qs = parse_qs(parsed.query)
        if parsed.path not in ROUTES:
            self.send_error(404)
            return
        if not self._token_ok(qs):
            self._send_json(err("token 无效"), 403)
            return
        try:
            length = int(self.headers.get("Content-Length") or 0)
            payload = json.loads(self.rfile.read(length) or b"{}")
            result = ROUTES[parsed.path](payload)
            if isinstance(result, dict) and "__file__" in result:
                fn = result["filename"].replace(" ", "_")
                self._send_file(result["__file__"], fn, result["content_type"])
            else:
                self._send_json(result)
        except core.CliError as e:
            self._send_json(err(str(e)))
        except Exception as e:  # noqa: BLE001
            self._send_json(err(f"{e}"))


def free_port(start=8765) -> int:
    for p in range(start, start + 30):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", p))
                return p
            except OSError:
                continue
    raise RuntimeError("找不到可用端口")


def port_in_use(p: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.3)
        return s.connect_ex(("127.0.0.1", p)) == 0


def main():
    if port_in_use(8765):
        print("!" * 62)
        print("  检测到 8765 端口已被占用：很可能有一个旧窗口还在跑旧代码。")
        print("  请先执行 stop_gui.bat（或关掉旧黑窗口），再启动本程序。")
        print("  本次会换一个端口启动，但旧页面不会自动更新。")
        print("!" * 62, flush=True)
    port = free_port()
    url = f"http://127.0.0.1:{port}/"
    srv = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print("=" * 62)
    print("  飞书项目 · 批量建单器（本地版）")
    print(f"  地址：{url}")
    print("  关闭：Ctrl+C；数据只在本机，登录态=你的 meegle 账号")
    print("=" * 62, flush=True)
    if "--no-browser" not in sys.argv:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")


if __name__ == "__main__":
    main()
