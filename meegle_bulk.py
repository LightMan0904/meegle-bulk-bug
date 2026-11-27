#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
meegle_bulk.py —— 批量创建飞书项目(BUG/缺陷)工作项

原理：飞书项目官方 CLI（@lark-project/meegle）用「你本人的 OAuth 身份」调用 OpenAPI，
      权限边界 = 你在该空间里的权限，不需要空间管理员给你开「Excel 导入」权限。

用法（在 config.json 所在目录执行）：
    python meegle_bulk.py check        # 1. 检查登录态
    python meegle_bulk.py inspect      # 2. 拉取字段元数据（生成 meta.json）
    python meegle_bulk.py template     # 3. 按你们空间真实字段生成 Excel 模板
    python meegle_bulk.py validate     # 4. 只校验，不写数据
    python meegle_bulk.py create       # 5. 预演（默认不落库）
    python meegle_bulk.py create --yes # 6. 真正创建（串行 + 断点续传 + 回写结果）

安全设计：默认 dry-run；只有显式 --yes 才会真的创建；每成功一条就落盘 state.json。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import re
import shutil
import subprocess
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.comments import Comment
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:  # pragma: no cover
    Workbook = None
    load_workbook = None

TZ = timezone(timedelta(hours=8))
HERE = Path(__file__).resolve().parent

# 中文 Windows 控制台可能是 GBK，遇到不支持的字符不要让程序崩掉
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors="replace")
    except Exception:  # noqa: BLE001
        pass

DEFAULT_CONFIG = {
    "project_key": "",
    "work_item_type": "",
    "excel": "",
    "sheet": "",
    "header_row": 1,
    "template_name": "",
    "rate_limit_seconds": 0.4,
    "max_retries": 4,
    "link_template": "",
    "meegle_cmd": "",
}

# ---------------------------------------------------------------- 基础工具


def now() -> str:
    return datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")


def die(msg: str, code: int = 1):
    print(f"[错误] {msg}", file=sys.stderr)
    sys.exit(code)


def info(msg: str):
    print(msg, flush=True)


def read_json(path: Path, default=None):
    if not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, data):
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def load_config(path: Path) -> dict:
    cfg = dict(DEFAULT_CONFIG)
    if path.exists():
        cfg.update(read_json(path, {}))
    # CLI 参数 / 环境变量可覆盖
    if os.environ.get("MEEGLE_PROJECT_KEY"):
        cfg["project_key"] = os.environ["MEEGLE_PROJECT_KEY"]
    if os.environ.get("MEEGLE_WORK_ITEM_TYPE"):
        cfg["work_item_type"] = os.environ["MEEGLE_WORK_ITEM_TYPE"]
    return cfg


# ---------------------------------------------------------------- CLI 封装


class CliError(RuntimeError):
    def __init__(self, msg: str, stdout: str = "", stderr: str = ""):
        super().__init__(msg)
        self.stdout = stdout or ""
        self.stderr = stderr or ""


def meegle_base(cfg: dict) -> list[str]:
    if cfg.get("meegle_cmd"):
        return [cfg["meegle_cmd"]]
    exe = shutil.which("meegle")
    if exe:
        return [exe]
    npx = shutil.which("npx.cmd") or shutil.which("npx")
    if npx:
        return [npx, "-y", "@lark-project/meegle@latest"]
    die("找不到 meegle CLI，请先安装：npm i -g @lark-project/meegle")


def parse_json_output(text: str):
    text = (text or "").strip()
    if not text:
        return None
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    start, end = text.find("{"), text.rfind("}")
    if start >= 0 and end > start:
        try:
            return json.loads(text[start : end + 1])
        except json.JSONDecodeError:
            pass
    return None


def run_cli(cfg: dict, args: list[str], timeout: int = 120, check: bool = True) -> dict:
    env = dict(os.environ)
    env.setdefault("MEEGLE_NO_UPDATE_CHECK", "1")
    cmd = meegle_base(cfg) + args + ["--format", "json"]
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            env=env,
        )
    except FileNotFoundError:
        raise CliError("无法启动 meegle CLI，请确认 Node.js 已安装")
    except subprocess.TimeoutExpired:
        raise CliError(f"命令超时（{timeout}s）：{' '.join(args[:2])}")

    out, err = proc.stdout or "", proc.stderr or ""
    if proc.returncode != 0:
        if check:
            raise CliError(f"meegle {' '.join(args[:2])} 失败：{(err or out).strip()[:800]}", out, err)
        return {"__error__": (err or out).strip()}
    data = parse_json_output(out)
    if data is None:
        raise CliError(f"无法解析 meegle 输出（前 400 字）：{out[:400]}", out, err)
    return data


# ---------------------------------------------------------------- 元数据解析

NAME_KEYS = ("field_name", "name", "label", "field_alias", "title")
KEY_KEYS = ("field_key", "key", "field_alias")
TYPE_KEYS = ("field_type_key", "field_type", "type", "field_type_id")
REQ_KEYS = ("is_required", "required", "required_flag")
ORDER_KEYS = ("order", "sort", "sort_order", "index", "position", "seq")
OPT_KEYS = ("option", "options", "option_list", "field_options", "option_items")


def walk(obj):
    if isinstance(obj, dict):
        yield obj
        for v in obj.values():
            yield from walk(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from walk(v)


def pick(d: dict, keys):
    for k in keys:
        if k in d and d[k] not in (None, "", []):
            return d[k]
    return None


def normalize_option(opt: dict) -> dict:
    label = pick(opt, ("label", "option_name", "name", "text", "value"))
    value = pick(opt, ("value", "option_id", "id", "key"))
    return {"label": str(label if label is not None else value),
            "value": str(value if value is not None else label)}


def flatten_options(raw_opts, leaves_only: bool = True) -> list[dict]:
    """把树形选项（tree_select）拍平成叶子列表；普通 select 原样返回。"""
    out: list[dict] = []

    def walk(node, acc):
        if not isinstance(node, dict):
            return
        disabled = node.get("disabled") or node.get("is_disabled") or node.get("disable")
        if disabled in (True, 1, "1", "true", "True"):
            return
        item = normalize_option(node)
        # 已经是规范化结构（带 path）时直接沿用，避免二次拍平丢层级
        if isinstance(node.get("path"), list) and node["path"]:
            cur = node["path"]
        else:
            cur = acc + [{"value": item["value"], "label": item["label"]}]
        kids = node.get("children") or node.get("child_options")
        if isinstance(kids, list) and kids:
            for k in kids:
                walk(k, cur)
            if not leaves_only:
                item["path"] = cur
                out.append(item)
        else:
            item["path"] = cur
            out.append(item)

    for o in (raw_opts or []):
        walk(o, [])
    seen, uniq = set(), []
    for o in out:
        if o["value"] in seen:
            continue
        seen.add(o["value"])
        uniq.append(o)
    return uniq


def extract_field_defs(resp) -> dict[str, dict]:
    """从任意形态的返回里抽出 字段key -> {name,type,required,options}"""
    found: dict[str, dict] = {}
    for d in walk(resp):
        key = pick(d, KEY_KEYS)
        name = pick(d, NAME_KEYS)
        if not key or not name:
            continue
        key, name = str(key), str(name)
        # 过滤掉形如 {"key": "..."} 的杂项
        if key in found and found[key].get("options"):
            continue
        ftype = pick(d, TYPE_KEYS)
        required = bool(pick(d, REQ_KEYS))
        raw_opts = pick(d, OPT_KEYS)
        opts = flatten_options(raw_opts) if isinstance(raw_opts, list) else []
        name = str(name).strip()
        raw_order = pick(d, ORDER_KEYS)
        try:
            order = int(raw_order) if raw_order is not None else None
        except (TypeError, ValueError):
            order = None
        related = (d.get("related_work_item_info") or d.get("related_work_item")
                   or d.get("related") or [])
        found[key] = {
            "field_key": key,
            "field_name": name,
            "field_type": str(ftype) if ftype is not None else "",
            "required": required,
            "options": opts,
            "order": order,
            "related": related if isinstance(related, list) else [],
        }
    return found


def merge_defs(*sources) -> dict:
    """合并多份字段定义：先出现的为准，但空字段（options/required/type）可被后面的补齐。"""
    out: dict[str, dict] = {}
    for src in sources:
        for k, v in (src or {}).items():
            cur = out.get(k)
            if not cur:
                out[k] = dict(v)
                continue
            for kk, vv in v.items():
                if kk not in cur or (not cur.get(kk) and vv not in (None, "", [], {}, False)):
                    cur[kk] = vv
    return out


def create_field_defs(meta: dict) -> dict:
    """只取「创建表单」里能填的字段；类型/选项用全量字段补齐。

    飞书项目的 meta-fields 会返回该工作项类型配置的**全部字段**
    （包括只在详情页、流程节点里出现的字段），而创建表单只暴露一部分。
    做模板/映射应该只用 create_fields。
    """
    created = extract_field_defs(meta.get("create_fields"))
    allfs = extract_field_defs(meta.get("fields"))
    if not created:
        return allfs
    return merge_defs(created, {k: v for k, v in allfs.items() if k in created})


NON_IMPORTABLE_TYPES = {
    "file", "multi-file", "attachment", "multi-attachment",
    "vote-boolean", "vote-option", "vote-option-multi", "computed", "formula",
}


def importable_fields(meta: dict) -> dict:
    """创建表单里能通过 API 写进去的字段（去掉附件、投票、计算字段）。"""
    return {
        k: f for k, f in create_field_defs(meta).items()
        if canon_type(f.get("field_type")) not in NON_IMPORTABLE_TYPES
    }


def order_fields(fields: dict, explicit: list | None = None) -> list[dict]:
    """按「创建表单」的顺序排列字段。

    优先用 explicit（界面上手工调整的顺序 / config.field_order），
    其次用接口自带的 order/sort 字段，最后保持接口返回顺序。
    """
    items = list(fields.values())
    if explicit:
        rank: dict[str, int] = {}
        for i, x in enumerate(explicit):
            rank[norm_name(x)] = i
        fallback = len(explicit) + 1000
        indexed = list(enumerate(items))
        indexed.sort(key=lambda pair: (
            rank.get(norm_name(pair[1]["field_key"]),
                     rank.get(norm_name(pair[1]["field_name"]), fallback)),
            pair[0],
        ))
        return [f for _, f in indexed]
    if any(isinstance(f.get("order"), int) for f in items):
        indexed = list(enumerate(items))
        indexed.sort(key=lambda pair: (
            0 if isinstance(pair[1].get("order"), int) else 1,
            pair[1].get("order") or 0,
            pair[0],
        ))
        return [f for _, f in indexed]
    return items


def extract_types(resp) -> list[dict]:
    out = []
    for d in walk(resp):
        key = pick(d, ("type_key", "work_item_type_key", "api_name"))
        name = pick(d, ("type_name", "name", "display_name"))
        if key and name and isinstance(key, str):
            out.append({"type_key": key, "type_name": str(name)})
    seen, uniq = set(), []
    for t in out:
        if t["type_key"] not in seen:
            seen.add(t["type_key"])
            uniq.append(t)
    return uniq


def fetch_meta(cfg: dict) -> dict:
    pk, wt = cfg["project_key"], cfg["work_item_type"]
    if not pk or not wt:
        die("config.json 里先填 project_key 和 work_item_type")
    info(f"→ 拉取空间 {pk} 的工作项类型…")
    types_resp = run_cli(cfg, ["workitem", "meta-types", "--project-key", pk])
    types = extract_types(types_resp)

    info(f"→ 拉取 {wt} 的创建字段…")
    create_resp = run_cli(cfg, ["workitem", "meta-create-fields", "--project-key", pk,
                                "--work-item-type", wt])
    create_fields = extract_field_defs(create_resp)

    info(f"→ 分页拉取 {wt} 的全部字段（每页 50）…")
    fields: dict[str, dict] = {}
    for page in range(1, 41):
        resp = run_cli(cfg, ["workitem", "meta-fields", "--project-key", pk,
                             "--work-item-type", wt, "--page-num", str(page)])
        page_fields = extract_field_defs(resp)
        new = {k: v for k, v in page_fields.items() if k not in fields}
        fields.update(new)
        if len(page_fields) < 50 or not new:
            break

    info("→ 拉取角色定义…")
    roles = []
    try:
        for page in range(1, 11):
            resp = run_cli(cfg, ["workitem", "meta-roles", "--project-key", pk,
                                 "--work-item-type", wt, "--page-num", str(page)])
            got = []
            for d in walk(resp):
                rid = pick(d, ("role_id", "role_key", "key"))
                rname = pick(d, ("role_name", "name"))
                if rid and rname:
                    got.append({"role_id": str(rid), "role_name": str(rname)})
            roles.extend(got)
            if len(got) < 50:
                break
    except CliError as e:
        info(f"  （角色读取失败，可忽略：{e}）")

    fields = merge_defs(create_fields, fields)
    meta = {
        "fetched_at": now(),
        "project_key": pk,
        "work_item_type": wt,
        "types": types,
        "create_fields": create_fields,
        "fields": fields,
        "roles": roles,
        "raw": {"meta_types": types_resp, "meta_create_fields": create_resp},
    }
    return meta


# ---------------------------------------------------------------- 映射


ALIASES = {
    "标题": ["name", "标题", "名称"],
    "名称": ["name", "标题"],
    "问题描述": ["description", "问题描述", "描述", "复现步骤"],
    "描述": ["description", "描述", "问题描述"],
    "问题管理类型": ["template", "问题管理类型", "模板"],
    "模板": ["template", "模板", "流程类型"],
    "优先级": ["priority", "优先级"],
    "报告人": ["reporter", "报告人"],
    "经办人": ["operator", "经办人", "处理人"],
    "所属模块": ["module", "所属模块", "模块"],
    "所属产品线": ["product_line", "所属产品线", "产品线"],
    "机型": ["machine_type", "机型", "设备型号"],
    "测试单类型": ["test_order_type", "测试单类型"],
}

def canon_type(t) -> str:
    """飞书项目里同一类型有时写 tree_select、有时写 tree-select，统一成短横线形式。"""
    return str(t or "").strip().lower().replace("_", "-").replace(" ", "-")


OPTION_TYPES = {"select", "single-select", "option", "single-option", "radio",
                "tree-select", "-work-item-template"}
MULTI_OPTION_TYPES = {"multi-select", "multi-option", "checkbox", "multi-text-select",
                      "tree-multi-select"}
USER_TYPES = {"user", "single-user", "member"}
MULTI_USER_TYPES = {"multi-user", "multi-member", "members"}
DATE_TYPES = {"date", "datetime", "precise-date", "schedule"}


def norm_header(h) -> str:
    s = str(h or "").strip()
    s = s.replace("\u3000", " ").replace("*", "").replace("＊", "")
    s = re.sub(r"[（(]\s*必填\s*[)）]", "", s)
    s = s.strip().rstrip(":：")
    return s.strip()


def norm_name(s: str) -> str:
    return re.sub(r"[\s_\-（）()·、,，。:：]*", "", str(s or "")).lower()


def guess_field(header: str, fields: dict) -> dict | None:
    h = norm_header(header)
    hn = norm_name(h)
    if not hn:
        return None
    # 1) 完全同名
    for f in fields.values():
        if norm_name(f["field_name"]) == hn:
            return f
    # 2) 别名
    for cand in ALIASES.get(h, []):
        cn = norm_name(cand)
        for f in fields.values():
            if norm_name(f["field_key"]) == cn or norm_name(f["field_name"]) == cn:
                return f
    # 3) 互相包含（长度>=2 才允许）
    if len(hn) >= 2:
        for f in fields.values():
            fn = norm_name(f["field_name"])
            if fn and (fn in hn or hn in fn) and abs(len(fn) - len(hn)) <= 4:
                return f
    return None


def field_mode(field: dict) -> str:
    if field.get("field_key") == "template":
        return "option"
    t = canon_type(field.get("field_type"))
    if t in OPTION_TYPES:
        return "option"
    if t in MULTI_OPTION_TYPES:
        return "multi_option"
    if t in USER_TYPES:
        return "user"
    if t in MULTI_USER_TYPES:
        return "multi_user"
    if t in DATE_TYPES:
        return "date"
    if t == "workitem-related-select":
        return "relation"
    if t == "workitem-related-multi-select":
        return "relation_multi"
    return "raw"


def build_mapping(headers: list[str], fields: dict, override: dict) -> dict:
    columns, unmapped = {}, []
    for idx, raw in enumerate(headers):
        h = norm_header(raw)
        if not h or h.startswith("#"):
            continue
        ov = override.get(raw) or override.get(h) or {}
        field = None
        if ov.get("field_key"):
            field = fields.get(ov["field_key"])
            if field is None:
                die(f"mapping.json 里 {h} 指定的 field_key={ov['field_key']} 在空间字段里不存在")
        else:
            field = guess_field(h, fields)
        if not field:
            unmapped.append(h)
            continue
        columns[h] = {
            "index": idx,
            "field_key": field["field_key"],
            "field_name": field["field_name"],
            "field_type": field.get("field_type", ""),
            "mode": ov.get("mode") or field_mode(field),
            "options": {norm_name(o["label"]): o["value"] for o in field.get("options", [])},
            "option_labels": {o["value"]: o["label"] for o in field.get("options", [])},
            "required": field.get("required", False),
        }
    return {"columns": columns, "unmapped": unmapped}


# ---------------------------------------------------------------- 值转换


def to_ms(value: str):
    s = str(value).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%Y/%m/%d %H:%M", "%Y/%m/%d"):
        try:
            dt = datetime.strptime(s, fmt).replace(tzinfo=TZ)
            return str(int(dt.timestamp() * 1000))
        except ValueError:
            continue
    return None


def split_multi(value: str) -> list[str]:
    return [x.strip() for x in re.split(r"[,，、;；/|\r\n]+", str(value)) if x.strip()]


class UserResolver:
    def __init__(self, cfg: dict, state_path: Path):
        self.cfg = cfg
        self.path = state_path
        self.cache = read_json(state_path, {}) or {}

    def resolve(self, name: str) -> str | None:
        key = norm_name(name)
        if not key:
            return None
        if key in self.cache:
            return self.cache[key]
        try:
            resp = run_cli(self.cfg, ["user", "search", "--user-keys", json.dumps([name], ensure_ascii=False),
                                      "--project-key", self.cfg["project_key"],
                                      "--need-all-status", "false"], timeout=60)
        except CliError:
            return None
        userkey = None
        for d in walk(resp):
            uk = pick(d, ("user_key", "userkey"))
            if uk:
                userkey = str(uk)
                break
        if userkey:
            self.cache[key] = userkey
            write_json(self.path, self.cache)
        return userkey


def convert_value(cell, col: dict, users: UserResolver | None):
    """返回 (ok, value_str, error)"""
    if cell is None:
        return True, None, None
    if isinstance(cell, datetime):
        raw = cell.strftime("%Y-%m-%d %H:%M") if (cell.hour or cell.minute) else cell.strftime("%Y-%m-%d")
    else:
        raw = str(cell).strip()
    if raw == "":
        return True, None, None

    mode = col["mode"]
    if mode == "option" or mode == "multi_option":
        if not col["options"]:
            return True, raw, None
        tree = bool(col.get("tree"))
        paths = col.get("paths") or {}

        def option_value(text: str):
            val = col["options"].get(norm_name(text))
            if val is None:
                return None, f"「{text}」不是 {col['field_name']} 的合法选项"
            if tree:
                # 级联字段必须传「整条路径」的 option 对象数组（根→叶），
                # 传裸 option_id 或单层对象都会报 ErrCascadeFieldValueDepthInvalid
                path = paths.get(norm_name(text)) or [{"value": val, "label": text}]
                return json.dumps([{"option_id": p["value"]} for p in path],
                                  ensure_ascii=False), None
            return val, None

        if mode == "option":
            v, err = option_value(raw)
            return (True, v, None) if err is None else (False, None, err)
        vals = []
        for part in split_multi(raw):
            v, err = option_value(part)
            if err:
                return False, None, err
            vals.append(json.loads(v) if tree else v)
        return True, json.dumps(vals, ensure_ascii=False), None

    if mode == "relation":
        return True, raw, None
    if mode == "relation_multi":
        ids = [x.strip() for x in split_multi(raw)]
        vals = [int(x) if x.isdigit() else x for x in ids]
        return True, json.dumps(vals, ensure_ascii=False), None

    if mode in ("user", "multi_user"):
        if users is None:
            return False, None, "人员字段需要登录态解析"
        names = [raw] if mode == "user" else split_multi(raw)
        keys = []
        for n in names:
            k = users.resolve(n)
            if not k:
                return False, None, f"找不到人员「{n}」"
            keys.append(k)
        return True, (keys[0] if mode == "user"
                      else json.dumps(keys, ensure_ascii=False)), None

    if mode == "date":
        ms = to_ms(raw)
        if ms is None:
            return False, None, f"日期格式无法识别：{raw}（要 YYYY-MM-DD）"
        return True, ms, None

    return True, raw, None


# ---------------------------------------------------------------- Excel


def cell_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def read_sheet(path: Path, sheet: str | None, header_row: int):
    if load_workbook is None:
        die("缺少 openpyxl，请先执行：python -m pip install openpyxl")
    if not path.exists():
        die(f"Excel 不存在：{path}")
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        die("Excel 是空的")
    headers = [cell_str(h) for h in rows[header_row - 1]]
    data = []
    for r in rows[header_row:]:
        if all(cell_str(c) == "" for c in r[: max(1, len(headers))]):
            continue
        data.append(list(r))
    return ws.title, [norm_header(h) for h in headers], headers, data


# ---------------------------------------------------------------- 命令


def cmd_check(cfg: dict, args):
    info("→ meegle auth status")
    data = run_cli(cfg, ["auth", "status"], check=False)
    info(json.dumps(data, ensure_ascii=False, indent=2))
    if data.get("__error__") or data.get("authenticated") is False:
        info("\n未登录或登录失效，请执行：\n  npx -y @lark-project/meegle@latest auth login --device-code --host project.feishu.cn\n")
        return 1
    info("\n→ meegle user me")
    me = run_cli(cfg, ["user", "me"], check=False)
    info(json.dumps(me, ensure_ascii=False, indent=2))
    return 0


def cmd_inspect(cfg: dict, args):
    meta = fetch_meta(cfg)
    out = Path(args.meta or (HERE / "meta.json"))
    write_json(out, meta)
    info(f"\n已保存 {out}")
    if meta["types"]:
        info("\n工作项类型：")
        for t in meta["types"]:
            info(f"  - {t['type_key']:<24} {t['type_name']}")
    fields = create_field_defs(meta)
    if fields:
        total = len(merge_defs(extract_field_defs(meta.get("fields")),
                               extract_field_defs(meta.get("create_fields"))))
        info(f"\n创建表单字段 {len(fields)} 个（该类型全部字段 {total} 个，* 为创建必填）：")
        for f in order_fields(fields):
            star = "*" if f.get("required") else " "
            opt = f"  可选值: {', '.join(o['label'] for o in f['options'][:6])}" if f["options"] else ""
            info(f"  {star} {f['field_key']:<28} {f['field_name']:<18} {f['field_type']:<18}{opt}")
    return 0


FIELD_WIDTHS = {
    "name": 30,
    "description": 55,
    "field_b7d8b9": 16,   # 订单号
    "field_96b3b7": 32,   # 出现bug程序
    "field_5b5335": 22,   # 评审意见
    "field_3cb243": 34,   # 备注
    "field_8db4f8": 14,   # 附件
}


def _col_width(item: dict) -> int:
    key = item.get("field_key", "")
    if key.startswith("role:"):
        return 12
    if key in FIELD_WIDTHS:
        return FIELD_WIDTHS[key]
    t = canon_type(item.get("field_type"))
    if t in OPTION_TYPES or t in MULTI_OPTION_TYPES:
        return 16
    if t.startswith("workitem-related"):
        return 20
    if t in NON_IMPORTABLE_TYPES:
        return 14
    return 18


def resolve_form_items(meta: dict, layout_items: list | None = None,
                       project_key: str | None = None, work_item_type: str | None = None):
    """把「页面布局」+ 空间元数据合成为有序的字段/角色列表。

    - 布局里记录页面顺序（含必填标记、说明、默认值）
    - 元数据里找字段/角色定义（名称、类型、选项）
    - 管理员新加、但布局里没有的创建字段，自动追加到末尾并标记 is_new
    - 布局里存在、元数据里已删掉的字段，进入 missing，不显示
    """
    allf = merge_defs(extract_field_defs(meta.get("fields")),
                      extract_field_defs(meta.get("create_fields")))
    created = extract_field_defs(meta.get("create_fields"))
    roles = {str(r.get("role_id")): r for r in (meta.get("roles") or []) if r.get("role_id")}

    items, seen, missing = [], set(), []
    for spec in (layout_items or []):
        kind = spec.get("kind", "field")
        key = str(spec.get("key") or "")
        if kind == "role":
            r = roles.get(key)
            if not r:
                missing.append(key)
                continue
            pk = project_key or meta.get("project_key") or ""
            wt = work_item_type or meta.get("work_item_type") or ""
            # role_owners / role_operate 都用 meta-roles 返回的裸 role_id（实测复合 key 会被拒）
            items.append({
                "kind": "role", "field_key": f"role:{key}", "role_id": key,
                "api_role_key": key,
                "api_role_key_full": f"role_{pk}_{wt}_{key}" if (pk and wt) else key,
                "field_name": spec.get("name") or r.get("role_name") or key,
                "field_type": "role", "required": bool(spec.get("required")),
                "options": [], "writable": True,
                "note": spec.get("note", ""), "default": spec.get("default", ""),
                "is_new": False,
            })
            seen.add(("role", key))
            continue

        f = allf.get(key)
        if not f:
            missing.append(key)
            continue
        cf = created.get(key)
        items.append({
            "kind": "field", "field_key": key, "field_name": f["field_name"],
            "field_type": f.get("field_type", ""),
            "required": bool(spec.get("required")) or bool((cf or {}).get("required")),
            "options": f.get("options", []),
            "related": f.get("related") or [],
            "writable": canon_type(f.get("field_type")) not in NON_IMPORTABLE_TYPES,
            "note": spec.get("note", ""), "default": spec.get("default", ""),
            "is_new": False,
        })
        seen.add(("field", key))

    # 管理员后加的创建字段：自动补在最后
    for k, f in created.items():
        if ("field", k) in seen:
            continue
        items.append({
            "kind": "field", "field_key": k, "field_name": f["field_name"],
            "field_type": f.get("field_type", ""),
            "required": bool(f.get("required")), "options": f.get("options", []),
            "related": f.get("related") or [],
            "writable": canon_type(f.get("field_type")) not in NON_IMPORTABLE_TYPES,
            "note": "管理员新增的创建字段", "default": "", "is_new": True,
        })
    return items, missing


def layout_items_from_names(meta: dict, names: list, required_names=None, notes=None):
    """按字段/角色名称生成布局项。

    用于首次运行、本地还没有 layouts/ 的情况：只依赖名称，不硬编码任何
    空间 ID、字段 ID、角色 ID；真正的 key 在运行时从 meta 里解析。
    """
    allf = merge_defs(extract_field_defs(meta.get("create_fields")),
                      extract_field_defs(meta.get("fields")))
    roles = meta.get("roles") or []
    req = set(required_names or [])
    note_map = notes or {}
    out = []
    for name in names or []:
        target = norm_name(name)
        f = next((v for v in allf.values()
                  if norm_name(v.get("field_name")) == target), None)
        if f:
            out.append({"kind": "field", "key": f["field_key"],
                        "name": f["field_name"], "required": name in req,
                        "note": note_map.get(name, ""), "default": ""})
            continue
        r = next((x for x in roles
                  if norm_name(x.get("role_name") or "") == target), None)
        if r:
            out.append({"kind": "role", "key": r["role_id"],
                        "name": r.get("role_name") or name, "required": name in req,
                        "note": note_map.get(name, ""), "default": ""})
    return out


def build_template_items(items: list[dict], meta: dict, sheet_title: str = "数据",
                         profiles: dict | None = None):
    """按给定顺序生成模板；每项包含 field_key/field_name/field_type/required/options/note/writable。"""
    if Workbook is None:
        raise RuntimeError("缺少 openpyxl，请先执行：python -m pip install openpyxl")
    if not items:
        raise ValueError("没有可用字段，请重新拉取元数据")

    wb = Workbook()
    ws = wb.active
    ws.title = (sheet_title or "数据")[:28]
    ws2 = wb.create_sheet("填写说明")
    ws_opts = wb.create_sheet("选项")
    ws_opts.sheet_state = "hidden"

    ws.append([(str(f["field_name"]).strip() + (" *" if f.get("required") else "")) for f in items])
    fill_req = PatternFill("solid", fgColor="C00000")
    fill_opt = PatternFill("solid", fgColor="1F4E79")
    fill_no = PatternFill("solid", fgColor="7A7F85")
    used_names: set[str] = set()
    opt_col = 0
    for i, f in enumerate(items, start=1):
        required = bool(f.get("required"))
        writable = f.get("writable", True)
        c = ws.cell(row=1, column=i)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = fill_req if required else (fill_opt if writable else fill_no)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        note = f.get("note") or ("必填：不填会创建失败" if required else "选填：可以留空")
        c.comment = Comment(note, "批量建单器")
        ws.column_dimensions[get_column_letter(i)].width = _col_width(f)

        if f.get("options") and writable:
            opt_col += 1
            letter = get_column_letter(opt_col)
            ws_opts.cell(row=1, column=opt_col, value=f["field_name"])
            opts = f["options"][:1000]
            for r, o in enumerate(opts, start=2):
                ws_opts.cell(row=r, column=opt_col, value=o["label"])
            base = "opt_" + re.sub(r"[^0-9A-Za-z_]", "_", f["field_key"])[:36]
            name, n = base, 1
            while name in used_names:
                n += 1
                name = f"{base}_{n}"
            used_names.add(name)
            ref = f"'{ws_opts.title}'!${letter}$2:${letter}${len(opts) + 1}"
            dv_kwargs = dict(
                type="list", allow_blank=True, showErrorMessage=True,
                errorTitle="选项不匹配",
                error="请从下拉里选择；文字必须和飞书项目里的选项完全一致",
            )
            try:
                wb.defined_names.add(DefinedName(name, attr_text=ref))
                dv = DataValidation(formula1=name, **dv_kwargs)
            except Exception:  # 名称管理器失败时退回内联列表
                labels = [o["label"] for o in opts if "," not in o["label"]][:30]
                dv = DataValidation(formula1='"' + ",".join(labels) + '"', **dv_kwargs)
            ws.add_data_validation(dv)
            dv.add(f"{get_column_letter(i)}2:{get_column_letter(i)}5000")
    ws.freeze_panes = "A2"

    # ---- 填写说明 ----
    ws2.append(["填写说明"])
    ws2["A1"].font = Font(bold=True, size=14)
    ws2.append(["表头顺序 = 飞书项目创建页顺序；红色表头 = 必填（*），蓝色 = 选填，灰色 = 不可通过 Excel 导入。"])
    ws2.append(["表头不要改名、不要增删列；不需要的列可以整列留空。"])
    ws2.append([])
    ws2.append(["序号", "字段", "必填", "类型", "可选值（下拉）", "说明"])
    for cell in ws2[ws2.max_row]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="E7E9EC")
    for idx, f in enumerate(items, start=1):
        vals = " / ".join(o["label"] for o in f.get("options", [])[:40])
        note = f.get("note") or ""
        ws2.append([idx, str(f["field_name"]).strip(),
                    "必填" if f.get("required") else "选填",
                    f.get("field_type", ""), vals, note])
    ws2.append([])

    if profiles:
        labels = [p.get("label") or k for k, p in profiles.items()]
        ws2.append(["按「测试单类型」区分必填"])
        ws2.append(["字段"] + labels)
        for cell in ws2[ws2.max_row]:
            cell.font = Font(bold=True)
        for f in items:
            row = [str(f["field_name"]).strip()]
            for p in profiles.values():
                if f.get("kind") == "role":
                    req = f.get("api_role_key_full") in (p.get("role_required") or [])
                else:
                    req = f["field_key"] in (p.get("required") or [])
                row.append("必填" if req else "选填")
            ws2.append(row)
        ws2.append([])

    notes = [
        "使用说明：",
        "1. 第一行是表头，从第二行开始每一行 = 一条工作项；不同行可以只填部分列。",
        "2. 选项类字段请用下拉值；手工输入时必须和选项文字完全一致（大小写、空格）。",
        "3. 人员字段（含解决人）填姓名或企业邮箱；多个解决人用「、」或逗号分隔；关联类字段可填工作项 ID，也可以留空，",
        "   上传后在网页「关联字段」里按行搜索选择（会自动转成 ID）。",
        "4. 附件列保留但无法随 Excel 导入，建单后单独上传；填了也不会生效。",
        "5. 字段/选项变了，点「刷新字段」重新拉取，再重新下载模板，不要复用旧文件。",
        "",
        "本次生成时间：" + now(),
        "空间：" + str(meta.get("project_key")),
        "工作项类型：" + str(meta.get("work_item_type")),
    ]
    for n in notes:
        ws2.append([n])
    ws2.column_dimensions["A"].width = 8
    ws2.column_dimensions["B"].width = 24
    ws2.column_dimensions["C"].width = 8
    ws2.column_dimensions["D"].width = 16
    ws2.column_dimensions["E"].width = 60
    ws2.column_dimensions["F"].width = 46
    for row in ws2.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    return wb, items


def build_template_wb(meta: dict, sheet_title: str = "数据", field_order: list | None = None):
    """CLI/兼容入口：只用创建表单里可写的字段生成模板。"""
    fields = importable_fields(meta)
    ordered = order_fields(fields, field_order)
    items = [{
        "kind": "field", "field_key": f["field_key"], "field_name": f["field_name"],
        "field_type": f.get("field_type", ""), "required": bool(f.get("required")),
        "options": f.get("options", []), "writable": True, "note": "", "default": "",
        "is_new": False,
    } for f in ordered]
    return build_template_items(items, meta, sheet_title)

def cmd_template(cfg: dict, args):
    meta = read_json(Path(getattr(args, "meta", None) or (HERE / "meta.json")))
    if not meta:
        die("还没有 meta.json，请先执行：python meegle_bulk.py inspect")
    order = [x.strip() for x in (getattr(args, "order", "") or "").split(",") if x.strip()]
    try:
        wb, cols = build_template_wb(meta, "数据", order or None)
    except (ValueError, RuntimeError) as e:
        die(str(e))
    out = Path(args.out or (HERE / "BUG批量导入_模板.xlsx"))
    wb.save(out)
    info(f"已生成模板：{out}\n列：{', '.join(c['field_name'] for c in cols)}")
    return 0


def load_pipeline(cfg: dict, args):
    meta = read_json(Path(args.meta) if args.meta else (HERE / "meta.json"))
    if not meta:
        die("还没有 meta.json，请先执行：python meegle_bulk.py inspect")
    excel = Path(args.excel or cfg["excel"] or (HERE / "BUG批量导入_模板.xlsx"))
    sheet, headers, raw_headers, rows = read_sheet(excel, cfg.get("sheet") or None,
                                                   int(cfg.get("header_row") or 1))
    # 兼容：meta.json 可能是 inspect 生成的规范化结构，也可能是手工改过的原始结构
    # 只把「创建表单且可写」的字段作为可映射候选
    all_fields = importable_fields(meta)

    override = read_json(Path(args.mapping) if args.mapping else (HERE / "mapping.json"), {}) or {}
    mapping = build_mapping(raw_headers, all_fields, override)
    return meta, excel, sheet, headers, raw_headers, rows, mapping, all_fields


def cmd_validate(cfg: dict, args):
    meta, excel, sheet, headers, raw_headers, rows, mapping, all_fields = load_pipeline(cfg, args)
    cols = mapping["columns"]
    info(f"Excel: {excel.name} / sheet={sheet} / {len(rows)} 行")
    info(f"已映射 {len(cols)} 列：")
    for h, c in cols.items():
        info(f"  {h:<16} → {c['field_key']:<24} ({c['field_type'] or c['mode']})")
    if mapping["unmapped"]:
        info(f"\n[忽略] 未匹配到字段的列：{', '.join(mapping['unmapped'])}")

    col_keys = {c["field_key"] for c in cols.values()}
    if cfg.get("template_name"):
        col_keys.add("template")  # 由 config 兜底填充
    missing_req = [f for f in all_fields.values()
                   if f.get("required") and f["field_key"] not in col_keys]
    global_error = bool(missing_req)
    if missing_req:
        info("\n[错误] 这些必填字段在 Excel 里没有对应列，创建会被拒绝：" +
             ", ".join(f"{f['field_name']}({f['field_key']})" for f in missing_req))

    users = UserResolver(cfg, HERE / "users.json")
    errors, warnings = 0, 0
    for i, row in enumerate(rows, start=2):
        row_err = []
        for h, c in cols.items():
            cell = row[c["index"]] if c["index"] < len(row) else None
            ok, val, err = convert_value(cell, c, users)
            if not ok:
                row_err.append(f"{h}: {err}")
            elif c.get("required") and val in (None, ""):
                row_err.append(f"{h}: 必填为空")
        if row_err:
            errors += 1
            if errors <= 20:
                info(f"  第 {i} 行 [X] " + "；".join(row_err[:3]))
    if errors > 20:
        info(f"  …… 其余 {errors - 20} 行同样有问题")
    if errors == 0 and not global_error:
        info(f"\n校验通过：{len(rows)} 行数据可以创建。")
    else:
        info(f"\n校验发现 {errors} 行有问题，先改 Excel。")
    return 1 if (errors or global_error) else 0


def fingerprint(row, cols) -> str:
    parts = []
    for key in ("name", "description"):
        for c in cols.values():
            if c["field_key"] == key:
                parts.append(cell_str(row[c["index"]]))
    return hashlib.sha1("|".join(parts).encode("utf-8")).hexdigest()[:12]


def build_payload(row, cols, meta, cfg, users=None) -> tuple[dict, list[str]]:
    fields, problems = [], []
    for h, c in cols.items():
        if c.get("kind") == "role":
            continue  # 角色走建单后的 role_operate，不进 fields
        if c.get("writable") is False:
            continue  # 附件等不可写列
        cell = row[c["index"]] if c["index"] < len(row) else None
        ok, val, err = convert_value(cell, c, users)
        if not ok:
            problems.append(f"{h}: {err}")
            continue
        if val in (None, "") and c.get("required") and c.get("default"):
            val = str(c["default"])
        if val in (None, ""):
            if c.get("required"):
                problems.append(f"{h}: 必填为空")
            continue
        fields.append({"field_key": c["field_key"], "field_value": str(val)})

    keys = {f["field_key"] for f in fields}
    # 模板字段（流程类型）必填，允许用 config.template_name 兜底
    want_template = cfg.get("template_name")
    if want_template and "template" not in keys:
        tpl_field = (meta.get("fields") or {}).get("template") or (meta.get("create_fields") or {}).get("template")
        if tpl_field and tpl_field.get("options"):
            for o in tpl_field["options"]:
                if norm_name(o["label"]) == norm_name(want_template):
                    fields.insert(0, {"field_key": "template", "field_value": o["value"]})
                    keys.add("template")
                    break
            else:
                problems.append(f"config.template_name「{want_template}」不在模板选项里")
    return {"fields": fields}, problems


def find_work_item_id(resp) -> str | None:
    for d in walk(resp):
        if isinstance(d.get("work_item_id"), (str, int)):
            return str(d["work_item_id"])
    for d in walk(resp):
        attr = d.get("work_item_attribute")
        if isinstance(attr, dict) and isinstance(attr.get("id"), (str, int)):
            return str(attr["id"])
    for k in ("id",):
        if isinstance(resp, dict) and isinstance(resp.get(k), (str, int)):
            return str(resp[k])
    return None


def cmd_create(cfg: dict, args):
    meta, excel, sheet, headers, raw_headers, rows, mapping, all_fields = load_pipeline(cfg, args)
    cols = mapping["columns"]
    if not cols:
        die("没有匹配到任何字段，检查表头或 mapping.json")

    users = UserResolver(cfg, HERE / "users.json")
    state_path = HERE / "state.json"
    state = read_json(state_path, {}) or {}
    results_path = HERE / "results.csv"

    start = int(args.start or 0)
    limit = int(args.limit or 0)
    targets = []
    for offset, row in enumerate(rows[start:], start=start + 2):
        if limit and len(targets) >= limit:
            break
        fp = f"{fingerprint(row, cols)}:{offset}"
        if state.get(fp, {}).get("status") == "ok" and not args.force:
            continue
        targets.append((offset, row, fp))

    info(f"待处理 {len(targets)} 行（已跳过成功过的 {len(state)} 条记录）")
    if not args.yes:
        info("当前是【预演模式】，不会真的创建。确认无误后加 --yes。\n")

    ok = fail = skip = 0
    created = []
    try:
        for n, (excel_row, row, fp) in enumerate(targets, start=1):
            payload, problems = build_payload(row, cols, meta, cfg, users)

            title = ""
            for c in cols.values():
                if c["field_key"] == "name":
                    title = cell_str(row[c["index"]])

            if problems:
                fail += 1
                info(f"[{n}/{len(targets)}] 行{excel_row} [X] {title[:30]} | {'；'.join(problems[:2])}")
                created.append({"excel_row": excel_row, "title": title, "status": "failed",
                                "id": "", "error": "；".join(problems), "time": now()})
                continue

            if not args.yes:
                info(f"[{n}/{len(targets)}] 行{excel_row} 预演 | {title[:40]}")
                if n <= 2 or args.show_payload:
                    info("    payload = " + json.dumps(payload, ensure_ascii=False)[:900])
                skip += 1
                continue

            params = json.dumps(payload, ensure_ascii=False)
            if len(params) > 20000:
                fail += 1
                info(f"[{n}/{len(targets)}] 行{excel_row} [X] {title[:30]} | 单条数据超过 20000 字符，请拆分问题描述")
                continue
            attempt, last_err, wid = 0, "", None
            while attempt <= int(cfg.get("max_retries") or 4):
                attempt += 1
                try:
                    resp = run_cli(cfg, ["workitem", "create",
                                         "--project-key", cfg["project_key"],
                                         "--work-item-type", cfg["work_item_type"],
                                         "--params", params], timeout=180)
                    wid = find_work_item_id(resp)
                    break
                except CliError as e:
                    last_err = str(e)
                    if any(s in last_err for s in ("429", "10429", "频繁", "超时", "timeout", "EOF")):
                        time.sleep(min(2 ** attempt, 30))
                        continue
                    break

            if wid:
                ok += 1
                state[fp] = {"status": "ok", "id": wid, "time": now(), "title": title}
                write_json(state_path, state)
                created.append({"excel_row": excel_row, "title": title, "status": "ok",
                                "id": wid, "error": "", "time": now()})
                link = ""
                if cfg.get("link_template"):
                    link = cfg["link_template"].format(project_key=cfg["project_key"],
                                                       work_item_type=cfg["work_item_type"], id=wid)
                info(f"[{n}/{len(targets)}] 行{excel_row} [OK] {wid} {title[:36]}{('  ' + link) if link else ''}")
            else:
                fail += 1
                state[fp] = {"status": "failed", "error": last_err[:400], "time": now(), "title": title}
                write_json(state_path, state)
                created.append({"excel_row": excel_row, "title": title, "status": "failed",
                                "id": "", "error": last_err[:400], "time": now()})
                info(f"[{n}/{len(targets)}] 行{excel_row} [X] {title[:30]} | {last_err[:160]}")
            time.sleep(float(cfg.get("rate_limit_seconds") or 0.4))
    except KeyboardInterrupt:
        info("\n已中断，进度已保存到 state.json，重新执行会跳过成功的行。")
    finally:
        if args.yes and created:
            writeback(excel, sheet, created)
            with open(results_path, "w", encoding="utf-8-sig", newline="") as f:
                import csv
                w = csv.DictWriter(f, fieldnames=["excel_row", "title", "status", "id", "error", "time"])
                w.writeheader()
                w.writerows(created)

    info(f"\n完成：成功 {ok}，失败 {fail}，预演 {skip}")
    if not args.yes:
        info("预演结束。确认无误后执行：python meegle_bulk.py create --yes")
    return 0 if fail == 0 else 2


def writeback(excel: Path, sheet: str, created: list[dict]):
    if load_workbook is None:
        return
    try:
        wb = load_workbook(excel)
        ws = wb[sheet] if sheet in wb.sheetnames else wb[wb.sheetnames[0]]
        base = ws.max_column
        extra = ["导入状态", "工作项ID", "错误信息", "导入时间"]
        for i, name in enumerate(extra, start=1):
            ws.cell(row=1, column=base + i, value=name).font = Font(bold=True)
        by_row = {c["excel_row"]: c for c in created}
        for r in range(2, ws.max_row + 1):
            c = by_row.get(r)
            if not c:
                continue
            ws.cell(row=r, column=base + 1, value=c["status"])
            ws.cell(row=r, column=base + 2, value=c["id"])
            ws.cell(row=r, column=base + 3, value=c["error"][:300])
            ws.cell(row=r, column=base + 4, value=c["time"])
        out = excel.with_name(excel.stem + "_已回写" + excel.suffix)
        wb.save(out)
        info(f"\n结果已回写到：{out}")
    except Exception as e:  # noqa: BLE001
        info(f"\n[警告] 回写 Excel 失败（文件可能正被占用）：{e}\n结果见 results.csv")


# ---------------------------------------------------------------- main


def main():
    parser = argparse.ArgumentParser(description="飞书项目批量创建 BUG（绕过 Excel 导入权限）")
    parser.add_argument("-c", "--config", default=str(HERE / "config.json"), help="配置文件")
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("check", help="检查登录态")
    p = sub.add_parser("inspect", help="拉取字段元数据")
    p.add_argument("--meta")
    p = sub.add_parser("template", help="生成 Excel 模板")
    p.add_argument("--out")
    p.add_argument("--meta")
    p.add_argument("--order", help="按这个顺序排字段，逗号分隔（字段名或 key）")
    p = sub.add_parser("validate", help="只校验 Excel")
    p.add_argument("--excel")
    p.add_argument("--meta")
    p.add_argument("--mapping")
    p = sub.add_parser("create", help="批量创建")
    p.add_argument("--excel")
    p.add_argument("--meta")
    p.add_argument("--mapping")
    p.add_argument("--start", type=int, default=0, help="从第几条数据开始（0 基）")
    p.add_argument("--limit", type=int, default=0, help="只处理前 N 条")
    p.add_argument("--yes", action="store_true", help="真的创建（默认只预演）")
    p.add_argument("--force", action="store_true", help="忽略 state.json 里的成功记录")
    p.add_argument("--show-payload", action="store_true")

    args = parser.parse_args()
    cfg = load_config(Path(args.config))
    if args.cmd != "check":
        if not cfg["project_key"] or not cfg["work_item_type"]:
            die("请先在 config.json 里填 project_key / work_item_type（可用 inspect 确认）")

    handler = {
        "check": cmd_check,
        "inspect": cmd_inspect,
        "template": cmd_template,
        "validate": cmd_validate,
        "create": cmd_create,
    }[args.cmd]
    sys.exit(handler(cfg, args))


if __name__ == "__main__":
    main()
