/* 飞书项目 · 批量建单器 —— 前端逻辑 */
const TOKEN = window.__TOKEN__;
const $ = (id) => document.getElementById(id);
const state = {
  projectKey: "", type: "", meta: null,
  upload: null, mapping: {}, autoMapping: {},
  jobId: "", timer: null, file: null,
  fieldOrder: [], defaultOrder: [], relations: {}, relCache: {}, relComplete: {},
  overrides: {}, selectedRows: new Set(), rowFilter: "", rowCheck: [], checkSeq: 0,
  checkTimer: null, editorRow: null, rowBound: false,
};

/* ---------------- 基础 ---------------- */

async function api(path, body) {
  const sep = path.includes("?") ? "&" : "?";
  const res = await fetch(path + sep + "t=" + encodeURIComponent(TOKEN), {
    method: body === undefined ? "GET" : "POST",
    headers: { "Content-Type": "application/json", "X-Token": TOKEN },
    body: body === undefined ? undefined : JSON.stringify(body || {}),
  });
  let data;
  try { data = await res.json(); } catch (e) { throw new Error("服务无响应（" + res.status + "）"); }
  if (!data.ok) throw new Error(data.error || "请求失败");
  return data.data;
}

async function apiDownload(path, body, filename) {
  const res = await fetch(path + "?t=" + encodeURIComponent(TOKEN), {
    method: "POST",
    headers: { "Content-Type": "application/json", "X-Token": TOKEN },
    body: JSON.stringify(body || {}),
  });
  const ct = res.headers.get("Content-Type") || "";
  // 后端出错时会返回 JSON，不能把它当文件存下来
  if (!res.ok || ct.includes("application/json")) {
    let detail = "HTTP " + res.status;
    try {
      const d = await res.json();
      if (d && d.error) detail = d.error;
    } catch (e) { /* ignore */ }
    alert("下载失败：" + detail);
    return;
  }
  const blob = await res.blob();
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 3000);
}

function msg(id, text, cls) {
  const el = $(id);
  el.textContent = text || "";
  el.className = "msg" + (cls ? " " + cls : "");
}

function unlock(...ids) { ids.forEach((id) => $(id).classList.remove("locked")); }
function lock(...ids) { ids.forEach((id) => $(id).classList.add("locked")); }

function esc(s) {
  return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

function toB64(buf) {
  const bytes = new Uint8Array(buf);
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

/* ---------------- 1. 登录 ---------------- */

async function checkStatus() {
  msg("statusMsg", "检查中…");
  $("btnStatus").disabled = true;
  try {
    const d = await api("/api/status");
    if (d.authenticated) {
      $("userBox").textContent = "已登录" + (d.user ? "：" + d.user : "");
      $("userBox").classList.add("on");
      msg("statusMsg", "登录态正常", "ok");
      $("loginHelp").classList.add("hidden");
      unlock("card2");
      if (!$("projectKey").value) $("projectKey").focus();
      loadProjects();
    } else {
      $("userBox").textContent = "未登录";
      $("userBox").classList.remove("on");
      msg("statusMsg", "未登录或已过期", "bad");
      $("loginHelp").classList.remove("hidden");
      lock("card2", "card3", "card4", "card5", "card6", "card7");
    }
  } catch (e) {
    msg("statusMsg", e.message, "bad");
    $("loginHelp").classList.remove("hidden");
  } finally {
    $("btnStatus").disabled = false;
  }
}

/* ---------------- 1.5 扫码登录（全图形化，不弹终端菜单） ---------------- */

const loginState = {
  deviceCode: "", clientId: "", interval: 5, left: 0,
  busy: false, done: false, pollTimer: null, tickTimer: null, gen: 0,
};

function openLogin() {
  $("loginModal").classList.remove("hidden");
  startLogin();
}

function closeLogin() {
  stopLoginTimers();
  $("loginModal").classList.add("hidden");
}

function stopLoginTimers() {
  clearTimeout(loginState.pollTimer);
  clearInterval(loginState.tickTimer);
  loginState.pollTimer = null;
  loginState.tickTimer = null;
}

async function startLogin() {
  stopLoginTimers();
  loginState.done = false;
  loginState.busy = false;
  loginState.deviceCode = "";
  $("qrCanvas").innerHTML = "";
  $("loginCode").textContent = "--------";
  $("loginLink").removeAttribute("href");
  $("loginStatus").className = "login-status";
  $("loginStatus").textContent = "正在生成二维码…";

  const gen = ++loginState.gen;
  try {
    const d = await api("/api/login/start", { host: $("hostSelect").value });
    if (gen !== loginState.gen) return;
    loginState.deviceCode = d.device_code;
    loginState.clientId = d.client_id;
    loginState.interval = Math.max(3, Number(d.interval) || 5);
    loginState.left = Number(d.expires_in) || 1800;
    $("loginCode").textContent = d.user_code || "--------";
    const url = d.verification_uri_complete || d.verification_uri || "";
    if (url) $("loginLink").href = url;
    drawQR(url);
    $("loginStatus").textContent = "等待扫码确认…（剩余 " + fmtLeft() + "）";
    startTick();
    schedulePoll();
  } catch (e) {
    $("loginStatus").className = "login-status bad";
    $("loginStatus").textContent = "生成失败：" + e.message;
  }
}

function drawQR(text) {
  const box = $("qrCanvas");
  box.innerHTML = "";
  if (!text) return;
  if (typeof QRCode === "undefined") {
    $("loginStatus").className = "login-status bad";
    $("loginStatus").textContent = "二维码组件未加载，请用下面的验证码在浏览器里登录";
    return;
  }
  // 260px + 白边，保证手机能扫
  new QRCode(box, {
    text, width: 260, height: 260,
    colorDark: "#000000", colorLight: "#ffffff",
    correctLevel: QRCode.CorrectLevel.M,
  });
}

function fmtLeft() {
  const m = Math.floor(loginState.left / 60);
  const s = loginState.left % 60;
  return m + ":" + String(s).padStart(2, "0");
}

function startTick() {
  loginState.tickTimer = setInterval(() => {
    if (loginState.done) return;
    if (loginState.left > 0) loginState.left--;
    if (loginState.left <= 0) {
      stopLoginTimers();
      $("loginStatus").className = "login-status bad";
      $("loginStatus").textContent = "二维码已过期，点「刷新二维码」重新生成";
      return;
    }
    if (!loginState.busy) {
      $("loginStatus").textContent = "等待扫码确认…（剩余 " + fmtLeft() + "）";
    }
  }, 1000);
}

function schedulePoll() {
  loginState.pollTimer = setTimeout(pollLogin, loginState.interval * 1000);
}

async function pollLogin() {
  if (loginState.done || !loginState.deviceCode) return;
  const gen = loginState.gen;
  loginState.busy = true;
  let state = "";
  try {
    const d = await api("/api/login/poll", {
      device_code: loginState.deviceCode,
      client_id: loginState.clientId,
    });
    if (gen !== loginState.gen) return; // 期间刷新过二维码，丢弃旧结果
    loginState.busy = false;
    state = (d.status || "").toLowerCase();
    if (state === "success") {
      loginState.done = true;
      stopLoginTimers();
      $("loginStatus").className = "login-status ok";
      $("loginStatus").textContent = "登录成功" + (d.user ? "：" + d.user : "");
      setTimeout(() => { closeLogin(); checkStatus(); }, 900);
      return;
    }
    if (state === "error") {
      $("loginStatus").className = "login-status bad";
      $("loginStatus").textContent = "登录失败：" + (d.error || "未知错误");
      stopLoginTimers();
      return;
    }
    if (state === "slow_down") loginState.interval += 5;
    if (state === "expired_token" || state === "expired") {
      $("loginStatus").className = "login-status bad";
      $("loginStatus").textContent = "二维码已过期，点「刷新二维码」重新生成";
      stopLoginTimers();
      return;
    }
    if (state === "access_denied") {
      $("loginStatus").className = "login-status bad";
      $("loginStatus").textContent = "你已在手机上拒绝授权，可点「刷新二维码」重试";
      stopLoginTimers();
      return;
    }
    if (loginState.left > 0) schedulePoll();
  } catch (e) {
    if (gen !== loginState.gen) return;
    loginState.busy = false;
    if (loginState.left > 0) schedulePoll();
    else {
      $("loginStatus").className = "login-status bad";
      $("loginStatus").textContent = "网络异常：" + e.message;
    }
  }
}

async function copyLoginCode() {
  const code = ($("loginCode").textContent || "").trim();
  try {
    await navigator.clipboard.writeText(code);
    $("btnCopyCode").textContent = "已复制";
    setTimeout(() => { $("btnCopyCode").textContent = "复制"; }, 1500);
  } catch (e) {
    $("loginStatus").textContent = "请手动选中验证码复制：" + code;
  }
}

/* ---------------- 2. 空间 / 类型 ---------------- */

async function loadProjects() {
  try {
    const d = await api("/api/projects");
    const list = $("projectList");
    list.innerHTML = (d.projects || []).map((p) =>
      `<option value="${esc(p.key)}">${esc(p.name)}</option>`).join("");
    if (!$("projectKey").value && d.projects && d.projects.length === 1) {
      $("projectKey").value = d.projects[0].key;
    }
  } catch (e) { /* 列不出空间不影响手填 */ }
}


async function loadTypes() {
  const pk = $("projectKey").value.trim();
  if (!pk) { msg("metaMsg", "请先填空间 key", "bad"); return; }
  $("btnTypes").disabled = true;
  msg("metaMsg", "加载工作项类型…");
  try {
    const d = await api("/api/types", { project_key: pk });
    const sel = $("typeSelect");
    sel.innerHTML = '<option value="">请选择…</option>' +
      d.types.map((t) => `<option value="${esc(t.type_key)}">${esc(t.type_name)}（${esc(t.type_key)}）</option>`).join("");
    sel.disabled = false;
    $("btnMeta").disabled = false;
    state.projectKey = pk;
    msg("metaMsg", `共 ${d.types.length} 种工作项类型`, "ok");
    // 只有一种类型时自动选中
    if (d.types.length === 1) { sel.value = d.types[0].type_key; }
  } catch (e) {
    msg("metaMsg", e.message, "bad");
  } finally {
    $("btnTypes").disabled = false;
  }
}

async function loadMeta(refresh = false) {
  const pk = $("projectKey").value.trim();
  const wt = $("typeSelect").value;
  if (!pk || !wt) { msg("metaMsg", "请选择空间和工作项类型", "bad"); return; }
  $("btnMeta").disabled = true;
  $("btnRefreshMeta").disabled = true;
  msg("metaMsg", refresh ? "重新拉取字段/选项/角色…" : "读取字段配置（首次约 5~10 秒）…");
  try {
    const d = await api("/api/meta", { project_key: pk, work_item_type: wt, refresh });
    state.projectKey = pk;
    state.type = wt;
    state.meta = d;
    state.defaultOrder = d.items.map((i) => i.field_key);
    state.fieldOrder = [...state.defaultOrder];
    $("btnResetOrder").disabled = false;
    $("btnRefreshMeta").disabled = false;
    renderFields(d);
    renderTemplateOptions(d);
    unlock("card3", "card4");
    $("btnTemplate").disabled = false;
    let info = `已加载 ${d.items.length} 个字段`;
    const profCount = d.profiles ? Object.keys(d.profiles).length : 0;
    if (profCount) info += `，已同步 ${profCount} 种测试单类型的必填规则`;
    if (d.missing && d.missing.length) info += `，${d.missing.length} 个已删除字段已跳过`;
    if (d.new_fields && d.new_fields.length) {
      info += `，新增：${d.new_fields.join("、")}（已加到最后）`;
    }
    msg("metaMsg", info, d.new_fields && d.new_fields.length ? "warn" : "ok");
    if (state.upload) renderMapping();
  } catch (e) {
    msg("metaMsg", e.message, "bad");
  } finally {
    $("btnMeta").disabled = false;
  }
}

function orderStorageKey() {
  return "meegle_order::" + state.projectKey + "::" + state.type;
}

function loadOrder(defaults) {
  try {
    const saved = JSON.parse(localStorage.getItem(orderStorageKey()) || "[]");
    const valid = saved.filter((k) => defaults.includes(k));
    return [...valid, ...defaults.filter((k) => !valid.includes(k))];
  } catch (e) {
    return [...defaults];
  }
}

function saveOrder() {
  try { localStorage.setItem(orderStorageKey(), JSON.stringify(state.fieldOrder)); } catch (e) { /* ignore */ }
}

function moveField(key, delta) {
  const arr = state.fieldOrder;
  const i = arr.indexOf(key);
  const j = i + delta;
  if (i < 0 || j < 0 || j >= arr.length) return;
  [arr[i], arr[j]] = [arr[j], arr[i]];
  saveOrder();
  persistLayout();
  renderFields(state.meta);
}

function resetOrder() {
  state.fieldOrder = [...state.defaultOrder];
  saveOrder();
  persistLayout();
  renderFields(state.meta);
}

function renderFields(d) {
  const items = d.items || [];
  const byKey = {};
  items.forEach((it) => { byKey[it.field_key] = it; });
  // 保持用户调过的顺序，新字段补在末尾
  const keys = state.fieldOrder.filter((k) => byKey[k]);
  items.forEach((it) => { if (!keys.includes(it.field_key)) keys.push(it.field_key); });
  state.fieldOrder = keys;
  const ordered = keys.map((k) => byKey[k]);

  const req = ordered.filter((f) => f.required);
  const opt = ordered.length - req.length;
  let text = `页面字段 ${ordered.length} 个：必填 ${req.length}、选填 ${opt}`;
  if (d.fetched_at) text += ` · 同步于 ${d.fetched_at}`;
  if (d.new_fields && d.new_fields.length) text += ` · 新增 ${d.new_fields.length} 个`;
  msg("fieldsSummary", text, d.new_fields && d.new_fields.length ? "warn" : "");

  const rows = ordered.map((f, idx) => {
    const tags = [];
    if (f.kind === "role") tags.push('<span class="tag">角色</span>');
    if (f.writable === false) tags.push('<span class="tag gray">不可导入</span>');
    if (f.is_new) tags.push('<span class="tag warn">新增</span>');
    const opts = (f.options || []);
    return `<tr>
      <td class="ord">${idx + 1}</td>
      <td>${esc(f.field_name)}${f.required ? '<span class="req">*</span>' : ""} ${tags.join("")}</td>
      <td>${f.required ? '<span class="req">必填</span>' : "选填"}</td>
      <td>${esc(f.field_type || f.kind || "")}</td>
      <td>${opts.length ? esc(opts.slice(0, 4).join(" / ")) + (opts.length > 4 ? " …" : "") : ""}</td>
      <td class="sample">${esc(f.note || "")}</td>
      <td class="ord-btns">
        <button class="btn mini" data-move="-1" data-key="${esc(f.field_key)}" title="上移">↑</button>
        <button class="btn mini" data-move="1" data-key="${esc(f.field_key)}" title="下移">↓</button>
      </td>
    </tr>`;
  }).join("");
  $("fieldsTable").innerHTML =
    "<thead><tr><th>#</th><th>字段</th><th>要求</th><th>类型</th><th>可选值</th><th>说明</th><th>排序</th></tr></thead><tbody>" + rows + "</tbody>";
  $("fieldsTable").querySelectorAll("button[data-move]").forEach((b) => {
    b.addEventListener("click", () => moveField(b.dataset.key, Number(b.dataset.move)));
  });
}

let layoutTimer = null;

function persistLayout() {
  clearTimeout(layoutTimer);
  layoutTimer = setTimeout(async () => {
    if (!state.meta) return;
    const items = state.fieldOrder.map((k) => {
      const it = (state.meta.items || []).find((x) => x.field_key === k) || {};
      return {
        kind: it.kind || "field",
        key: it.role_id || it.field_key || k,
        name: it.field_name, required: !!it.required,
        note: it.note || "", default: it.default || "",
      };
    });
    try {
      await api("/api/layout", {
        project_key: state.projectKey, work_item_type: state.type, items,
      });
    } catch (e) { /* 本地顺序仍保存在 localStorage */ }
  }, 300);
}

function renderTemplateOptions(d) {
  const sel = $("templateSelect");
  const cur = sel.value;
  const opts = d.template_options || [];
  sel.innerHTML = '<option value="">（不指定）</option>' +
    opts.map((o) => `<option value="${esc(o)}">${esc(o)}</option>`).join("");
  const tplItem = (d.items || []).find((i) => i.field_key === "template");
  const prefer = (tplItem && tplItem.default) || opts[0] || "";
  if (cur && opts.includes(cur)) sel.value = cur;
  else if (prefer) sel.value = prefer;
}

/* ---------------- 3. 模板 ---------------- */

async function downloadTemplate() {
  try {
    await apiDownload("/api/template",
      { project_key: state.projectKey, work_item_type: state.type, order: state.fieldOrder },
      `批量建单模板_${state.type}.xlsx`);
  } catch (e) { alert(e.message); }
}

/* ---------------- 4. 上传 ---------------- */

async function onFile(file, sheet) {
  if (!file) return;
  state.file = file;
  msg("uploadMsg", "解析中…");
  const buf = await file.arrayBuffer();
  try {
    const d = await api("/api/upload", {
      name: file.name, data_b64: toB64(buf), sheet: sheet || "",
      project_key: state.projectKey, work_item_type: state.type,
    });
    state.upload = d;
    state.relations = {};
    state.overrides = {};
    state.selectedRows = new Set();
    state.rowFilter = "";
    state.rowCheck = [];
    state.editorRow = null;
    state.mapping = { ...d.mapping };
    state.autoMapping = { ...d.mapping };

    // 工作表切换
    if (d.sheets.length > 1) {
      const sel = $("sheetSelect");
      sel.innerHTML = d.sheets.map((s) => `<option ${s === d.sheet ? "selected" : ""}>${esc(s)}</option>`).join("");
      $("sheetWrap").classList.remove("hidden");
    } else {
      $("sheetWrap").classList.add("hidden");
    }

    renderPreview();
    renderMapping();
    unlock("card5", "card6", "card7");
    msg("uploadMsg", `${file.name} · 共 ${d.row_count} 行数据`, "ok");
  } catch (e) {
    msg("uploadMsg", e.message, "bad");
  }
}

function renderPreview() {
  const d = state.upload;
  if (!d) return;
  const head = d.headers.map((h, i) => `<th>${esc(h || "列" + (i + 1))}</th>`).join("");
  const all = d.rows || d.preview || [];
  const body = all.map((r) => "<tr>" + d.headers.map((_, i) =>
    `<td class="sample">${esc(r[i] || "")}</td>`).join("") + "</tr>").join("");
  $("previewTable").innerHTML = `<thead><tr>${head}</tr></thead><tbody>${body}</tbody>`;
  $("previewWrap").classList.remove("hidden");
}

/* ---------------- 5. 映射 ---------------- */

function renderMapping() {
  const d = state.upload, meta = state.meta;
  if (!d || !meta) return;
  const sample = d.preview[0] || [];
  const items = meta.items || [];
  const opts = (idx) => {
    const cur = state.mapping[String(idx)] || "";
    const list = ['<option value="">— 忽略此列 —</option>'];
    for (const f of items) {
      if (f.writable === false) continue; // 附件等不可导入
      const sel = f.field_key === cur ? " selected" : "";
      const suffix = f.kind === "role" ? "（角色）" : (f.required ? " *" : "");
      list.push(`<option value="${esc(f.field_key)}"${sel}>${esc(f.field_name)}${suffix}</option>`);
    }
    return list.join("");
  };
  const rows = d.headers.map((h, i) => {
    const cur = state.mapping[String(i)] || "";
    const f = items.find((x) => x.field_key === cur);
    return `<tr>
      <td>${esc(h || "列" + (i + 1))}</td>
      <td class="sample">${esc(sample[i] || "")}</td>
      <td><select data-col="${i}">${opts(i)}</select></td>
      <td>${f ? (f.required ? '<span class="req">必填</span>' : "") : '<span class="muted">未映射</span>'}</td>
    </tr>`;
  }).join("");
  $("mapTable").innerHTML =
    "<thead><tr><th>表格列</th><th>示例值</th><th>映射到字段</th><th>要求</th></tr></thead><tbody>" + rows + "</tbody>";

  $("mapTable").querySelectorAll("select").forEach((sel) => {
    sel.addEventListener("change", (e) => {
      const idx = e.target.dataset.col;
      if (e.target.value) state.mapping[idx] = e.target.value;
      else delete state.mapping[idx];
      updateMapMsg();
      renderMapping();
    });
  });
  updateMapMsg();
  renderRowCheck();
}

function cellText(v) {
  return v == null ? "" : String(v);
}

function relationColumns() {
  const meta = state.meta, up = state.upload;
  if (!meta || !up) return [];
  const byKey = {};
  (meta.items || []).forEach((it) => { byKey[it.field_key] = it; });
  const out = [];
  Object.keys(state.mapping).forEach((idx) => {
    const it = byKey[state.mapping[idx]];
    if (!it || !it.related || !it.related.length || it.writable === false) return;
    out.push({ index: Number(idx), item: it, header: up.headers[Number(idx)] || "" });
  });
  return out;
}

function relKey(project, target) { return project + "|" + target; }

function fillRelList(items) {
  $("relList").innerHTML = (items || []).slice(0, 200).map((o) =>
    `<option value="${esc(o.id)}">${esc(o.name)}</option>`).join("");
}

async function ensureRelCache(project, target) {
  const key = relKey(project, target);
  if (state.relCache[key]) return state.relCache[key];
  const d = await api("/api/relation_options", {
    project_key: project, target_type: target, q: "",
  });
  state.relCache[key] = d.items || [];
  state.relComplete[key] = !!d.complete;
  return state.relCache[key];
}

function bindRelationInput(inp) {
  let timer = null;
  const project = inp.dataset.project, target = inp.dataset.target;
  const key = relKey(project, target);
  const loadAll = async () => {
    try {
      const items = await ensureRelCache(project, target);
      fillRelList(items);
      inp.setAttribute("list", "relList");
    } catch (e) { /* 失败不阻塞 */ }
  };
  inp.addEventListener("focus", loadAll);
  inp.addEventListener("input", () => {
    inp.setAttribute("list", "relList");
    const q = inp.value.trim().toLowerCase();
    const items = state.relCache[key] || [];
    if (!items.length) { loadAll(); return; }
    if (!q) { fillRelList(items); return; }
    const hits = items.filter((o) => String(o.name).toLowerCase().includes(q));
    fillRelList(hits);
    if (hits.length || state.relComplete[key]) return;
    clearTimeout(timer);
    timer = setTimeout(async () => {
      try {
        const d = await api("/api/relation_options", {
          project_key: project, target_type: target, q: inp.value.trim(),
        });
        fillRelList(d.items || []);
        inp.setAttribute("list", "relList");
      } catch (e) { /* ignore */ }
    }, 300);
  });
  inp.addEventListener("change", () => {
    const r = inp.dataset.row, f = inp.dataset.field, v = inp.value.trim();
    state.relations[r] = state.relations[r] || {};
    if (v) state.relations[r][f] = v;
    else delete state.relations[r][f];
  });
}

function renderRelations() {
  const cols = relationColumns();
  const box = $("relBox");
  if (!cols.length) { box.classList.add("hidden"); return; }
  box.classList.remove("hidden");
  const rows = state.upload.rows || [];
  const head = `<tr><th>行</th><th>标题</th>${cols.map((c) =>
    `<th>${esc(c.item.field_name)}</th>`).join("")}</tr>`;
  const body = rows.map((row, r) => {
    const cells = cols.map((c) => {
      const rel = (state.relations[r] || {})[c.item.field_key] || "";
      const existing = rel || cellText(row[c.index]);
      const target = (c.item.related && c.item.related[0]) || {};
      return `<td><input class="rel-input" data-row="${r}"
        data-field="${esc(c.item.field_key)}"
        data-target="${esc(target.work_item_type || "")}"
        data-project="${esc(target.project_key || state.projectKey)}"
        value="${esc(existing)}" placeholder="搜索工作项…" autocomplete="off"></td>`;
    }).join("");
    return `<tr><td class="ord">${r + 2}</td>
      <td class="sample">${esc(cellText(row[0]))}</td>${cells}</tr>`;
  }).join("");
  $("relTable").innerHTML = `<thead>${head}</thead><tbody>${body}</tbody>`;
  $("relTable").querySelectorAll(".rel-input").forEach(bindRelationInput);
}

function updateMapMsg() {
  const meta = state.meta;
  if (!meta) return;
  const mapped = new Set(Object.values(state.mapping));
  const items = meta.items || [];
  const missing = items.filter((f) => f.required && f.writable !== false
    && !mapped.has(f.field_key) && f.field_key !== "template");
  if (!missing.length) msg("mapMsg", `已映射 ${mapped.size} 个字段，必填项齐全`, "ok");
  else msg("mapMsg", `已映射 ${mapped.size} 个字段；还缺必填：${missing.map((f) => f.field_name).join("、")}`, "warn");
}

function remap() {
  if (!state.upload) return;
  state.mapping = { ...state.autoMapping };
  renderMapping();
}

/* ---------------- 5.5 逐行检查 / 行内编辑 / 批量修改 ---------------- */

function fieldByKey(key) {
  return ((state.meta && state.meta.items) || []).find((it) => it.field_key === key);
}

function relationFields() {
  return ((state.meta && state.meta.items) || []).filter(
    (it) => it.kind === "field" && (it.related || []).length && it.writable !== false);
}

function mappedIndex(key) {
  for (const [i, k] of Object.entries(state.mapping)) {
    if (k === key) return Number(i);
  }
  return -1;
}

function rowValue(index, key) {
  const o = state.overrides[String(index)];
  if (o && Object.prototype.hasOwnProperty.call(o, key)) return o[key];
  const ci = mappedIndex(key);
  if (ci < 0 || !state.upload) return "";
  const row = state.upload.rows[index] || [];
  return row[ci] == null ? "" : String(row[ci]);
}

function setOverride(index, key, value) {
  const r = String(index);
  state.overrides[r] = state.overrides[r] || {};
  const s = value == null ? "" : String(value).trim();
  if (s) state.overrides[r][key] = s;
  else delete state.overrides[r][key];
  if (!Object.keys(state.overrides[r]).length) delete state.overrides[r];
  scheduleCheck();
}

function scheduleCheck() {
  clearTimeout(state.checkTimer);
  state.checkTimer = setTimeout(refreshRowCheck, 350);
}

function rowSearchText(r) {
  let t = (r.title || "") + " " + (r.type || "");
  for (const f of relationFields()) t += " " + rowValue(r.index, f.field_key);
  return t.toLowerCase();
}

async function refreshRowCheck() {
  if (!state.upload || !state.meta) return;
  const seq = ++state.checkSeq;
  msg("rowCheckMsg", "检查必填项…");
  try {
    const d = await api("/api/check_rows", {
      upload_id: state.upload.upload_id, mapping: state.mapping,
      overrides: state.overrides,
    });
    if (seq !== state.checkSeq) return;
    state.rowCheck = d.rows || [];
    renderRowCheckTable();
    msg("rowCheckMsg", d.profiles_ready ? "" : "必填规则首次加载中，稍后再看…", d.profiles_ready ? "ok" : "warn");
  } catch (e) {
    if (seq === state.checkSeq) msg("rowCheckMsg", e.message, "bad");
  }
}

function renderRowCheck() {
  const box = $("rowCheckBox");
  if (!box) return;
  if (!state.upload || !state.meta) { box.classList.add("hidden"); return; }
  box.classList.remove("hidden");
  if (!state.rowBound) { bindRowCheckOnce(); state.rowBound = true; }
  refreshRowCheck();
}

function renderRowCheckTable() {
  const all = state.rowCheck || [];
  const filter = state.rowFilter.trim().toLowerCase();
  const visible = all.filter((r) => !filter || rowSearchText(r).includes(filter));
  const head = `<thead><tr>
    <th><input type="checkbox" id="rowCheckAll"></th><th>行</th><th>标题</th>
    <th>测试单类型</th><th>缺失必填</th><th></th></tr></thead>`;
  const body = visible.map((r) => {
    const picked = state.selectedRows.has(r.index) ? " picked" : "";
    const bad = r.missing_count ? " bad" : "";
    const miss = r.missing.length
      ? r.missing.map((m) => `<span class="miss-tag">${esc(m.field_name)}</span>`).join("")
      : '<span class="ok-tag">完整</span>';
    return `<tr class="${picked}${bad}" data-row="${r.index}">
      <td class="sel"><input type="checkbox" class="row-pick" data-index="${r.index}" ${picked ? "checked" : ""}></td>
      <td>${r.row}</td>
      <td class="sample"><a href="javascript:void 0" class="row-title" data-index="${r.index}">${esc(r.title || "")}</a></td>
      <td>${esc(r.type || "")}</td>
      <td class="status">${miss}</td>
      <td><button class="btn small row-edit" data-index="${r.index}">编辑</button></td>
    </tr>`;
  }).join("");
  $("rowCheckTable").innerHTML = head + `<tbody>${body || '<tr><td colspan="6" class="muted">没有匹配的 BUG</td></tr>'}</tbody>`;
  const missing = all.filter((r) => r.missing_count).length;
  $("checkSummary").textContent = `共 ${all.length} 行，${missing} 行缺必填`;
  $("rowSelCount").textContent = `已选 ${state.selectedRows.size} 行`;
  $("btnBulkEdit").disabled = state.selectedRows.size === 0;
  const selAll = $("rowCheckAll");
  if (selAll) {
    const picked = visible.filter((r) => state.selectedRows.has(r.index)).length;
    selAll.checked = visible.length > 0 && picked === visible.length;
  }
  $("rowSelectAll").checked = visible.length > 0 &&
    visible.every((r) => state.selectedRows.has(r.index));
}

function bindRowCheckOnce() {
  $("rowFilter").addEventListener("input", () => {
    state.rowFilter = $("rowFilter").value;
    renderRowCheckTable();
  });
  $("rowSelectAll").addEventListener("change", () => {
    const on = $("rowSelectAll").checked;
    const filter = state.rowFilter.trim().toLowerCase();
    const visible = (state.rowCheck || []).filter((r) => !filter || rowSearchText(r).includes(filter));
    for (const r of visible) {
      if (on) state.selectedRows.add(r.index); else state.selectedRows.delete(r.index);
    }
    renderRowCheckTable();
  });
  $("btnBulkEdit").addEventListener("click", openBulk);
  $("bulkCancel").addEventListener("click", closeBulk);
  $("bulkMask").addEventListener("click", closeBulk);
  $("bulkConfirm").addEventListener("click", confirmBulk);
  $("rowCheckTable").addEventListener("change", (e) => {
    if (!e.target.classList.contains("row-pick")) return;
    const i = Number(e.target.dataset.index);
    if (e.target.checked) state.selectedRows.add(i); else state.selectedRows.delete(i);
    renderRowCheckTable();
  });
  $("rowCheckTable").addEventListener("click", (e) => {
    const el = e.target.closest(".row-title, .row-edit");
    if (el) openRowEditor(Number(el.dataset.index));
  });
  $("rowCheckTable").addEventListener("change", (e) => {
    if (e.target.id !== "rowCheckAll") return;
    const visible = (state.rowCheck || []).filter((r) =>
      !state.rowFilter || rowSearchText(r).includes(state.rowFilter.trim().toLowerCase()));
    for (const r of visible) {
      if (e.target.checked) state.selectedRows.add(r.index); else state.selectedRows.delete(r.index);
    }
    renderRowCheckTable();
  });
}

function controlHtml(it, index, val) {
  const key = esc(it.field_key);
  const ft = String(it.field_type || "");
  const role = it.kind === "role";
  if (role) return `<input data-key="${key}" value="${esc(val)}" placeholder="姓名或企业邮箱">`;
  if ((it.related || []).length) {
    const rel = it.related[0] || {};
    return `<input class="edit-rel" data-key="${key}" data-project="${esc(rel.project_key || "")}"
      data-target="${esc(rel.work_item_type || "")}" value="${esc(val)}"
      placeholder="搜索工作项…" autocomplete="off">`;
  }
  if (ft === "multi-file") return `<input value="不可导入" disabled>`;
  if (ft === "tree-select" || ft === "tree-multi-select") {
    const leaves = (it.options || []).filter((o) => o.path && o.path.length);
    if (leaves.length) return `<div class="tree-cascade" data-tree-key="${key}"></div>`;
  }
  if (ft.includes("select") || ft.includes("option") || ft.includes("template") || it.field_key === "template") {
    const opts = it.options || [];
    if (ft.includes("multi")) {
      const cur = val.split(/[,，、;；/|]+/).map((s) => s.trim()).filter(Boolean);
      return `<select data-key="${key}" multiple size="4">${opts.map((o) =>
        `<option value="${esc(o.label)}" ${cur.includes(o.label) ? "selected" : ""}>${esc(o.label)}</option>`).join("")}</select>`;
    }
    return `<select data-key="${key}"><option value="">请选择</option>${opts.map((o) =>
      `<option value="${esc(o.label)}" ${o.label === val ? "selected" : ""}>${esc(o.label)}</option>`).join("")}</select>`;
  }
  if (ft === "multi-text") return `<textarea data-key="${key}" placeholder="多个用换行分隔">${esc(val)}</textarea>`;
  return `<input data-key="${key}" value="${esc(val)}">`;
}

function initTree(box, item, index) {
  const key = item.field_key;
  const leaves = (item.options || []).filter((o) => o.path && o.path.length);
  const depth = Math.max(...leaves.map((o) => o.path.length));
  const cur = rowValue(index, key);
  const curPath = (leaves.find((o) => o.label === cur) || {}).path || null;
  box.innerHTML = Array.from({ length: depth }, (_, level) =>
    `<select data-level="${level}"></select>`).join("");
  const fill = (level, prefix) => {
    const sel = box.querySelector(`[data-level="${level}"]`);
    if (!sel) return;
    const labels = [];
    for (const o of leaves) {
      if (!prefix.every((x, i) => o.path[i] && o.path[i].label === x)) continue;
      const lab = o.path[level] && o.path[level].label;
      if (lab && !labels.includes(lab)) labels.push(lab);
    }
    sel.innerHTML = '<option value="">请选择</option>' +
      labels.map((l) => `<option value="${esc(l)}">${esc(l)}</option>`).join("");
    if (curPath && curPath[level]) sel.value = curPath[level].label;
    else sel.value = "";
  };
  fill(0, []);
  if (curPath && curPath[0]) {
    for (let level = 1; level < depth; level += 1) {
      fill(level, curPath.slice(0, level).map((p) => p.label));
    }
    for (let level = 0; level < depth; level += 1) {
      const sel = box.querySelector(`[data-level="${level}"]`);
      if (sel && curPath[level]) sel.value = curPath[level].label;
    }
  }
  box.addEventListener("change", (e) => {
    const level = Number(e.target.dataset.level);
    const labels = [];
    for (let i = 0; i <= level; i += 1) {
      const v = box.querySelector(`[data-level="${i}"]`).value;
      if (!v) break;
      labels.push(v);
    }
    for (let i = level + 1; i < depth; i += 1) {
      box.querySelector(`[data-level="${i}"]`).innerHTML = '<option value="">请选择</option>';
    }
    if (labels.length === level + 1 && level < depth - 1) fill(level + 1, labels);
    if (labels.length === depth) setOverride(index, key, labels[depth - 1]);
  });
}

function openRowEditor(index, keepScroll) {
  if (!state.upload || !state.meta) return;
  state.editorRow = index;
  const row = state.upload.rows[index] || [];
  const title = row[0] == null ? "" : String(row[0]);
  const items = state.meta.items || [];
  const html = items.map((it) => {
    const val = rowValue(index, it.field_key);
    return `<div class="field-edit"><label>${esc(it.field_name)}${it.required ? ' <span class="req">*</span>' : ""}</label>${controlHtml(it, index, val)}</div>`;
  }).join("");
  const box = $("rowEditor");
  box.innerHTML = `<div class="row-editor-head"><b>第 ${index + 2} 行：${esc(title)}</b>
    <button class="btn small" id="rowEditClose">收起</button></div>
    <div class="row-editor-grid">${html}</div><datalist id="editRelList"></datalist>`;
  box.classList.remove("hidden");
  if (!keepScroll) box.scrollIntoView({ block: "nearest" });
  $("rowEditClose").onclick = () => { state.editorRow = null; box.classList.add("hidden"); box.innerHTML = ""; };

  box.querySelectorAll("[data-tree-key]").forEach((el) => {
    const item = fieldByKey(el.dataset.treeKey);
    if (item) initTree(el, item, index);
  });
  box.querySelectorAll("[data-key]").forEach((el) => {
    if (el.classList.contains("edit-rel")) bindEditRel(el, index);
    el.addEventListener("change", () => {
      let v = el.value;
      if (el.multiple) v = [...el.selectedOptions].map((o) => o.value).join(",");
      setOverride(index, el.dataset.key, v);
    });
  });
}

function bindEditRel(inp, index) {
  let timer = null;
  const load = async () => {
    try {
      const d = await api("/api/relation_options", {
        project_key: inp.dataset.project, target_type: inp.dataset.target,
        q: inp.value.trim(),
      });
      $("editRelList").innerHTML = (d.items || []).map((o) =>
        `<option value="${esc(o.id)}">${esc(o.name)}</option>`).join("");
      inp.setAttribute("list", "editRelList");
    } catch (e) { /* ignore */ }
  };
  inp.addEventListener("focus", load);
  inp.addEventListener("input", () => {
    inp.setAttribute("list", "editRelList");
    clearTimeout(timer);
    timer = setTimeout(load, 250);
  });
}

function openBulk() {
  if (!state.selectedRows.size) return;
  const fields = relationFields();
  if (!fields.length) { msg("rowCheckMsg", "没有可批量修改的关联字段", "warn"); return; }
  $("bulkCount").textContent = `将修改已选中的 ${state.selectedRows.size} 行`;
  $("bulkFields").innerHTML = fields.map((it) => `<div class="bulk-field">
    <label>${esc(it.field_name)}</label>
    <div class="bulk-line">
      <input class="bulk-rel" data-key="${esc(it.field_key)}"
        data-project="${esc((it.related[0] || {}).project_key || "")}"
        data-target="${esc((it.related[0] || {}).work_item_type || "")}"
        placeholder="输入关键字搜索飞书工作项…" autocomplete="off">
      <button type="button" class="btn small bulk-search">搜索</button>
    </div>
  </div>`).join("");
  $("bulkList").innerHTML = "";
  $("bulkFields").querySelectorAll(".bulk-rel").forEach(bindBulkRel);
  $("bulkFields").querySelectorAll(".bulk-search").forEach((btn) => {
    btn.addEventListener("click", () => btn.parentElement.querySelector(".bulk-rel").focus());
  });
  $("bulkModal").classList.remove("hidden");
}

function bindBulkRel(inp) {
  let timer = null;
  const load = async () => {
    try {
      const d = await api("/api/relation_options", {
        project_key: inp.dataset.project, target_type: inp.dataset.target,
        q: inp.value.trim(),
      });
      $("bulkList").innerHTML = (d.items || []).map((o) =>
        `<option value="${esc(o.id)}">${esc(o.name)}</option>`).join("");
      inp.setAttribute("list", "bulkList");
    } catch (e) { /* ignore */ }
  };
  inp.addEventListener("focus", load);
  inp.addEventListener("input", () => {
    inp.setAttribute("list", "bulkList");
    clearTimeout(timer);
    timer = setTimeout(load, 250);
  });
}

function closeBulk() { $("bulkModal").classList.add("hidden"); }

function confirmBulk() {
  const vals = [];
  $("bulkFields").querySelectorAll(".bulk-rel").forEach((inp) => {
    if (inp.value.trim()) vals.push([inp.dataset.key, inp.value.trim()]);
  });
  if (!vals.length) { closeBulk(); return; }
  for (const idx of state.selectedRows) {
    for (const [key, value] of vals) setOverride(idx, key, value);
  }
  closeBulk();
  if (state.editorRow != null) openRowEditor(state.editorRow, true);
  refreshRowCheck();
}

/* ---------------- 6. 校验 ---------------- */

async function validate() {
  if (!state.upload) return;
  $("btnValidate").disabled = true;
  msg("validateMsg", "校验中（人员字段需要联网解析）…");
  try {
    const d = await api("/api/validate", { upload_id: state.upload.upload_id, mapping: state.mapping, overrides: state.overrides });
    $("issuesBox").classList.remove("hidden");
    if (!d.issue_rows) {
      $("issuesBox").innerHTML = `<div class="issue s-ok">全部 ${d.valid_rows} 行通过校验，可以执行了。</div>`;
      msg("validateMsg", `通过 ${d.valid_rows} 行`, "ok");
    } else {
      $("issuesBox").innerHTML = d.issues.map((it) => `<div class="issue">
          <b>第 ${it.row} 行</b> ${esc(String(it.title || "").slice(0, 40))} ——
          ${it.errors.map((e) => esc(e.col) + "：" + esc(e.msg)).join("；")}
        </div>`).join("") + (d.issue_rows > d.issues.length ? `<div class="issue">…还有 ${d.issue_rows - d.issues.length} 行</div>` : "");
      msg("validateMsg", `${d.issue_rows} 行有问题，${d.valid_rows} 行通过`, "bad");
    }
  } catch (e) {
    msg("validateMsg", e.message, "bad");
  } finally {
    $("btnValidate").disabled = false;
  }
}

/* ---------------- 7. 执行 ---------------- */

async function start() {
  if (!state.upload) return;
  const dry = $("dryRun").checked;
  const body = {
    upload_id: state.upload.upload_id,
    mapping: state.mapping,
    overrides: state.overrides,
    dry_run: dry,
    delay: parseFloat($("delayInput").value || "0.4"),
    limit: parseInt($("limitInput").value || "0", 10),
    template_name: $("templateSelect").value,
  };
  $("btnStart").disabled = true;
  msg("resultMsg", "");
  try {
    const d = await api("/api/run", body);
    state.jobId = d.job_id;
    $("progressBox").classList.remove("hidden");
    $("logBox").classList.remove("hidden");
    $("resultWrap").classList.add("hidden");
    $("resultActions").classList.add("hidden");
    $("btnStop").classList.remove("hidden");
    $("logBox").textContent = dry ? "预演开始（不会真正创建）…" : "开始创建…";
    poll();
  } catch (e) {
    msg("resultMsg", e.message, "bad");
    $("btnStart").disabled = false;
  }
}

async function poll() {
  if (!state.jobId) return;
  try {
    const job = await api("/api/job?id=" + encodeURIComponent(state.jobId));
    const pct = job.total ? Math.round((job.done / job.total) * 100) : 0;
    $("progressFill").style.width = pct + "%";
    $("progressText").textContent =
      `${job.done} / ${job.total}（${pct}%） · 成功 ${job.ok} · 失败 ${job.failed} · 预演 ${job.planned}` +
      (job.elapsed ? ` · 用时 ${job.elapsed}s` : "");
    $("logBox").textContent = (job.logs || []).join("\n");
    $("logBox").scrollTop = $("logBox").scrollHeight;

    if (job.results && job.results.length) {
      const rows = job.results.slice(-100).map((r) => `<tr>
        <td>${r.row}</td><td class="sample">${esc(String(r.title || "").slice(0, 40))}</td>
        <td class="${r.status === "成功" ? "s-ok" : r.status === "失败" ? "s-bad" : ""}">${esc(r.status)}</td>
        <td class="mono">${esc(r.id || "")}</td>
        <td class="sample">${esc(r.error || "")}</td></tr>`).join("");
      $("resultTable").innerHTML =
        "<thead><tr><th>行号</th><th>标题</th><th>状态</th><th>工作项ID</th><th>错误</th></tr></thead><tbody>" + rows + "</tbody>";
      $("resultWrap").classList.remove("hidden");
    }

    if (["done", "stopped", "error"].includes(job.status)) {
      $("btnStart").disabled = false;
      $("btnStop").classList.add("hidden");
      $("resultActions").classList.remove("hidden");
      const label = job.status === "done" ? "执行完成" : job.status === "stopped" ? "已停止" : "执行异常";
      msg("resultMsg", `${label}：成功 ${job.ok}，失败 ${job.failed}${job.dry_run ? "，预演 " + job.planned : ""}`, job.status === "error" ? "bad" : "ok");
      if (job.error) $("logBox").textContent += "\n" + job.error;
      clearTimeout(state.timer);
      state.timer = null;
      return;
    }
    state.timer = setTimeout(poll, 800);
  } catch (e) {
    msg("resultMsg", e.message, "bad");
    $("btnStart").disabled = false;
  }
}

async function stop() {
  if (!state.jobId) return;
  try { await api("/api/stop", { id: state.jobId }); msg("resultMsg", "正在停止…", "warn"); }
  catch (e) { msg("resultMsg", e.message, "bad"); }
}

function downloadResults() {
  if (!state.jobId) return;
  apiDownload("/api/results", { id: state.jobId }, "批量建单结果.xlsx");
}

/* ---------------- 绑定 ---------------- */

document.addEventListener("DOMContentLoaded", () => {
  $("btnLogin").addEventListener("click", openLogin);
  $("btnCloseLogin").addEventListener("click", closeLogin);
  $("btnCloseLogin2").addEventListener("click", closeLogin);
  $("loginMask").addEventListener("click", closeLogin);
  $("btnRelogin").addEventListener("click", startLogin);
  $("btnCopyCode").addEventListener("click", copyLoginCode);
  $("hostSelect").addEventListener("change", startLogin);
  $("btnStatus").addEventListener("click", checkStatus);
  $("btnTypes").addEventListener("click", loadTypes);
  $("btnMeta").addEventListener("click", loadMeta);
  $("btnTemplate").addEventListener("click", downloadTemplate);
  $("btnResetOrder").addEventListener("click", resetOrder);
  $("btnRefreshMeta").addEventListener("click", () => loadMeta(true));
  $("btnRemap").addEventListener("click", remap);
  $("btnValidate").addEventListener("click", validate);
  $("btnStart").addEventListener("click", start);
  $("btnStop").addEventListener("click", stop);
  $("btnDownload").addEventListener("click", downloadResults);

  $("projectKey").addEventListener("keydown", (e) => { if (e.key === "Enter") loadTypes(); });
  $("typeSelect").addEventListener("change", () => { $("btnMeta").disabled = !$("typeSelect").value; });

  $("fileInput").addEventListener("change", (e) => {
    const f = e.target.files && e.target.files[0];
    if (f) onFile(f, "");
  });
  $("sheetSelect").addEventListener("change", (e) => {
    if (state.file) onFile(state.file, e.target.value);
  });

  const params = new URLSearchParams(location.search);
  if (params.has("demo") || params.has("guide")) bootDemo(params.has("guide"));
  else checkStatus();
  if (params.has("login")) openLogin();
});

/* ---------------- 演示模式（?demo=1，不联网） ---------------- */

async function bootDemo(isGuide = false) {
  let d;
  try { d = await (await fetch(isGuide ? "/static/guide.json" : "/static/demo.json")).json(); }
  catch (e) { msg("statusMsg", "演示数据缺失", "bad"); return; }

  state.projectKey = d.project_key;
  state.type = d.type;
  state.meta = d.meta;
  state.defaultOrder = (d.meta.items || []).map((i) => i.field_key);
  state.fieldOrder = [...state.defaultOrder];
  state.upload = d.upload;
  state.upload.rows = d.upload.rows || d.upload.preview || [];
  state.mapping = { ...d.mapping };
  state.autoMapping = { ...d.mapping };

  $("projectKey").value = d.project_key;
  $("typeSelect").innerHTML = `<option>${esc(d.type)}</option>`;
  $("typeSelect").disabled = false;
  $("btnMeta").disabled = false;
  $("userBox").textContent = isGuide ? "已登录：示例用户" : "已登录：演示用户";
  $("userBox").classList.add("on");
  msg("statusMsg", isGuide ? "登录态正常（界面示意，数据为示例）" : "登录态正常（演示数据，不会调用 API）", "ok");
  unlock("card2", "card3", "card4", "card5", "card6", "card7");
  $("btnTemplate").disabled = false;
  $("btnResetOrder").disabled = false;
  $("btnRefreshMeta").disabled = false;

  renderFields(d.meta);
  renderTemplateOptions(d.meta);
  if (d.meta.template_options.length) $("templateSelect").value = d.meta.template_options[0];
  renderPreview();
  if (isGuide) $("previewWrap").classList.add("hidden");
  renderMapping();
  msg("uploadMsg", `${d.upload.filename} · 共 ${d.upload.row_count} 行数据${isGuide ? "（示例）" : "（演示）"}`, "ok");

  $("issuesBox").classList.remove("hidden");
  if (!d.issues || !d.issues.length) {
    $("issuesBox").innerHTML = `<div class="issue s-ok">全部 ${d.upload.row_count || ""} 行通过校验，可以执行了。</div>`;
  } else {
    $("issuesBox").innerHTML = d.issues.map((it) => `<div class="issue">
      <b>第 ${it.row} 行</b> ${esc(it.title)} ——
      ${it.errors.map((e) => esc(e.col) + "：" + esc(e.msg)).join("；")}
    </div>`).join("");
  }
  msg("validateMsg", d.validateMsg || "2 行通过，1 行有问题（演示）", d.issues && d.issues.length ? "warn" : "ok");

  $("progressBox").classList.remove("hidden");
  $("progressFill").style.width = "100%";
  $("progressText").textContent = d.progressText || "3 / 3（100%） · 成功 2 · 失败 1 · 预演 0 · 用时 12.4s";
  $("resultWrap").classList.remove("hidden");
  $("resultTable").innerHTML = "<thead><tr><th>行号</th><th>标题</th><th>状态</th><th>工作项ID</th><th>错误</th></tr></thead><tbody>" +
    d.results.map((r) => `<tr><td>${r.row}</td><td class="sample">${esc(r.title)}</td>
      <td class="${r.status === "成功" ? "s-ok" : "s-bad"}">${esc(r.status)}</td>
      <td class="mono">${esc(r.id)}</td><td class="sample">${esc(r.error)}</td></tr>`).join("") + "</tbody>";
  $("logBox").classList.remove("hidden");
  $("logBox").textContent = d.logsText || [
    "行2 成功：6812345678901 蓝牙偶发断连-001",
    "行3 成功：6812345678902 U盘识别失败-002",
    "行4 失败：机型：不是 机型 的合法选项",
  ].join("\n");
  $("resultActions").classList.remove("hidden");
  msg("resultMsg", d.resultMsg || "执行完成：成功 2，失败 1（演示）", "ok");
}
