// 飞书多维表「视图同步」插件
// 读取当前视图（SDK，免密钥）→ 同步到目标（同文件走 SDK 免密钥；跨文件/电子表格走开放 API）。
// 注意：本文件在飞书多维表「自定义插件」的 iframe 中运行时，bitable 才存在；
// 单独打开只会看到友好提示，不会崩溃。

const state = {
  bitable: null,
  ready: false,
  source: null,        // { tableName, viewName, fields:[{id,name,type}], rows:[{recordId, values:{name:val}}] }
  target: { mode: 'same', type: 'bitable' },
  targetFields: [],    // [{id, name}] 同文件/跨多维表用；电子表格时存列字母
  mapping: [],         // [{src, tgt}]
  keyField: '',
  syncMode: 'full',
  token: null,
  tokenExpire: 0,
};

const $ = (id) => document.getElementById(id);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ---------------- UI 基础 ----------------
function setStatus(text, kind) {
  const el = $('status');
  el.textContent = text;
  el.className = 'status' + (kind ? ' ' + kind : '');
}
function log(msg, kind) {
  const el = $('log');
  const line = document.createElement('div');
  if (kind) line.className = kind;
  const t = new Date().toLocaleTimeString('zh-CN', { hour12: false });
  line.textContent = `[${t}] ${msg}`;
  el.appendChild(line);
  el.scrollTop = el.scrollHeight;
}
function showResult(msg, err) {
  const el = $('result');
  el.textContent = msg;
  el.className = 'result' + (err ? ' err' : '');
  el.classList.remove('hidden');
}

// ---------------- 开放 API 客户端（跨文件 / 电子表格） ----------------
const API_BASE = 'https://open.feishu.cn/open-apis';

async function getToken() {
  const now = Date.now();
  if (state.token && now < state.tokenExpire - 5000) return state.token;
  const id = $('appId').value.trim();
  const secret = $('appSecret').value.trim();
  if (!id || !secret) throw new Error('请先填写 App ID 与 App Secret');
  const r = await fetch(`${API_BASE}/auth/v3/tenant_access_token/internal`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
    body: JSON.stringify({ app_id: id, app_secret: secret }),
  });
  const j = await r.json();
  if (j.code !== 0) throw new Error('获取 tenant_access_token 失败：' + j.msg);
  state.token = j.tenant_access_token;
  state.tokenExpire = now + (j.expire || 7200) * 1000;
  return state.token;
}

let _lastWrite = 0;
async function api(method, path, { body, params, throttle = false } = {}) {
  if (throttle) {
    const wait = 800 - (Date.now() - _lastWrite);
    if (wait > 0) await sleep(wait);
  }
  const token = await getToken();
  const url = new URL(API_BASE + path);
  if (params) Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, v));
  const headers = { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json; charset=utf-8' };
  let attempt = 0;
  while (true) {
    const r = await fetch(url, { method, headers, body: body ? JSON.stringify(body) : undefined });
    const j = await r.json().catch(() => ({}));
    if (j.code === 0) { _lastWrite = Date.now(); return j.data; }
    if ([1254290, 1254291, 1254607].includes(j.code) && attempt < 5) {
      const d = 1000 * (1 << attempt);
      log(`限流(${j.code})，退避 ${d}ms 重试…`, 'err');
      await sleep(d); attempt++; continue;
    }
    throw new Error(`接口失败 [${method} ${path}] code=${j.code} ${j.msg}`);
  }
}

// ---------------- 读取当前视图（SDK，免密钥） ----------------
async function loadSource() {
  const base = state.bitable.base;
  const table = await base.getActiveTable();
  let view, viewId;
  try {
    view = await base.getActiveView();
    viewId = view.id;
  } catch {
    const sel = await base.getSelection();
    viewId = sel.viewId;
  }
  const fields = (await table.getFields()) || [];
  const fieldMeta = fields.map((f) => ({ id: f.id, name: f.name, type: f.type }));
  const idMap = Object.fromEntries(fieldMeta.map((f) => [f.id, f.name]));

  const rows = [];
  let pageToken;
  let res;
  do {
    res = await table.getRecords({ pageSize: 500, viewId, pageToken });
    for (const rec of res.records || []) {
      const values = {};
      for (const [fid, val] of Object.entries(rec.fields || {})) {
        const nm = idMap[fid];
        if (nm) values[nm] = val;
      }
      rows.push({ recordId: rec.recordId, values });
    }
    pageToken = res.pageToken;
  } while (res.hasMore);

  state.source = {
    tableName: table.name || '(当前表)',
    viewName: view?.name || '(当前视图)',
    fields: fieldMeta,
    rows,
  };
  renderSource();
  buildKeyOptions();
  log(`已读取视图「${state.source.viewName}」：${rows.length} 条记录 / ${fieldMeta.length} 个字段`);
}

function renderSource() {
  const s = state.source;
  $('srcBody').innerHTML = `
    <div class="meta-line">数据表：<b>${escapeHtml(s.tableName)}</b> · 视图：<b>${escapeHtml(s.viewName)}</b></div>
    <div class="meta-line">记录 <b>${s.rows.length}</b> 条 · 字段 <b>${s.fields.length}</b> 个</div>
    <div class="chips">${s.fields.map((f) => `<span class="chip">${escapeHtml(f.name)}</span>`).join('')}</div>`;
}

// ---------------- 目标 schema ----------------
async function ensureTargetFields() {
  if (state.target.mode === 'same') {
    const id = $('sameTable').value;
    if (!id) throw new Error('请选择目标数据表');
    const t = await state.bitable.base.getTableById(id);
    const fs = (await t.getFields()) || [];
    state.targetFields = fs.map((f) => ({ id: f.id, name: f.name }));
  } else if (state.target.type === 'bitable') {
    const app = $('tgtAppToken').value.trim();
    const tbl = $('tgtTableId').value.trim();
    if (!app || !tbl) throw new Error('请填写目标多维表 App Token 与 Table ID');
    const fs = [];
    let pageToken;
    let d;
    do {
      const params = { page_size: '100' };
      if (pageToken) params.page_token = pageToken;
      d = await api('GET', `/bitable/v1/apps/${app}/tables/${tbl}/fields`, { params });
      const items = d.items || (Array.isArray(d) ? d : []);
      for (const it of items) fs.push({ id: it.field_id || it.id, name: it.field_name || it.name });
      pageToken = d.page_token;
    } while (d.has_more);
    state.targetFields = fs;
  } else {
    // 电子表格：目标为列字母
    state.targetFields = genColumns(52).map((c) => ({ id: c, name: c }));
  }
  return state.targetFields;
}

// ---------------- 字段映射 UI ----------------
function buildMappingUI() {
  const list = $('mapList');
  list.innerHTML = '';
  const tgtIsSheet = state.target.mode === 'cross' && state.target.type === 'sheet';
  const dlId = 'tgtList';
  // 复用同一个 datalist，避免重复追加
  let dl = $(dlId);
  if (!dl) { dl = document.createElement('datalist'); dl.id = dlId; document.body.appendChild(dl); }
  dl.innerHTML = state.targetFields.map((f) => `<option value="${escapeHtml(f.name)}">`).join('');
  const srcOptions = state.source
    ? state.source.fields.map((f) => `<option value="${escapeHtml(f.name)}">${escapeHtml(f.name)}</option>`).join('')
    : '';
  state.mapping.forEach((m, i) => {
    const row = document.createElement('div');
    row.className = 'map-row';
    row.innerHTML = `
      <select class="src-sel" data-i="${i}">${srcOptions}</select>
      <span class="arrow">→</span>
      <input class="tgt" data-i="${i}" list="${dlId}" placeholder="${tgtIsSheet ? '列如 B' : '目标字段名'}" value="${escapeHtml(m.tgt)}" />
      <button class="rm" data-i="${i}" title="移除">×</button>`;
    list.appendChild(row);
  });
  list.querySelectorAll('.src-sel').forEach((sel) => sel.addEventListener('change', (e) => {
    state.mapping[+e.target.dataset.i].src = e.target.value;
  }));
  list.querySelectorAll('.tgt').forEach((inp) => inp.addEventListener('input', (e) => {
    state.mapping[+e.target.dataset.i].tgt = e.target.value.trim();
  }));
  list.querySelectorAll('.rm').forEach((b) => b.addEventListener('click', (e) => {
    state.mapping.splice(+e.target.dataset.i, 1);
    buildMappingUI();
  }));
}

function addMap() {
  if (!state.source) return;
  state.mapping.push({ src: state.source.fields[0]?.name || '', tgt: '' });
  buildMappingUI();
}

async function autoMatch() {
  try {
    await ensureTargetFields();
  } catch (e) {
    log('自动匹配：未能读取目标字段，仅按名称猜测（' + e.message + '）', 'err');
  }
  const names = new Set(state.targetFields.map((f) => f.name));
  state.mapping = state.source.fields
    .filter((f) => !names.size || names.has(f.name))
    .map((f) => ({ src: f.name, tgt: names.has(f.name) ? f.name : '' }));
  buildMappingUI();
  log(`自动匹配完成：${state.mapping.filter((m) => m.tgt).length}/${state.mapping.length} 个字段已对应目标`);
}

function buildKeyOptions() {
  const sel = $('keyField');
  sel.innerHTML = '<option value="">（不使用主键 / 全量按追加）</option>' +
    state.source.fields.map((f) => `<option value="${escapeHtml(f.name)}">${escapeHtml(f.name)}</option>`).join('');
  if (!state.keyField && state.source.fields.length) {
    state.keyField = state.source.fields[0].name;
    sel.value = state.keyField;
  }
  sel.onchange = () => (state.keyField = sel.value);
}

// ---------------- 同步引擎 ----------------
function buildPayload() {
  const active = state.mapping.filter((m) => m.tgt);
  if (!active.length) throw new Error('请先建立至少一个字段映射');
  // 解析同文件模式下的 目标字段名→id
  const nameToId = Object.fromEntries(state.targetFields.map((f) => [f.name, f.id]));
  const isSheet = state.target.mode === 'cross' && state.target.type === 'sheet';
  const isSame = state.target.mode === 'same';

  return state.source.rows.map((row) => {
    const fields = {};
    for (const m of active) {
      const v = row.values[m.src];
      if (v === undefined || v === null || v === '') continue;
      if (isSheet) fields[m.tgt] = toSheetValue(v);
      else if (isSame) fields[nameToId[m.tgt]] = v;       // SDK 用字段 id
      else fields[m.tgt] = v;                              // 跨多维表用字段名
    }
    return { recordId: row.recordId, fields, key: row.values[state.keyField] };
  }).filter((p) => Object.keys(p.fields).length > 0);
}

function toSheetValue(v) {
  if (v == null) return '';
  if (typeof v === 'object') {
    if (Array.isArray(v)) return v.map((x) => (x && x.text) ? x.text : (typeof x === 'object' ? JSON.stringify(x) : x)).join(', ');
    if (v.text != null) return v.text;
    return JSON.stringify(v);
  }
  return v;
}

async function runSync() {
  const btn = $('runBtn');
  btn.disabled = true;
  $('result').classList.add('hidden');
  try {
    const payload = buildPayload();
    const total = payload.length;
    if (!total) throw new Error('没有可同步的数据（检查映射与目标字段）');
    $('progressWrap').classList.remove('hidden');
    setProgress(0);
    log(`开始同步：模式=${state.syncMode === 'full' ? '全量覆盖' : '增量'}, 记录 ${total} 条`);

    if (state.target.mode === 'same') await runSame(payload);
    else if (state.target.type === 'bitable') await runCrossBitable(payload);
    else await runCrossSheet(payload);

    setProgress(100);
    showResult(`✅ 同步完成：新增 ${cnt.add} · 更新 ${cnt.upd} · 删除 ${cnt.del}`);
    log(`同步完成：新增 ${cnt.add} / 更新 ${cnt.upd} / 删除 ${cnt.del}`, 'ok');
  } catch (e) {
    log('同步失败：' + e.message, 'err');
    showResult('❌ 同步失败：' + e.message, true);
    setStatus('同步失败', 'err');
  } finally {
    btn.disabled = false;
  }
}

const cnt = { add: 0, upd: 0, del: 0 };
function setProgress(p) { $('progressBar').style.width = p + '%'; }

// —— 同文件（SDK 直写，免密钥） ——
async function runSame(payload) {
  const t = await state.bitable.base.getTableById($('sameTable').value);
  cnt.add = cnt.upd = cnt.del = 0;
  const delMissing = state.syncMode === 'incr' && $('delMissing').checked;

  if (state.syncMode === 'full') {
    let pt; const exist = []; let r;
    do { r = await t.getRecords({ pageSize: 500, pageToken: pt }); exist.push(...r.records); pt = r.pageToken; } while (r.hasMore);
    for (const rec of exist) { await t.deleteRecord(rec.recordId); await sleep(60); cnt.del++; }
    log(`全量：已清空目标 ${exist.length} 条旧记录`);
  } else {
    const exist = [];
    let pt; let r; do { r = await t.getRecords({ pageSize: 500, pageToken: pt }); exist.push(...r.records); pt = r.pageToken; } while (r.hasMore);
    const idMap = Object.fromEntries(state.targetFields.map((f) => [f.name, f.id]));
    const keyName = state.keyField;
    const keyId = idMap[keyName];
    const tgtMap = new Map();
    for (const r of exist) {
      const kv = r.fields?.[keyId];
      if (kv != null) tgtMap.set(kv, r.recordId);
    }
    const srcKeys = new Set(payload.map((p) => p.key));
    for (const p of payload) {
      const rid = tgtMap.get(p.key);
      if (rid) { await t.setRecord({ recordId: rid, fields: p.fields }); cnt.upd++; }
      else { await t.addRecord({ fields: p.fields }); cnt.add++; }
      await sleep(60);
    }
    if (delMissing) {
      for (const [kv, rid] of tgtMap) if (!srcKeys.has(kv)) { await t.deleteRecord(rid); cnt.del++; await sleep(60); }
    }
  }

  if (state.syncMode === 'full') {
    for (const p of payload) { await t.addRecord({ fields: p.fields }); cnt.add++; await sleep(60); }
  }
}

// —— 跨多维表（开放 API） ——
async function runCrossBitable(payload) {
  const app = $('tgtAppToken').value.trim();
  const tbl = $('tgtTableId').value.trim();
  cnt.add = cnt.upd = cnt.del = 0;
  const delMissing = state.syncMode === 'incr' && $('delMissing').checked;

  const readAll = async () => {
    const out = []; let pt; let d;
    do {
      const params = { page_size: '500', automatic_fields: 'true' };
      if (pt) params.page_token = pt;
      d = await api('GET', `/bitable/v1/apps/${app}/tables/${tbl}/records`, { params, throttle: true });
      for (const it of (d.items || [])) out.push({ recordId: it.record_id, fields: it.fields || {} });
      pt = d.page_token;
    } while (d.has_more);
    return out;
  };

  const appendBatch = async (recs) => {
    if (!recs.length) return;
    await api('POST', `/bitable/v1/apps/${app}/tables/${tbl}/records`,
      { body: { records: recs.map((f) => ({ fields: f })) }, throttle: true });
  };
  const updateBatch = async (recs) => {
    if (!recs.length) return;
    await api('POST', `/bitable/v1/apps/${app}/tables/${tbl}/records/batch_update`,
      { body: { records: recs.map((r) => ({ record_id: r.rid, fields: r.fields })) }, throttle: true });
  };
  const deleteBatch = async (rids) => {
    if (!rids.length) return;
    await api('POST', `/bitable/v1/apps/${app}/tables/${tbl}/records/batch_delete`,
      { body: { records: rids.map((rid) => ({ record_id: rid })) }, throttle: true });
  };

  if (state.syncMode === 'full') {
    const exist = await readAll();
    for (let i = 0; i < exist.length; i += 100) { await deleteBatch(exist.slice(i, i + 100).map((r) => r.recordId)); setProgress(Math.round((i / Math.max(exist.length, 1)) * 50)); }
    cnt.del = exist.length;
    log(`全量：已删除目标 ${exist.length} 条旧记录`);
    for (let i = 0; i < payload.length; i += 100) { await appendBatch(payload.slice(i, i + 100).map((p) => p.fields)); cnt.add += payload.slice(i, i + 100).length; setProgress(50 + Math.round((i / Math.max(payload.length, 1)) * 50)); }
  } else {
    const exist = await readAll();
    const keyName = state.keyField;
    const tgtMap = new Map();
    for (const r of exist) { const kv = r.fields?.[keyName]; if (kv != null) tgtMap.set(kv, r.recordId); }
    const srcKeys = new Set(payload.map((p) => p.key));
    const toUpd = [], toAdd = [];
    for (const p of payload) {
      const rid = tgtMap.get(p.key);
      if (rid) toUpd.push({ rid, fields: p.fields }); else toAdd.push(p.fields);
    }
    for (let i = 0; i < toUpd.length; i += 100) await updateBatch(toUpd.slice(i, i + 100));
    for (let i = 0; i < toAdd.length; i += 100) await appendBatch(toAdd.slice(i, i + 100));
    cnt.upd = toUpd.length; cnt.add = toAdd.length;
    if (delMissing) {
      const delRids = [...tgtMap.entries()].filter(([kv]) => !srcKeys.has(kv)).map(([, rid]) => rid);
      for (let i = 0; i < delRids.length; i += 100) await deleteBatch(delRids.slice(i, i + 100));
      cnt.del = delRids.length;
    }
  }
}

// —— 跨电子表格（开放 API） ——
async function runCrossSheet(payload) {
  const token = $('tgtSheetToken').value.trim();
  const sheetId = $('tgtSheetId').value.trim() || '0';
  cnt.add = cnt.upd = cnt.del = 0;
  const delMissing = state.syncMode === 'incr' && $('delMissing').checked;

  // mapping 顺序即写入列顺序；表头用「源字段名」(可读)，数据按目标列字母定位
  const map = state.mapping.filter((m) => m.tgt);
  const cols = map.map((m) => m.tgt);                 // 目标列字母，按映射顺序
  const firstCol = minColLetter(cols);                // 最左列
  const lastColL = maxColLetter(cols);                // 最右列
  const header = map.map((m) => m.src);               // 源字段名作为表头
  const dataRows = payload.map((p) => cols.map((c) => p.fields[c] ?? ''));
  const total = payload.length;

  const write = async (values, startRow = 1) => {
    const range = `${sheetId}!${firstCol}${startRow}:${lastColL}${startRow + values.length - 1}`;
    await api('POST', `/sheets/v2/spreadsheets/${token}/values_batch_update`,
      { body: { valueRanges: [{ range, values }] }, throttle: true });
  };

  if (state.syncMode === 'full') {
    await write([header, ...dataRows]);
    cnt.add = dataRows.length;
    log(`全量：已写入电子表格 ${dataRows.length} 行（含表头）`);
    return;
  }

  // 增量：读取现有表头与数据，按主键（源字段名）所在列匹配
  const readRange = await api('GET', `/sheets/v2/spreadsheets/${token}/values/${sheetId}!${firstCol}1:${lastColL}1000`);
  const grid = (readRange?.valueRange?.values) || [];
  const existingHeader = grid[0] || [];
  const keyColIdx = existingHeader.indexOf(state.keyField); // 源字段名在表头中的位置
  const tgtMap = new Map();
  for (let r = 1; r < grid.length; r++) {
    const kv = grid[r][keyColIdx];
    if (kv != null && kv !== '') tgtMap.set(String(kv), r + 1); // 行号（1-based）
  }
  const srcKeys = new Set(payload.map((p) => String(p.key)));
  const updates = [];   // 已存在行：原地更新
  const appends = [];   // 新增行：统一追加到末尾
  payload.forEach((p, idx) => {
    const rownum = tgtMap.get(String(p.key));
    const rowVals = cols.map((c) => p.fields[c] ?? '');
    if (rownum) { updates.push({ range: `${sheetId}!${firstCol}${rownum}:${lastColL}${rownum}`, values: [rowVals] }); cnt.upd++; }
    else { appends.push(rowVals); cnt.add++; }
    if (idx % 20 === 0) setProgress(Math.round((idx / total) * 80));
  });
  // 批量更新（每批最多 100 个 range）
  for (let i = 0; i < updates.length; i += 100) {
    await api('POST', `/sheets/v2/spreadsheets/${token}/values_batch_update`,
      { body: { valueRanges: updates.slice(i, i + 100) }, throttle: true });
  }
  // 追加新增行到现有数据之后的空行
  if (appends.length) await write(appends, grid.length + 1);
  if (delMissing) log('电子表格「删除缺失」未执行（避免破坏行结构，仅更新/追加）', 'err');
  log(`增量：更新 ${cnt.upd} / 新增 ${cnt.add}`);
}

// ---------------- 工具 ----------------
function genColumns(n) {
  const a = 'A'.charCodeAt(0); const out = [];
  for (let i = 0; i < n; i++) {
    if (i < 26) out.push(String.fromCharCode(a + i));
    else out.push('A' + String.fromCharCode(a + (i - 26)));
  }
  return out;
}
function lastCol(n) { return genColumns(n).pop(); }
// 列字母 ⇄ 序号（支持 AA、AB…）
function colToNum(col) { let n = 0; for (const c of col) n = n * 26 + (c.charCodeAt(0) - 64); return n; }
function numToCol(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
function minColLetter(cols) { return numToCol(cols.reduce((m, c) => Math.min(m, colToNum(c)), Infinity)); }
function maxColLetter(cols) { return numToCol(cols.reduce((m, c) => Math.max(m, colToNum(c)), 0)); }
function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// ---------------- UI 事件绑定 ----------------
function bindUI() {
  $('modeSeg').addEventListener('click', (e) => {
    const b = e.target.closest('.seg-btn'); if (!b) return;
    state.target.mode = b.dataset.mode;
    [...e.currentTarget.children].forEach((x) => x.classList.toggle('active', x === b));
    $('sameBox').classList.toggle('hidden', state.target.mode !== 'same');
    $('crossBox').classList.toggle('hidden', state.target.mode !== 'cross');
    refreshTargetFieldsUI();
  });
  $('typeSeg').addEventListener('click', (e) => {
    const b = e.target.closest('.seg-btn'); if (!b) return;
    state.target.type = b.dataset.type;
    [...e.currentTarget.children].forEach((x) => x.classList.toggle('active', x === b));
    $('bitableInputs').classList.toggle('hidden', state.target.type !== 'bitable');
    $('sheetInputs').classList.toggle('hidden', state.target.type !== 'sheet');
    refreshTargetFieldsUI();
  });
  $('modeSyncSeg').addEventListener('click', (e) => {
    const b = e.target.closest('.seg-btn'); if (!b) return;
    state.syncMode = b.dataset.sync;
    [...e.currentTarget.children].forEach((x) => x.classList.toggle('active', x === b));
  });
  $('autoMap').addEventListener('click', autoMatch);
  $('addMap').addEventListener('click', addMap);
  $('runBtn').addEventListener('click', runSync);

  // 同文件：选表即加载目标字段
  $('sameTable').addEventListener('change', refreshTargetFieldsUI);
  // 跨文件：输入框失焦时尝试预载目标字段（用于自动匹配建议）
  ['tgtAppToken', 'tgtTableId', 'tgtSheetToken', 'tgtSheetId'].forEach((id) =>
    $(id).addEventListener('blur', () => { if (state.target.mode === 'cross') refreshTargetFieldsUI(); }));
}

async function refreshTargetFieldsUI() {
  try {
    if (state.target.mode === 'same') {
      // 填充本文件表清单（排除当前表可选，这里列出全部）
      const list = await state.bitable.base.getTableList();
      const tables = list.tableList || list.tables || (Array.isArray(list) ? list : []);
      $('sameTable').innerHTML = tables.map((t) => `<option value="${t.id}">${escapeHtml(t.name)}</option>`).join('');
      await ensureTargetFields();
    } else if (state.target.mode === 'cross' && state.target.type === 'sheet') {
      await ensureTargetFields();
    } /* 跨多维表字段在“自动匹配”时按填写的 token 加载 */
    buildMappingUI();
  } catch (e) {
    log('加载目标字段失败：' + e.message, 'err');
  }
}

// ---------------- 启动 ----------------
async function init() {
  bindUI();
  try {
    const mod = await import('https://cdn.jsdelivr.net/npm/@lark-base-open/js-sdk/+esm');
    state.bitable = mod.bitable;
  } catch (e) {
    setStatus('未在飞书环境', 'err');
    $('srcBody').innerHTML = '<p class="muted">本插件需在飞书多维表中以「自定义插件」方式打开（粘贴本页网址）。单独打开无法连接多维表。</p>';
    $('runBtn').disabled = true;
    return;
  }
  try {
    await state.bitable.base;          // 等待插件上下文就绪
    await loadSource();
    await refreshTargetFieldsUI();
    setStatus('已连接', 'ok');
    $('runBtn').disabled = false;
  } catch (e) {
    setStatus('连接失败', 'err');
    log('初始化失败：' + e.message, 'err');
  }
}

init();
