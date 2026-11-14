/**
 * 冰箱溫度記錄表系統 — Google 雲端同步後端（Apps Script，綁定在試算表上）
 *
 * 試算表分頁：設備、溫度紀錄（每天每時段每台一列）、設定。可直接在試算表查看；手動修改也會同步回 App。
 * 每列有「更新時間」(毫秒) 與「版本」：App 以更新時間決定新舊，以版本只下載有變動的列。
 * 密碼：第一次從 App 連線（ping）時設定，存成 SHA-256 雜湊在指令碼屬性 PW_HASH。
 */

const TABLES = {
  fridges: {
    sheet: '設備',
    head: ['ID', '名稱', '類型', '下限', '上限', '停用', '已刪除', '排序', '更新時間', '版本'],
    text: [1, 2, 3],
    key: (r) => r.id,
    valid: (r) => !!r.id,
    toRow: (r, rev) => [r.id, str_(r.name), TYPE_N[r.type] || '其他', nn_(r.lo), nn_(r.hi), r.off ? '是' : '', r.deleted ? '是' : '', num_(r.ord), num_(r.at), rev],
    fromRow: (a) => ({ id: str_(a[0]), name: str_(a[1]), type: TYPE_K[str_(a[2])] || 'other', lo: nul_(a[3]), hi: nul_(a[4]), off: yes_(a[5]), deleted: yes_(a[6]), ord: num_(a[7]), at: num_(a[8]) }),
  },
  recs: {
    sheet: '溫度紀錄',
    head: ['日期', '時段ID', '時段', '設備ID', '設備名稱', '溫度', '狀態', '記錄時間', '記錄人', '異常處置', '複測溫度', '更新時間', '版本'],
    text: [1, 2, 3, 4, 5, 8, 9, 10],
    key: (r) => r.d + '|' + r.sid + '|' + r.fid,
    valid: (r) => /^\d{4}-\d{2}-\d{2}$/.test(r.d) && !!r.fid && !!r.sid,
    toRow: (r, rev, names) => {
      const t = nul_(r.t);
      return [r.d, r.sid, names.slot[r.sid] || '', r.fid, names.fridge[r.fid] || '', t == null ? '' : t, t == null ? '' : r.x ? '異常' : '正常',
        str_(r.ts), str_(r.by), str_(r.act), nn_(r.rt), num_(r.at), rev];
    },
    fromRow: (a) => ({ d: ymd_(a[0]), sid: str_(a[1]), fid: str_(a[3]), t: nul_(a[5]), x: str_(a[6]) === '異常' ? 1 : 0, ts: str_(a[7]), by: str_(a[8]), act: str_(a[9]), rt: nul_(a[10]), at: num_(a[11]) }),
  },
};
const ORDER = ['fridges', 'recs'];
const CONF_KEYS = ['shop', 'slots', 'staff'];
const TYPE_N = { chill: '冷藏', freeze: '冷凍', hot: '熱藏', other: '其他' };
const TYPE_K = { 冷藏: 'chill', 冷凍: 'freeze', 熱藏: 'hot', 其他: 'other' };

function str_(v) { return v == null ? '' : String(v); }
function num_(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function nul_(v) { if (v === '' || v == null) return null; const n = Number(v); return isFinite(n) ? n : null; }
function nn_(v) { const n = nul_(v); return n == null ? '' : n; }
function yes_(v) { return v === true || v === '是' || v === 'TRUE' || v === 'true' || v === 1; }
function ymd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = str_(v).trim().replace(/\//g, '-');
  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : s;
}
function props_() { return PropertiesService.getScriptProperties(); }
function rev_() { return num_(props_().getProperty('REV')); }
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'fridgetemp:' + s, Utilities.Charset.UTF_8)
    .map((b) => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet_(t) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(t.sheet);
  if (!sh) {
    sh = ss.insertSheet(t.sheet);
    sh.getRange(1, 1, 1, t.head.length).setValues([t.head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function read_(t) {
  const sh = sheet_(t), n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, t.head.length).getValues().filter((a) => str_(a[0]) !== '') : [];
}
function write_(t, rows) {
  const sh = sheet_(t), w = t.head.length, last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, w).clearContent();
  if (!rows.length) return;
  // 文字欄先設成純文字，避免 = + 開頭被當公式、日期被轉成日期格式
  t.text.forEach((c) => sh.getRange(2, c, rows.length, 1).setNumberFormat('@'));
  sh.getRange(2, 1, rows.length, w).setValues(rows);
}

// 設定（店名、時段、記錄人）：整份覆蓋，存在指令碼屬性，並寫一份到「設定」分頁方便查看
function conf_() {
  const p = props_(), out = {};
  CONF_KEYS.forEach((k) => { try { out[k] = JSON.parse(p.getProperty('CONF_' + k) || 'null'); } catch (e) { out[k] = null; } });
  return out;
}
function writeConfSheet_(conf) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('設定');
  if (!sh) sh = ss.insertSheet('設定');
  const slots = conf.slots && Array.isArray(conf.slots.v) ? conf.slots.v.filter((s) => s && !s.deleted) : [];
  const rows = [['項目', '內容'], ['店名', conf.shop ? str_(conf.shop.v) : ''],
    ['記錄時段', slots.map((s) => s.name + (s.time ? ' ' + s.time : '')).join('、')],
    ['記錄人', conf.staff && Array.isArray(conf.staff.v) ? conf.staff.v.join('、') : ''],
    ['說明', '這個分頁只供查看，請在 App 的「設定」修改。']];
  const last = sh.getLastRow();
  if (last > 0) sh.getRange(1, 1, last, 2).clearContent();
  sh.getRange(1, 2, rows.length, 1).setNumberFormat('@');
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
}

function doGet() {
  return ContentService.createTextOutput('冰箱溫度記錄表系統 雲端服務運作中');
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, err: '要求格式錯誤' }); }
  const p = props_();
  const stored = p.getProperty('PW_HASH');
  const pw = str_(req.pw);
  if (!pw) return json_({ ok: false, err: '請輸入密碼' });
  let claimed = false;
  if (!stored) {
    if (req.a !== 'ping') return json_({ ok: false, err: '雲端尚未設定密碼，請重新連線' });
    if (pw.length < 4) return json_({ ok: false, err: '密碼至少 4 個字' });
    p.setProperty('PW_HASH', hash_(pw));
    claimed = true;
  } else if (hash_(pw) !== stored) {
    return json_({ ok: false, err: '密碼錯誤', auth: true });
  }

  if (req.a === 'ping') {
    return json_({
      ok: true, claimed: claimed, rev: rev_(), sheet: SpreadsheetApp.getActive().getUrl(),
      counts: { fridges: read_(TABLES.fridges).filter((a) => !yes_(a[6])).length, recs: read_(TABLES.recs).filter((a) => str_(a[5]) !== '').length },
    });
  }
  if (req.a === 'rev') return json_({ ok: true, rev: rev_() });
  if (req.a === 'setpw') {
    const np = str_(req.newpw);
    if (np.length < 4) return json_({ ok: false, err: '新密碼至少 4 個字' });
    p.setProperty('PW_HASH', hash_(np));
    return json_({ ok: true });
  }
  if (req.a === 'sync') {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try { return json_(sync_(req)); } finally { lock.releaseLock(); }
  }
  return json_({ ok: false, err: '未知的動作' });
}

function names_(fridgeRows, conf) {
  const n = { fridge: {}, slot: {}, ord: {}, slotOrd: {} };
  fridgeRows.forEach((a) => { n.fridge[str_(a[0])] = str_(a[1]); n.ord[str_(a[0])] = num_(a[7]); });
  const slots = conf.slots && Array.isArray(conf.slots.v) ? conf.slots.v : [];
  slots.forEach((s) => { if (s && s.id) { n.slot[s.id] = str_(s.name); n.slotOrd[s.id] = str_(s.time) || '99'; } });
  return n;
}

function sync_(req) {
  const p = props_();
  const ch = req.changes || {};
  const cur = rev_(), next = cur + 1;
  const data = {}, changed = {};
  for (const name of ORDER) {
    const t = TABLES[name];
    const rows = read_(t);
    const idx = {};
    rows.forEach((a, i) => { idx[t.key(t.fromRow(a))] = i; });
    data[name] = { rows: rows, idx: idx };
  }
  // 設定先處理，紀錄列才能帶入最新的時段名稱
  const conf = conf_();
  const cc = ch.conf || {};
  CONF_KEYS.forEach((k) => {
    const v = cc[k];
    if (v && num_(v.at) > num_(conf[k] && conf[k].at)) {
      conf[k] = { v: v.v, at: num_(v.at), rev: next };
      p.setProperty('CONF_' + k, JSON.stringify(conf[k]));
      changed.conf = true;
    }
  });
  let names = names_(data.fridges.rows, conf);
  for (const name of ORDER) {
    const t = TABLES[name], D = data[name], atCol = t.head.length - 2;
    for (const r of ch[name] || []) {
      if (!r || !t.valid(r)) continue;
      const k = t.key(r), i = D.idx[k];
      if (i === undefined) { D.idx[k] = D.rows.length; D.rows.push(t.toRow(r, next, names)); changed[name] = true; }
      else if (num_(r.at) > num_(D.rows[i][atCol])) { D.rows[i] = t.toRow(r, next, names); changed[name] = true; }
    }
    if (name === 'fridges' && changed.fridges) names = names_(data.fridges.rows, conf);
  }
  if (changed.fridges || changed.conf) {
    // 設備或時段改名，紀錄分頁的名稱一起更新（只改顯示欄，不動版本）
    data.recs.rows.forEach((a) => {
      const sn = names.slot[str_(a[1])], fn = names.fridge[str_(a[3])];
      if (sn != null && a[2] !== sn) { a[2] = sn; changed.recNames = true; }
      if (fn != null && a[4] !== fn) { a[4] = fn; changed.recNames = true; }
    });
  }
  const any = changed.fridges || changed.recs || changed.conf;
  if (changed.fridges) {
    data.fridges.rows.sort((a, b) => num_(a[7]) - num_(b[7]));
    write_(TABLES.fridges, data.fridges.rows);
  }
  if (changed.recs || changed.recNames) {
    const so = (a) => names.slotOrd[str_(a[1])] || '99';
    data.recs.rows.sort((a, b) => {
      const da = ymd_(a[0]), db = ymd_(b[0]);
      if (da !== db) return da < db ? -1 : 1;
      if (so(a) !== so(b)) return so(a) < so(b) ? -1 : 1;
      return num_(names.ord[str_(a[3])]) - num_(names.ord[str_(b[3])]);
    });
    write_(TABLES.recs, data.recs.rows);
  }
  if (changed.conf) writeConfSheet_(conf);
  const rev = any ? next : cur;
  if (any) p.setProperty('REV', String(rev));

  // 回傳用戶端還沒有的列（版本 > since）；since 比目前版本還大（試算表被重設過）就全部回傳
  let since = num_(req.since);
  if (since > rev) since = 0;
  const out = { ok: true, rev: rev, full: since === 0, sheet: SpreadsheetApp.getActive().getUrl(), conf: {} };
  for (const name of ORDER) {
    const t = TABLES[name], vc = t.head.length - 1;
    out[name] = data[name].rows.filter((a) => num_(a[vc]) > since).map(t.fromRow);
  }
  CONF_KEYS.forEach((k) => { if (conf[k] && num_(conf[k].rev) > since) out.conf[k] = { v: conf[k].v, at: conf[k].at }; });
  return out;
}

/** 在試算表上手動修改時：更新該列的更新時間與版本（溫度紀錄會重新判斷正常／異常），App 下次同步就會拿到 */
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    let t = null;
    for (const name of ORDER) if (TABLES[name].sheet === sh.getName()) t = TABLES[name];
    if (!t) return;
    const r0 = Math.max(2, e.range.getRow()), r1 = e.range.getRow() + e.range.getNumRows() - 1;
    if (r1 < r0) return;
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const rev = rev_() + 1, w = t.head.length, now = Date.now();
      const vals = sh.getRange(r0, 1, r1 - r0 + 1, w).getValues();
      let range = {};
      if (t === TABLES.recs) read_(TABLES.fridges).forEach((a) => { range[str_(a[0])] = { lo: nul_(a[3]), hi: nul_(a[4]) }; });
      vals.forEach((a) => {
        if (str_(a[0]) === '') return;
        if (t === TABLES.recs) {
          a[0] = ymd_(a[0]);
          const v = nul_(a[5]), g = range[str_(a[3])] || {};
          a[6] = v == null ? '' : ((g.lo != null && v < g.lo) || (g.hi != null && v > g.hi)) ? '異常' : '正常';
        }
        a[w - 2] = now; a[w - 1] = rev;
      });
      t.text.forEach((c) => sh.getRange(r0, c, vals.length, 1).setNumberFormat('@'));
      sh.getRange(r0, 1, vals.length, w).setValues(vals);
      props_().setProperty('REV', String(rev));
    } finally { lock.releaseLock(); }
  } catch (err) {
    console.error(err);
  }
}

/** 在編輯器手動執行一次，完成授權並建立分頁 */
function setup() {
  ORDER.forEach((n) => sheet_(TABLES[n]));
  const def = SpreadsheetApp.getActive().getSheetByName('工作表1') || SpreadsheetApp.getActive().getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && SpreadsheetApp.getActive().getSheets().length > 1) SpreadsheetApp.getActive().deleteSheet(def);
  Logger.log('完成，目前版本 ' + rev_() + '，密碼' + (props_().getProperty('PW_HASH') ? '已設定' : '尚未設定（第一次從 App 連線時設定）'));
}

/** 忘記密碼時在編輯器執行：清除密碼，下次從 App 連線時重新設定 */
function resetPassword() {
  props_().deleteProperty('PW_HASH');
  Logger.log('已清除密碼');
}
