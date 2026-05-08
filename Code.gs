/**
 * SONG PWA — Google Apps Script 後端 v1.0.0
 * 部署：「以網頁應用程式部署」→ 執行身分=自己 / 存取權=任何人
 *
 * 設定：
 *   1. 將 SHEET_ID 替換為你的 Google 試算表 ID
 *   2. 試算表會在第一次呼叫時自動建立兩個分頁
 */
const SHEET_ID = 'PUT_YOUR_SHEET_ID_HERE';
const SONGS_SHEET = '歌單';
const SONGS_HEADER = ['ID', '歌名', '歌手', 'YouTubeID', '縮圖', '時長', '加入時間', '標籤', '播放次數', '最後播放', '歌詞LRC', '備註'];
const SETTINGS_SHEET = '設定';
const SETTINGS_HEADER = ['key', 'value'];
const TZ = 'Asia/Taipei';

// ============ Dispatch ============

function doGet(e) {
  return _dispatch_(e, 'GET');
}

function doPost(e) {
  return _dispatch_(e, 'POST');
}

function _dispatch_(e, method) {
  const params = e && e.parameter ? e.parameter : {};
  let body = {};
  if (method === 'POST' && e.postData && e.postData.contents) {
    try { body = JSON.parse(e.postData.contents); } catch (_) {}
  }
  const action = params.action || body.action || '';
  const callback = params.callback || '';

  let result;
  try {
    switch (action) {
      case 'ping':         result = { ok: true, time: new Date().toISOString() }; break;
      case 'getSongs':     result = getSongs(); break;
      case 'addSong':      result = addSong(body); break;
      case 'updateSong':   result = updateSong(body); break;
      case 'deleteSong':   result = deleteSong(body); break;
      case 'saveLyrics':   result = saveLyrics(body); break;
      case 'incrementPlay':result = incrementPlay(body); break;
      case 'getSetting':   result = getSetting(params.key || body.key); break;
      case 'setSetting':   result = setSetting(body); break;
      default: result = { ok: false, error: 'unknown action: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: String(err && err.message || err) };
  }

  const json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ============ Sheet helpers ============

function _ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function _ensureSheet_(name, header) {
  const ss = _ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.setFrozenRows(1);
  } else {
    const cur = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), header.length)).getValues()[0];
    let needFix = false;
    for (let i = 0; i < header.length; i++) if (cur[i] !== header[i]) { needFix = true; break; }
    if (needFix) sh.getRange(1, 1, 1, header.length).setValues([header]);
  }
  return sh;
}

function _idx_(header) {
  const m = {};
  header.forEach((h, i) => m[h] = i);
  return m;
}

function _readAll_(sh) {
  const lr = sh.getLastRow();
  if (lr < 2) return { header: sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0], rows: [] };
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rows = sh.getRange(2, 1, lr - 1, header.length).getValues();
  return { header, rows };
}

function _rowToObj_(row, idx) {
  const o = {};
  Object.keys(idx).forEach(k => { o[k] = row[idx[k]]; });
  return o;
}

function _objToRow_(obj, header) {
  return header.map(h => obj[h] !== undefined ? obj[h] : '');
}

function _newId_() {
  return 's' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function _normDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss');
  return String(v);
}

// ============ Songs ============

function getSongs() {
  const sh = _ensureSheet_(SONGS_SHEET, SONGS_HEADER);
  const { header, rows } = _readAll_(sh);
  const idx = _idx_(header);
  const songs = rows.map(r => {
    const o = _rowToObj_(r, idx);
    o['加入時間'] = _normDate_(o['加入時間']);
    o['最後播放'] = _normDate_(o['最後播放']);
    return o;
  }).filter(o => o['ID']);
  return { ok: true, songs };
}

function addSong(body) {
  if (!body || !body['YouTubeID']) return { ok: false, error: 'YouTubeID required' };
  const sh = _ensureSheet_(SONGS_SHEET, SONGS_HEADER);
  const { header, rows } = _readAll_(sh);
  const idx = _idx_(header);
  const dup = rows.find(r => r[idx['YouTubeID']] === body['YouTubeID']);
  if (dup) return { ok: false, error: 'duplicate', id: dup[idx['ID']] };
  const id = _newId_();
  const obj = {
    'ID': id,
    '歌名': body['歌名'] || '',
    '歌手': body['歌手'] || '',
    'YouTubeID': body['YouTubeID'],
    '縮圖': body['縮圖'] || ('https://i.ytimg.com/vi/' + body['YouTubeID'] + '/mqdefault.jpg'),
    '時長': body['時長'] || '',
    '加入時間': Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'),
    '標籤': body['標籤'] || '',
    '播放次數': 0,
    '最後播放': '',
    '歌詞LRC': body['歌詞LRC'] || '',
    '備註': body['備註'] || ''
  };
  sh.appendRow(_objToRow_(obj, header));
  return { ok: true, id, song: obj };
}

function _findRow_(sh, idx, id) {
  const lr = sh.getLastRow();
  if (lr < 2) return -1;
  const ids = sh.getRange(2, idx['ID'] + 1, lr - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}

function updateSong(body) {
  if (!body || !body['ID']) return { ok: false, error: 'ID required' };
  const sh = _ensureSheet_(SONGS_SHEET, SONGS_HEADER);
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = _idx_(header);
  const row = _findRow_(sh, idx, body['ID']);
  if (row < 0) return { ok: false, error: 'not found' };
  Object.keys(body).forEach(k => {
    if (k === 'ID' || k === 'action') return;
    if (idx[k] !== undefined) sh.getRange(row, idx[k] + 1).setValue(body[k]);
  });
  return { ok: true };
}

function deleteSong(body) {
  if (!body || !body['ID']) return { ok: false, error: 'ID required' };
  const sh = _ensureSheet_(SONGS_SHEET, SONGS_HEADER);
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = _idx_(header);
  const row = _findRow_(sh, idx, body['ID']);
  if (row < 0) return { ok: false, error: 'not found' };
  sh.deleteRow(row);
  return { ok: true };
}

function saveLyrics(body) {
  if (!body || !body['ID']) return { ok: false, error: 'ID required' };
  return updateSong({ ID: body['ID'], '歌詞LRC': body['歌詞LRC'] || '' });
}

function incrementPlay(body) {
  if (!body || !body['ID']) return { ok: false, error: 'ID required' };
  const sh = _ensureSheet_(SONGS_SHEET, SONGS_HEADER);
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = _idx_(header);
  const row = _findRow_(sh, idx, body['ID']);
  if (row < 0) return { ok: false, error: 'not found' };
  const cur = Number(sh.getRange(row, idx['播放次數'] + 1).getValue()) || 0;
  sh.getRange(row, idx['播放次數'] + 1).setValue(cur + 1);
  sh.getRange(row, idx['最後播放'] + 1).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'));
  return { ok: true, count: cur + 1 };
}

// ============ Settings ============

function getSetting(key) {
  if (!key) return { ok: false, error: 'key required' };
  const sh = _ensureSheet_(SETTINGS_SHEET, SETTINGS_HEADER);
  const { rows } = _readAll_(sh);
  const r = rows.find(x => x[0] === key);
  return { ok: true, key, value: r ? r[1] : '' };
}

function setSetting(body) {
  if (!body || !body.key) return { ok: false, error: 'key required' };
  const sh = _ensureSheet_(SETTINGS_SHEET, SETTINGS_HEADER);
  const { rows } = _readAll_(sh);
  const i = rows.findIndex(x => x[0] === body.key);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(body.value || '');
  else sh.appendRow([body.key, body.value || '']);
  return { ok: true };
}
