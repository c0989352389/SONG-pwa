/**
 * SONG PWA — Google Apps Script 後端 v2.0.0
 * 部署：以網頁應用程式 → 執行=自己 / 存取=任何人
 *
 * 設定：
 *   1. 將 SHEET_ID 替換為你的試算表 ID
 *   2. 試算表會自動建立 4 個分頁
 */
const SHEET_ID = '1THqrOQNwSGeYrlVvx9j928iEwptESlH-pvdyMxMfLMU';

const SONGS_SHEET = '歌單';
const SONGS_HEADER = ['ID','歌名','歌手','YouTubeID','縮圖','時長','加入時間','標籤','播放次數','最後播放','歌詞LRC','備註','翻譯JSON','解釋JSON','段落JSON','歌單IDs','BPM','歌詞偏移','歌單內排序JSON','語言','性別','專輯','更新時間'];

const PL_SHEET = '歌單分組';
const PL_HEADER = ['ID','名稱','排序','顏色','更新時間'];

const SETTINGS_SHEET = '設定';
const SETTINGS_HEADER = ['key','value'];

const TZ = 'Asia/Taipei';

// ============ Dispatch ============

function doGet(e)  { return _dispatch_(e, 'GET'); }
function doPost(e) { return _dispatch_(e, 'POST'); }

function _dispatch_(e, method) {
  const params = (e && e.parameter) ? e.parameter : {};
  let body = {};
  if (method === 'POST' && e && e.postData && e.postData.contents) {
    try { body = JSON.parse(e.postData.contents); } catch (_) {}
  }
  // params 也可帶 body 欄位（JSONP 用）
  Object.keys(params).forEach(k => { if (body[k] === undefined && k !== 'action' && k !== 'callback') body[k] = params[k]; });
  const action = params.action || body.action || '';
  const callback = params.callback || '';

  let result;
  try {
    switch (action) {
      case 'ping':           result = { ok:true, time: new Date().toISOString() }; break;
      case 'getSongs':       result = getSongs(); break;
      case 'addSong':        // legacy alias
      case 'upsertSong':     result = upsertSong(body); break;
      case 'updateSong':     result = updateSong(body); break;
      case 'deleteSong':     result = deleteSong(body); break;
      case 'saveLyrics':     result = saveLyrics(body); break;
      case 'incrementPlay':  result = incrementPlay(body); break;
      case 'getPlaylists':   result = getPlaylists(); break;
      case 'upsertPlaylist': result = upsertPlaylist(body); break;
      case 'deletePlaylist': result = deletePlaylist(body); break;
      case 'getSetting':     result = getSetting(body.key); break;
      case 'setSetting':     result = setSetting(body); break;
      default: result = { ok:false, error:'unknown action: ' + action };
    }
  } catch (err) {
    result = { ok:false, error: String(err && err.message || err) };
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
    return sh;
  }
  // 確保所有 header 欄位都存在(只追加,不破壞既有資料)
  const lastCol = sh.getLastColumn();
  const cur = sh.getRange(1, 1, 1, Math.max(lastCol, 1)).getValues()[0];
  let needAdd = false;
  for (let i = 0; i < header.length; i++) {
    if (cur[i] !== header[i]) { needAdd = true; break; }
  }
  if (needAdd) {
    // 找出缺少的 header 欄位,追加到末尾
    const existing = new Set(cur.filter(x => x));
    let nextCol = lastCol + 1;
    for (const h of header) {
      if (!existing.has(h)) {
        sh.getRange(1, nextCol).setValue(h);
        nextCol++;
      }
    }
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
  const lc = sh.getLastColumn();
  const header = sh.getRange(1, 1, 1, lc).getValues()[0];
  if (lr < 2) return { header, rows: [] };
  const rows = sh.getRange(2, 1, lr - 1, lc).getValues();
  return { header, rows };
}

function _rowToObj_(row, header) {
  const o = {};
  for (let i = 0; i < header.length; i++) o[header[i]] = row[i];
  return o;
}

function _newId_(p) {
  return (p || 's') + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function _normDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss');
  return String(v);
}

function _findRowById_(sh, id) {
  const lr = sh.getLastRow();
  if (lr < 2) return -1;
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = _idx_(header);
  if (idx['ID'] === undefined) return -1;
  const ids = sh.getRange(2, idx['ID'] + 1, lr - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}

// ============ Songs ============

function getSongs() {
  const sh = _ensureSheet_(SONGS_SHEET, SONGS_HEADER);
  const { header, rows } = _readAll_(sh);
  const songs = rows.map(r => {
    const o = _rowToObj_(r, header);
    o['加入時間'] = _normDate_(o['加入時間']);
    o['最後播放'] = _normDate_(o['最後播放']);
    return o;
  }).filter(o => o['ID']);
  return { ok:true, songs };
}

function upsertSong(body) {
  if (!body) return { ok:false, error:'body required' };
  const sh = _ensureSheet_(SONGS_SHEET, SONGS_HEADER);
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = _idx_(header);
  let id = body['ID'];

  // 反查 YouTubeID 看是否已存在(避免相同影片重複)
  if ((!id || _findRowById_(sh, id) < 0) && body['YouTubeID']) {
    const lr = sh.getLastRow();
    if (lr >= 2 && idx['YouTubeID'] !== undefined) {
      const vals = sh.getRange(2, idx['YouTubeID'] + 1, lr - 1, 1).getValues();
      for (let i = 0; i < vals.length; i++) {
        if (vals[i][0] === body['YouTubeID']) { id = sh.getRange(i + 2, idx['ID'] + 1).getValue(); break; }
      }
    }
  }

  let row = id ? _findRowById_(sh, id) : -1;

  const nowIso = new Date().toISOString();
  if (!id || row < 0) {
    if (!id) id = _newId_();
    const obj = {};
    SONGS_HEADER.forEach(k => obj[k] = body[k] !== undefined ? body[k] : '');
    obj['ID'] = id;
    if (!obj['加入時間']) obj['加入時間'] = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    if (!obj['縮圖'] && obj['YouTubeID']) obj['縮圖'] = 'https://i.ytimg.com/vi/' + obj['YouTubeID'] + '/mqdefault.jpg';
    if (obj['播放次數'] === '') obj['播放次數'] = 0;
    obj['更新時間'] = body['更新時間'] || nowIso;
    sh.appendRow(header.map(h => obj[h] !== undefined ? obj[h] : ''));
    return { ok:true, id, created:true };
  }

  // 版本檢查:若雲端版本比本次推送新 → 拒絕(但仍接受播放次數/最後播放等統計類更新)
  if (idx['更新時間'] !== undefined) {
    const cloudT = sh.getRange(row, idx['更新時間'] + 1).getValue();
    const incomingT = body['更新時間'];
    if (cloudT && incomingT && String(cloudT) > String(incomingT)) {
      // 只允許特定統計欄位更新(播放次數/最後播放)
      const allowed = ['播放次數','最後播放'];
      Object.keys(body).forEach(k => {
        if (allowed.includes(k) && idx[k] !== undefined) sh.getRange(row, idx[k] + 1).setValue(body[k]);
      });
      return { ok:true, id, rejected:true, reason:'older version' };
    }
  }

  Object.keys(body).forEach(k => {
    if (k === 'action' || k === 'callback') return;
    if (idx[k] !== undefined) sh.getRange(row, idx[k] + 1).setValue(body[k]);
  });
  // 設定本次更新時間(若 body 沒帶就用 now)
  if (idx['更新時間'] !== undefined && !body['更新時間']) {
    sh.getRange(row, idx['更新時間'] + 1).setValue(nowIso);
  }
  return { ok:true, id, updated:true };
}

function updateSong(body) { return upsertSong(body); }

function deleteSong(body) {
  if (!body || !body['ID']) return { ok:false, error:'ID required' };
  const sh = _ensureSheet_(SONGS_SHEET, SONGS_HEADER);
  const row = _findRowById_(sh, body['ID']);
  if (row < 0) return { ok:false, error:'not found' };
  sh.deleteRow(row);
  return { ok:true };
}

function saveLyrics(body) {
  if (!body || !body['ID']) return { ok:false, error:'ID required' };
  return upsertSong({ ID: body['ID'], '歌詞LRC': body['歌詞LRC'] || '' });
}

function incrementPlay(body) {
  if (!body || !body['ID']) return { ok:false, error:'ID required' };
  const sh = _ensureSheet_(SONGS_SHEET, SONGS_HEADER);
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = _idx_(header);
  const row = _findRowById_(sh, body['ID']);
  if (row < 0) return { ok:false, error:'not found' };
  const cur = Number(sh.getRange(row, idx['播放次數'] + 1).getValue()) || 0;
  sh.getRange(row, idx['播放次數'] + 1).setValue(cur + 1);
  if (idx['最後播放'] !== undefined) sh.getRange(row, idx['最後播放'] + 1).setValue(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'));
  return { ok:true, count: cur + 1 };
}

// ============ Playlists ============

function getPlaylists() {
  const sh = _ensureSheet_(PL_SHEET, PL_HEADER);
  const { header, rows } = _readAll_(sh);
  const playlists = rows.map(r => _rowToObj_(r, header)).filter(o => o['ID']);
  return { ok:true, playlists };
}

function upsertPlaylist(body) {
  if (!body) return { ok:false, error:'body required' };
  const sh = _ensureSheet_(PL_SHEET, PL_HEADER);
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = _idx_(header);
  let id = body['ID'];
  if (!id) id = _newId_('pl');
  const row = _findRowById_(sh, id);
  const nowIso = new Date().toISOString();
  if (row < 0) {
    const obj = {};
    PL_HEADER.forEach(h => obj[h] = (h === 'ID' ? id : (body[h] !== undefined ? body[h] : '')));
    obj['更新時間'] = body['更新時間'] || nowIso;
    sh.appendRow(header.map(h => obj[h] !== undefined ? obj[h] : ''));
    return { ok:true, id, created:true };
  }
  // 版本檢查
  if (idx['更新時間'] !== undefined) {
    const cloudT = sh.getRange(row, idx['更新時間'] + 1).getValue();
    const incomingT = body['更新時間'];
    if (cloudT && incomingT && String(cloudT) > String(incomingT)) {
      return { ok:true, id, rejected:true, reason:'older version' };
    }
  }
  Object.keys(body).forEach(k => {
    if (k === 'action' || k === 'callback') return;
    if (idx[k] !== undefined) sh.getRange(row, idx[k] + 1).setValue(body[k]);
  });
  if (idx['更新時間'] !== undefined && !body['更新時間']) {
    sh.getRange(row, idx['更新時間'] + 1).setValue(nowIso);
  }
  return { ok:true, id, updated:true };
}

function deletePlaylist(body) {
  if (!body || !body['ID']) return { ok:false, error:'ID required' };
  const sh = _ensureSheet_(PL_SHEET, PL_HEADER);
  const row = _findRowById_(sh, body['ID']);
  if (row < 0) return { ok:false, error:'not found' };
  sh.deleteRow(row);
  return { ok:true };
}

// ============ Settings ============

function getSetting(key) {
  if (!key) return { ok:false, error:'key required' };
  const sh = _ensureSheet_(SETTINGS_SHEET, SETTINGS_HEADER);
  const { rows } = _readAll_(sh);
  const r = rows.find(x => x[0] === key);
  return { ok:true, key, value: r ? r[1] : '' };
}

function setSetting(body) {
  if (!body || !body.key) return { ok:false, error:'key required' };
  const sh = _ensureSheet_(SETTINGS_SHEET, SETTINGS_HEADER);
  const { rows } = _readAll_(sh);
  const i = rows.findIndex(x => x[0] === body.key);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(body.value || '');
  else sh.appendRow([body.key, body.value || '']);
  return { ok:true };
}
