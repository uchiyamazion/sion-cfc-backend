/**
 * シオンテクノス株式会社 - フロン排出抑制法一括管理システム API
 * * GitHub & CLASP 管理用メインソースコード
 */

// GETリクエスト：機器情報の取得（AppSheetやマスター参照用）
function doGet(e) {
  const __g = authGuard_(e); if (__g) return __g;
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('機器台帳');
    if (!sheet) {
      return createJsonResponse({ status: 'error', message: '機器台帳シートが見つかりません。' });
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    const result = rows.map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
    
    return createJsonResponse({ status: 'success', data: result });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  }
}

// POSTリクエスト：スマホアプリ等からのデータ書込・登録処理
function doPost(e) {
  const __g = authGuard_(e); if (__g) return __g;
  try {
    if (!e.postData || !e.postData.contents) {
      return createJsonResponse({ status: 'error', message: 'リクエストボディが空です。' });
    }
    
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    const payload = requestData.payload;
    
    let result;
    switch (action) {
      case 'recordCheck': // 簡易点検・法定点検の記録
        result = recordMaintenance(payload);
        break;
      case 'registerMachine': // 新規機器の台帳登録
        result = registerNewMachine(payload);
        break;
      default:
        return createJsonResponse({ status: 'error', message: '無効なアクションです: ' + action });
    }
    
    return createJsonResponse({ status: 'success', data: result });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  }
}

// 点検記録をスプレッドシートに追記する処理
function recordMaintenance(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('点検記録');
  if (!sheet) throw new Error('点検記録シートが見つかりません。');
  
  sheet.appendRow([
    new Date(),                                         // A: 点検日時
    payload.machineId,                                  // B: 機器管理番号
    payload.inspector,                                  // C: 点検者
    payload.checkType || '簡易点検',                     // D: 点検種別
    payload.isNormal ? '異常なし' : '異常あり',           // E: 簡易点検結果
    payload.leakageDetected ? '漏洩あり' : '漏洩なし',    // F: 法定漏洩有無
    payload.memo || '',                                 // G: 備考・所見
    payload.imageUrl || ''                              // H: 現場写真URL
  ]);
  
  // 機器台帳側の最終点検日を自動更新
  updateLastCheckDate(payload.machineId);
  
  return '点検記録を正常に保存しました。';
}

// 新しいフロン機器を台帳に登録する処理
function registerNewMachine(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('機器台帳');
  if (!sheet) throw new Error('機器台帳シートが見つかりません。');
  
  sheet.appendRow([
    payload.machineId,        // A: 機器管理番号
    payload.facilityName,     // B: 施設名
    payload.location,         // C: 設置場所
    payload.manufacturer,     // D: メーカー
    payload.modelName,        // E: 型式
    payload.serialNumber,     // F: 製造番号
    payload.gasType,          // G: フロン類の種類
    payload.capacity,         // H: 初期充填量(kg)
    new Date()                // I: 登録・最終点検日
  ]);
  
  return '新規機器を台帳に登録しました。';
}

// 機器台帳の最終点検日列(I列)を更新する内部関数
function updateLastCheckDate(machineId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('機器台帳');
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == machineId) {
      sheet.getRange(i + 1, 9).setValue(new Date()); // 9列目 = I列
      break;
    }
  }
}

// JSON出力を生成する共通ヘルパー
function createJsonResponse(outputObject) {
  return ContentService.createTextOutput(JSON.stringify(outputObject))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════
// AuthGuard.gs ― Googleログインの本人確認（GAS側）  v2
//
// ■ 許可リスト
//   ・共通の「アクセス許可」スプレッドシート（シート名「許可アカウント」）を、管理画面から編集します。
//     スクリプト プロパティ ACL_SHEET_ID にそのスプレッドシートのIDを登録してください。
//   ・予備として、スクリプト プロパティ ALLOWED_EMAILS（カンマ区切り）も有効です。
//   ・管理者：シートの「権限」列が「管理者」の人、またはスクリプト プロパティ ADMIN_EMAILS の人。
//   ・許可リストが空のときは全員を拒否します。
//
// ■ 初回だけ：エディタで関数 authorizeOnce を選んで「実行」し、権限を承認してください。
//   （名前の末尾が _ の関数は、実行メニューに表示されません）
//
// ■ doGet(e) と doPost(e) の先頭に、次の2行を入れる
//   const g = authGuard_(e); if (g) return g;
// ════════════════════════════════════════════════

const GOOGLE_CLIENT_ID = '782190614995-5d4fur401smd1gpnnvs43h3m2q2q2ar3.apps.googleusercontent.com';
const ACL_SHEET_NAME = '許可アカウント';

// 初回だけ実行する関数：外部通信とスプレッドシートの権限の承認画面を出す
function authorizeOnce() {
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=test', { muteHttpExceptions: true });
  CacheService.getScriptCache().put('authorize_test', '1', 5);
  const id = PropertiesService.getScriptProperties().getProperty('ACL_SHEET_ID');
  if (id) { const sh = aclSheet_(); Logger.log('許可シート: ' + sh.getParent().getName() + ' / 登録 ' + Math.max(0, sh.getLastRow() - 1) + ' 件'); }
  Logger.log('承認OK。ACL_SHEET_ID=' + (id ? '設定あり' : '未設定（ALLOWED_EMAILSだけで動作）'));
}

// ---------- 許可リスト ----------
function listProp_(name) {
  const v = PropertiesService.getScriptProperties().getProperty(name) || '';
  return v.split(/[\s,;]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
}
function aclSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('ACL_SHEET_ID');
  if (!id) return null;
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(ACL_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ACL_SHEET_NAME);
    sh.getRange(1, 1, 1, 4).setValues([['メールアドレス', '名前', '権限', '登録日時']]);
  }
  return sh;
}
function aclRows_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('acl_rows');
  if (hit) { try { return JSON.parse(hit); } catch (e) { /* 再取得 */ } }
  let rows = [];
  try {
    const sh = aclSheet_();
    if (sh && sh.getLastRow() > 1) {
      rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
        .map(r => ({ email: String(r[0]).trim().toLowerCase(), name: String(r[1] || ''), role: String(r[2] || '') }))
        .filter(r => r.email);
    }
  } catch (err) { rows = []; }
  cache.put('acl_rows', JSON.stringify(rows), 60);
  return rows;
}
function allowedEmails_() { return listProp_('ALLOWED_EMAILS').concat(aclRows_().map(r => r.email)); }
function adminEmails_() { return listProp_('ADMIN_EMAILS').concat(aclRows_().filter(r => r.role === '管理者').map(r => r.email)); }

// ---------- IDトークンの確認 ----------
// 戻り値：{ ok:true, email } | { ok:false, reason:'unauthorized' } | { ok:false, reason:'forbidden', email }
function verifyIdToken_(idToken) {
  if (!idToken) return { ok: false, reason: 'unauthorized' };
  const cache = CacheService.getScriptCache();
  const key = 'idt_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken)).slice(0, 40);
  let info = null;
  const hit = cache.get(key);
  if (hit) { try { info = JSON.parse(hit); } catch (e) { info = null; } }
  if (!info) {
    const res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return { ok: false, reason: 'unauthorized' };
    const r = JSON.parse(res.getContentText());
    info = { aud: r.aud, iss: r.iss, email: r.email, email_verified: r.email_verified, exp: r.exp };
    const ttl = Math.min(300, Number(info.exp) - Math.floor(Date.now() / 1000));
    if (ttl > 0) cache.put(key, JSON.stringify(info), ttl);
  }
  const now = Math.floor(Date.now() / 1000);
  const iss = String(info.iss || '');
  if (info.aud !== GOOGLE_CLIENT_ID) return { ok: false, reason: 'unauthorized' };
  if (iss !== 'accounts.google.com' && iss !== 'https://accounts.google.com') return { ok: false, reason: 'unauthorized' };
  if (Number(info.exp) <= now) return { ok: false, reason: 'unauthorized' };
  if (String(info.email_verified) !== 'true') return { ok: false, reason: 'unauthorized' };
  const email = String(info.email || '').toLowerCase();
  if (allowedEmails_().indexOf(email) < 0) return { ok: false, reason: 'forbidden', email: email };
  return { ok: true, email: email };
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- 管理画面用の操作（acl_list / acl_add / acl_remove）。管理者だけが使える ----------
function aclHandle_(action, p, a) {
  if (!a.email) return jsonOut_({ status: 'error', success: false, error: a.reason, data: a.reason, message: a.reason });
  if (adminEmails_().indexOf(a.email) < 0) {
    return jsonOut_({ status: 'error', success: false, error: 'not_admin', data: 'not_admin', message: '管理者のみ操作できます' });
  }
  const props = { me: a.email };
  if (action === 'acl_list') {
    props.rows = aclRows_();
    props.propEmails = listProp_('ALLOWED_EMAILS');
    props.propAdmins = listProp_('ADMIN_EMAILS');
    props.sheetReady = !!PropertiesService.getScriptProperties().getProperty('ACL_SHEET_ID');
    return jsonOut_({ status: 'success', success: true, data: props });
  }
  const sh = aclSheet_();
  if (!sh) return jsonOut_({ status: 'error', success: false, error: 'no_sheet', data: 'no_sheet', message: 'スクリプト プロパティ ACL_SHEET_ID が未設定です' });
  const email = String(p.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return jsonOut_({ status: 'error', success: false, error: 'bad_email', data: 'bad_email', message: 'メールアドレスの形式が正しくありません' });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const last = sh.getLastRow();
    const emails = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase()) : [];
    const idx = emails.indexOf(email);
    if (action === 'acl_add') {
      const name = String(p.name || '').slice(0, 60);
      const role = p.role === '管理者' ? '管理者' : '';
      if (idx >= 0) sh.getRange(idx + 2, 1, 1, 3).setValues([[email, name, role]]);
      else sh.appendRow([email, name, role, new Date()]);
    } else if (action === 'acl_remove') {
      if (email === a.email) return jsonOut_({ status: 'error', success: false, error: 'self', data: 'self', message: '自分自身は削除できません' });
      if (idx >= 0) sh.deleteRow(idx + 2);
    } else {
      return jsonOut_({ status: 'error', success: false, error: 'bad_action', data: 'bad_action', message: '不明な操作です' });
    }
  } finally { lock.releaseLock(); }
  CacheService.getScriptCache().remove('acl_rows');
  props.rows = aclRows_();
  return jsonOut_({ status: 'success', success: true, data: props });
}

// 許可されていれば null、そうでなければエラー応答（JSON）を返す
function authGuard_(e) {
  const p = (e && e.parameter) || {};
  let token = p.idToken || '';
  if (!token && e && e.postData && e.postData.contents) {
    try { token = (JSON.parse(e.postData.contents) || {}).idToken || ''; } catch (err) { /* JSON以外は無視 */ }
  }
  const a = verifyIdToken_(token);
  if (/^acl_/.test(p.action || '')) return aclHandle_(p.action, p, a);
  if (a.ok) return null;
  return jsonOut_({ status: 'error', success: false, ok: false, error: a.reason, data: a.reason, message: a.reason });
}
