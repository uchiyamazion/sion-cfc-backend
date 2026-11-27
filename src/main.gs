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
// AuthGuard.gs ― Googleログインの本人確認（GAS側）  v3
//
// ■ 許可リスト
//   ・共通の「アクセス許可」スプレッドシート（シート名「許可アカウント」）を、管理画面から編集します。
//     スクリプト プロパティ ACL_SHEET_ID にそのスプレッドシートのIDを登録してください。
//   ・予備として、スクリプト プロパティ ALLOWED_EMAILS（カンマ区切り）も有効です。
//   ・管理者：シートの「権限」列が「管理者」の人、またはスクリプト プロパティ ADMIN_EMAILS の人。
//   ・許可リストが空のときは全員を拒否します。
//   ・許可のない人は、画面の「利用を申請する」から申請できます。申請はシート「申請中」に入り、
//     管理画面で承認すると、許可アカウントに移ります（管理者にはメールで通知されます）。
//
// ■ 初回だけ：エディタで関数 authorizeOnce を選んで「実行」し、権限を承認してください。
//   （名前の末尾が _ の関数は、実行メニューに表示されません）
//
// ■ doGet(e) と doPost(e) の先頭に、次の2行を入れる
//   const g = authGuard_(e); if (g) return g;
// ════════════════════════════════════════════════

const GOOGLE_CLIENT_ID = '782190614995-5d4fur401smd1gpnnvs43h3m2q2q2ar3.apps.googleusercontent.com';
const ACL_SHEET_NAME = '許可アカウント';
const REQ_SHEET_NAME = '申請中';
const ADMIN_PAGE_URL = 'https://uchiyamazion.github.io/sion-portal/access/';

// 初回だけ実行する関数：外部通信とスプレッドシートの権限の承認画面を出す
function authorizeOnce() {
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=test', { muteHttpExceptions: true });
  CacheService.getScriptCache().put('authorize_test', '1', 5);
  MailApp.getRemainingDailyQuota();   // 管理者へのメール通知の権限の承認用
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
function reqSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('ACL_SHEET_ID');
  if (!id) return null;
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(REQ_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(REQ_SHEET_NAME);
    sh.getRange(1, 1, 1, 3).setValues([['メールアドレス', '名前', '申請日時']]);
  }
  return sh;
}
function reqRows_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .map(r => ({ email: String(r[0]).trim().toLowerCase(), name: String(r[1] || ''), at: r[2] instanceof Date ? r[2].getTime() : 0 }))
    .filter(r => r.email);
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

// ---------- 利用申請（許可のない人も、Googleにログインしていれば申請できる） ----------
function aclRequestHandle_(p, a) {
  if (!a.email) return jsonOut_({ status: 'error', success: false, error: a.reason, data: a.reason, message: a.reason });
  if (a.ok) return jsonOut_({ status: 'success', success: true, data: { state: 'already_allowed' } });
  const sh = reqSheet_();
  if (!sh) return jsonOut_({ status: 'error', success: false, error: 'no_sheet', data: 'no_sheet', message: '申請を受け付けられません。管理者に直接連絡してください' });
  const cache = CacheService.getScriptCache();
  const rk = 'req_' + a.email;
  if (cache.get(rk)) return jsonOut_({ status: 'success', success: true, data: { state: 'already_requested' } });
  cache.put(rk, '1', 60);
  const name = String(p.name || '').slice(0, 60);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let isNew = false;
  try {
    const rows = reqRows_(sh);
    if (rows.some(r => r.email === a.email)) return jsonOut_({ status: 'success', success: true, data: { state: 'already_requested' } });
    if (rows.length >= 200) return jsonOut_({ status: 'error', success: false, error: 'full', data: 'full', message: '申請が多すぎます。管理者に直接連絡してください' });
    sh.appendRow([a.email, name, new Date()]);
    isNew = true;
  } finally { lock.releaseLock(); }
  if (isNew) notifyAdmins_(a.email, name);
  return jsonOut_({ status: 'success', success: true, data: { state: 'requested' } });
}
function notifyAdmins_(email, name) {
  try {
    const to = adminEmails_().filter((v, i, arr) => arr.indexOf(v) === i);
    if (!to.length) return;
    MailApp.sendEmail({
      to: to.join(','),
      subject: '【アクセス申請】' + email,
      body: '利用の申請がありました。\n\nメールアドレス：' + email + '\n名前：' + (name || '（未入力）') +
            '\n\n承認・却下は、次の管理画面で行ってください。\n' + ADMIN_PAGE_URL
    });
  } catch (err) { /* メール送信に失敗しても申請自体は有効 */ }
}

// ---------- 管理画面用の操作（acl_list / acl_add / acl_remove / acl_approve / acl_reject）。管理者だけが使える ----------
function aclHandle_(action, p, a) {
  const bad = (code, msg) => jsonOut_({ status: 'error', success: false, error: code, data: code, message: msg });
  if (!a.email) return bad(a.reason, a.reason);
  if (adminEmails_().indexOf(a.email) < 0) return bad('not_admin', '管理者のみ操作できます');
  const props = { me: a.email };
  const sheetId = PropertiesService.getScriptProperties().getProperty('ACL_SHEET_ID');
  const result = () => {
    CacheService.getScriptCache().remove('acl_rows');
    props.rows = aclRows_();
    props.pending = reqRows_(reqSheet_());
    props.propEmails = listProp_('ALLOWED_EMAILS');
    props.propAdmins = listProp_('ADMIN_EMAILS');
    props.sheetReady = !!sheetId;
    return jsonOut_({ status: 'success', success: true, data: props });
  };
  if (action === 'acl_list') {
    if (!sheetId) { props.rows = []; props.pending = []; props.propEmails = listProp_('ALLOWED_EMAILS'); props.propAdmins = listProp_('ADMIN_EMAILS'); props.sheetReady = false; return jsonOut_({ status: 'success', success: true, data: props }); }
    return result();
  }
  const sh = aclSheet_();
  if (!sh) return bad('no_sheet', 'スクリプト プロパティ ACL_SHEET_ID が未設定です');
  const email = String(p.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return bad('bad_email', 'メールアドレスの形式が正しくありません');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const last = sh.getLastRow();
    const emails = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase()) : [];
    const idx = emails.indexOf(email);
    const psh = reqSheet_();
    const pend = reqRows_(psh);
    const pidx = pend.findIndex(r => r.email === email);
    const dropPending = () => { if (pidx >= 0) psh.deleteRow(pidx + 2); };
    if (action === 'acl_add' || action === 'acl_approve') {
      const name = String(p.name || (pidx >= 0 ? pend[pidx].name : '') || '').slice(0, 60);
      const role = p.role === '管理者' ? '管理者' : '';
      if (idx >= 0) sh.getRange(idx + 2, 1, 1, 3).setValues([[email, name, role]]);
      else sh.appendRow([email, name, role, new Date()]);
      dropPending();
    } else if (action === 'acl_reject') {
      dropPending();
    } else if (action === 'acl_remove') {
      if (email === a.email) return bad('self', '自分自身は削除できません');
      if (idx >= 0) sh.deleteRow(idx + 2);
    } else {
      return bad('bad_action', '不明な操作です');
    }
  } finally { lock.releaseLock(); }
  return result();
}

// 許可されていれば null、そうでなければエラー応答（JSON）を返す
function authGuard_(e) {
  const p = (e && e.parameter) || {};
  let token = p.idToken || '';
  if (!token && e && e.postData && e.postData.contents) {
    try { token = (JSON.parse(e.postData.contents) || {}).idToken || ''; } catch (err) { /* JSON以外は無視 */ }
  }
  const a = verifyIdToken_(token);
  if (p.action === 'acl_request') return aclRequestHandle_(p, a);
  if (/^acl_/.test(p.action || '')) return aclHandle_(p.action, p, a);
  if (a.ok) return null;
  return jsonOut_({ status: 'error', success: false, ok: false, error: a.reason, data: a.reason, message: a.reason });
}
