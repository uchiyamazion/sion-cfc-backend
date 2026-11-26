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
// AuthGuard.gs ― Googleログインの本人確認（GAS側）
// 画面から送られたGoogleのIDトークンを確認し、許可リストにあるアカウントだけを通す。
//
// 【設定】Apps Scriptの「プロジェクトの設定 → スクリプト プロパティ」に
//   ALLOWED_EMAILS = a@gmail.com,b@gmail.com（カンマ区切り）を登録する。未設定のときは全員を拒否する。
// 【使い方】doGet(e) と doPost(e) の先頭に、次の2行を入れる。
//   const g = authGuard_(e); if (g) return g;
// ════════════════════════════════════════════════

const GOOGLE_CLIENT_ID = '782190614995-5d4fur401smd1gpnnvs43h3m2q2q2ar3.apps.googleusercontent.com';

function allowedEmails_() {
  const v = PropertiesService.getScriptProperties().getProperty('ALLOWED_EMAILS') || '';
  return v.split(/[\s,;]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
}

// 戻り値：{ ok:true, email } または { ok:false, reason:'unauthorized' | 'forbidden' }
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
  if (allowedEmails_().indexOf(email) < 0) return { ok: false, reason: 'forbidden' };
  return { ok: true, email: email };
}

// 許可されていれば null、そうでなければエラー応答（JSON）を返す
function authGuard_(e) {
  let token = (e && e.parameter && e.parameter.idToken) || '';
  if (!token && e && e.postData && e.postData.contents) {
    try { token = (JSON.parse(e.postData.contents) || {}).idToken || ''; } catch (err) { /* JSON以外は無視 */ }
  }
  const a = verifyIdToken_(token);
  if (a.ok) return null;
  return ContentService.createTextOutput(JSON.stringify({
    status: 'error', success: false, ok: false, error: a.reason, data: a.reason, message: a.reason
  })).setMimeType(ContentService.MimeType.JSON);
}
