/**
 * シオンテクノス株式会社 - フロン排出抑制法一括管理システム API
 * * GitHub & CLASP 管理用メインソースコード
 */

// GETリクエスト：機器情報の取得（AppSheetやマスター参照用）
function doGet(e) {
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
