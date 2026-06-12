// ==========================================
// SION CFC Cloud - フロントエンドロジック
// ==========================================

// ★★★ ここに先ほど取得したGASのウェブアプリURLを貼り付けます ★★★
const GAS_API_URL = "YOUR_GAS_URL_HERE";

// 画面読み込み時にデータ取得を実行
document.addEventListener("DOMContentLoaded", () => {
    if (GAS_API_URL === "https://script.google.com/macros/s/AKfycbzBjLmHCUbf53orKDA4_eMNlqzPBenIHt2JLCP7-AZ3pXMJTaQua41wlU-uaSRxBweG/exec") {
        showError("エラー: app.js を開いて GAS_API_URL を設定してください。");
        return;
    }
    fetchMachineData();
});

// GASから機器データを取得して画面に表示する関数
async function fetchMachineData() {
    const tbody = document.getElementById('machine-list');
    tbody.innerHTML = '<tr><td colspan="4" class="p-10 text-center text-gray-500"><i class="fa-solid fa-circle-notch fa-spin text-2xl mb-3 text-cyan-500"></i><br>データを同期中...</td></tr>';

    try {
        // GAS APIへGETリクエストを送信
        const response = await fetch(GAS_API_URL);
        const result = await response.json();

        if (result.status === 'success') {
            renderTable(result.data);
            document.getElementById('machine-count').innerText = result.data.length;
        } else {
            console.error("API Error:", result.message);
            showError("データの取得に失敗しました: " + result.message);
        }
    } catch (error) {
        console.error("Network Error:", error);
        showError("通信エラーが発生しました。GAS側のデプロイ設定（アクセス権限が「全員」になっているか）等を確認してください。");
    }
}

// テーブルにHTMLを描画する関数
function renderTable(data) {
    const tbody = document.getElementById('machine-list');
    tbody.innerHTML = ''; // プレースホルダーを消去

    // 1行目のヘッダー情報を除外（GAS側で処理していない場合のお守り）
    const validData = data.filter(item => item['機器管理番号'] !== '機器管理番号');

    if (validData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4" class="p-10 text-center text-gray-500">登録された機器がありません。スプレッドシートにデータを入力してください。</td></tr>';
        document.getElementById('machine-count').innerText = "0";
        return;
    }

    validData.forEach(machine => {
        const tr = document.createElement('tr');
        tr.className = "border-b border-gray-700 hover:bg-gray-750 transition-colors";
        
        // プロパティ名はスプレッドシートの1行目（ヘッダー）と完全に一致させる必要があります
        const machineId = machine['機器管理番号'] || '未設定';
        const location = (machine['施設名'] || '') + ' ' + (machine['設置場所'] || '');
        const model = (machine['メーカー'] || '') + ' ' + (machine['型式'] || '');
        
        tr.innerHTML = `
            <td class="p-4 font-mono text-cyan-300 font-medium">${machineId}</td>
            <td class="p-4 text-gray-200">${location}</td>
            <td class="p-4 text-gray-400 text-sm">${model}</td>
            <td class="p-4 text-center">
                <span class="px-3 py-1 bg-green-900/50 text-green-400 text-xs rounded-full border border-green-700/50 flex items-center justify-center w-max mx-auto">
                    <span class="w-1.5 h-1.5 rounded-full bg-green-400 mr-1.5 animate-pulse"></span> 正常稼働
                </span>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function showError(message) {
    document.getElementById('machine-list').innerHTML = 
        `<tr><td colspan="4" class="p-10 text-center text-red-400 bg-red-900/10 border-t border-b border-red-900/30"><i class="fa-solid fa-triangle-exclamation text-2xl mb-2"></i><br>${message}</td></tr>`;
}
