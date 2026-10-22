/**
 * プロジェクト内の全コードを取得し、ダウンロード・コピー用ダイアログを表示する
 */
function showCodeExportDialog() {
  // 【重要】OAuthスコープにDrive読み取り権限を自動認識させるための参照（実行はされません）
  if (false) {
    DriveApp.getFileById("dummy");
  }

  const scriptId = ScriptApp.getScriptId();
  const url = `https://script.google.com/feeds/download/export?id=${scriptId}&format=json`;
  
  const options = {
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    const errorMsg = `コードの取得に失敗しました。(HTTPステータス: ${responseCode})\nレスポンス: ${response.getContentText()}`;
    Logger.log(errorMsg);
    SpreadsheetApp.getUi().alert(errorMsg);
    return;
  }
  
  const projectData = JSON.parse(response.getContentText());
  const files = projectData.files;
  
  // テキスト全体の組み立て
  let combinedText = `GASプロジェクトコード一式\n出力日時: ${new Date().toLocaleString("ja-JP")}\n\n`;
  
  files.forEach(file => {
    const ext = file.type === "server_js" ? ".gs" : (file.type === "html" ? ".html" : "");
    combinedText += `==================================================\n`;
    combinedText += `FILE: ${file.name}${ext}\n`;
    combinedText += `==================================================\n`;
    combinedText += (file.source || "") + "\n\n\n";
  });
  
  // HTMLダイアログ用のテンプレート作成
  const template = HtmlService.createTemplate(getExportHtmlContent());
  template.codeText = combinedText;
  template.defaultFileName = `GAS_Codes_${SpreadsheetApp.getActiveSpreadsheet().getName()}_${Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd_HHmm")}.txt`;
  
  const html = template.evaluate().setWidth(650).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'GASコード一式の出力');
}

function getExportHtmlContent() {
  return `
  <!DOCTYPE html>
  <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: sans-serif; padding: 15px; margin: 0; }
        .btn-group { margin-bottom: 15px; display: flex; gap: 10px; }
        button {
          padding: 8px 16px;
          border-radius: 4px;
          border: none;
          cursor: pointer;
          font-weight: bold;
          font-size: 13px;
        }
        .btn-dl { background-color: #1a73e8; color: white; }
        .btn-dl:hover { background-color: #1557b0; }
        .btn-copy { background-color: #34a853; color: white; }
        .btn-copy:hover { background-color: #2b8242; }
        textarea {
          width: 100%;
          height: 380px;
          box-sizing: border-box;
          font-family: monospace;
          font-size: 11px;
          padding: 10px;
          border: 1px solid #ccc;
          border-radius: 4px;
          resize: none;
        }
      </style>
    </head>
    <body>
      <div class="btn-group">
        <button class="btn-dl" onclick="downloadFile()">💾 PCにテキスト保存 (保存先を選択)</button>
        <button class="btn-copy" onclick="copyToClipboard()">📋 全文をクリップボードにコピー</button>
      </div>
      <textarea id="codeArea" readonly><?= codeText ?></textarea>

      <script>
        function downloadFile() {
          const text = document.getElementById("codeArea").value;
          const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
          const link = document.createElement("a");
          link.href = URL.createObjectURL(blob);
          link.download = "<?= defaultFileName ?>";
          link.click();
        }

        function copyToClipboard() {
          const area = document.getElementById("codeArea");
          area.select();
          document.execCommand("copy");
          alert("クリップボードにコピーしました！そのまま貼り付けできます。");
        }
      </script>
    </body>
  </html>
  `;
}