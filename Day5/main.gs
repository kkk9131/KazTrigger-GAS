function onEdit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = e.range;
  const editedRow = range.getRow();
  const editedCol = range.getColumn();

  const triggerRow = 113;
  const triggerCol = 1;

  // ✅ 自作フラグで「処理中かどうか」を管理
  const scriptProperties = PropertiesService.getScriptProperties();
  const isRunning = scriptProperties.getProperty("IS_RUNNING");

  if (isRunning === "true") return;

  if (editedRow === triggerRow && editedCol === triggerCol && range.getValue() === true) {
    try {
      // 処理中フラグON
      scriptProperties.setProperty("IS_RUNNING", "true");

      const originalSheet = e.source.getActiveSheet();
      const originalName = originalSheet.getName();

      const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd");
      const baseName = `${today}_積込`;

      let i = 1;
      let newName = `${baseName}_${i}`;
      while (ss.getSheetByName(newName)) {
        i++;
        newName = `${baseName}_${i}`;
      }

      const copiedSheet = ss.getSheetByName(originalName).copyTo(ss);
      const newSheet = ss.getSheets()[ss.getSheets().length - 1];
      newSheet.setName(newName);

      const startRow = 2;
      const endRow = 110;
      const clearCols = [3, 4, 5];

      clearCols.forEach(col => {
        newSheet.getRange(startRow, col, endRow - startRow + 1).clearContent();
      });

      newSheet.getRange(startRow, 4, endRow - startRow + 1).insertCheckboxes();
      newSheet.getRange(113, 1).insertCheckboxes().setValue(false); // A113
      newSheet.getRange(113, 2).insertCheckboxes().setValue(false); // B113
      newSheet.getRange(113, 3).clearContent(); // C113
      newSheet.getRange(113, 4).clearContent(); // D113
      newSheet.getRange(113, 5).clearContent(); // E113

      originalSheet.getRange(triggerRow, triggerCol).setValue(false);

      ss.toast(`✅ シート「${newName}」を1枚だけ追加しました`, "完了", 4);
    } finally {
      // 最後にフラグを解除
      scriptProperties.deleteProperty("IS_RUNNING");
    }

    return;
  }

  // --- 通知処理（省略：前のままでOK） ---
}
function sendLineNotification(name, timestamp) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token) {
    Logger.log("❌ LINE_TOKENが設定されていません");
    return false;
  }

  if (!(timestamp instanceof Date)) {
    timestamp = new Date(timestamp);
  }

  const timeStr = Utilities.formatDate(timestamp, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  const message = `📦 積み込み完了通知\n担当者: ${name}\n時刻: ${timeStr}`;

  try {
    const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/broadcast", {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + token
      },
      payload: JSON.stringify({
        messages: [{ type: "text", text: message }]
      }),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const body = response.getContentText();
    Logger.log(`📤 LINE通知レスポンス: ${code} - ${body}`);

    return (code >= 200 && code < 300);
  } catch (err) {
    Logger.log("❌ LINE通知エラー: " + err);
    return false;
  }
}
function setLineToken() {
  const token = "自分のLineトークン";
  PropertiesService.getScriptProperties().setProperty('LINE_TOKEN', token);
  Logger.log("✅ LINE_TOKENを設定しました");
  
  SpreadsheetApp.getUi().alert("✅ LINE_TOKENの設定が完了しました");
}

function testSendNotification() {
  const name = "テスト担当者";
  const timestamp = new Date();
  const result = sendLineNotification(name, timestamp);
  if (result) {
    SpreadsheetApp.getUi().alert("✅ LINE通知のテスト送信に成功しました！");
  } else {
    SpreadsheetApp.getUi().alert("❌ テスト通知に失敗しました。ログを確認してください。");
  }
}
function resetInput(sheet) {
  const checkStartRow = 2;
  const checkEndRow = 110;
  const checkCol = 4; // D列（チェックボックス）
  const quantityCol = 3; // C列（数量）

  const completeCheckRow = 113;
  const completeCheckCol = 2; // B列（完了チェック）
  const nameCol = 3; // C列（担当者名）
  const timeCol = 4; // D列（完了時間）

  for (let i = checkStartRow; i <= checkEndRow; i++) {
    sheet.getRange(i, checkCol).setValue(false);      // チェック解除
    sheet.getRange(i, quantityCol).clearContent();    // 数量クリア
  }

  sheet.getRange(completeCheckRow, nameCol).clearContent();  // 担当者
  sheet.getRange(completeCheckRow, timeCol).clearContent();  // 時刻
  sheet.getRange(completeCheckRow, completeCheckCol).setValue(false); // 完了チェック

  SpreadsheetApp.getActiveSpreadsheet().toast("✅ 入力を初期化しました", "リセット完了", 5);
}