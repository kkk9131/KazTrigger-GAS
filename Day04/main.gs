/**
 * チェックボックスが押されたときに動くメイン関数
 * 各種処理（通知・記録・リセット）を呼び出す入口
 */
function onEdit(e) {
  const sheetName = '進捗管理';
  const checkCol = 4;
  const timeCol = 5;
  const notifiedCol = 6;

  const sheet = e.source.getSheetByName(sheetName);
  const activeSheet = e.source.getActiveSheet();

  if (!activeSheet || activeSheet.getName() !== sheetName) return;

  const range = e.range;
  if (range.getColumn() !== checkCol || range.getRow() === 1) return;

  const row = range.getRow();
  const checkboxValue = range.getValue();
  if (checkboxValue !== true) return;

  const taskName = sheet.getRange(row, 1).getValue();
  const notified = sheet.getRange(row, notifiedCol).getValue();
  const sheetNameLabel = sheet.getName();

  if (notified !== "通知済") {
    const timestamp = new Date();
    sheet.getRange(row, timeCol).setValue(timestamp);

    if (sendLineNotification(taskName, sheetNameLabel)) {
      sheet.getRange(row, notifiedCol).setValue("通知済");
      recordToMonthlySheet(taskName, timestamp);

      if (taskName === "最終確認") {
        const sheetsToReset = ["現場情報", "リーダー指示書", "アシスタント指示書", "進捗管理"];
        resetMultipleSheets(sheetsToReset);
      }
    }
  }
}
// --- LINE通知を送る関数 ---
function sendLineNotification(taskName, sheetNameLabel) {
  try {
    const props = PropertiesService.getScriptProperties();
    let accessToken = props.getProperty('LINE_TOKEN');

    if (!accessToken) {
      SpreadsheetApp.getActiveSpreadsheet().toast("LINEトークンが未設定です", "エラー", 5);
      return false;
    }

    let 現場名 = "", 作業者 = "";
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const siteSheet = ss.getSheetByName("現場情報");
      if (siteSheet) {
        現場名 = siteSheet.getRange("B4").getValue() || "未設定";
        const 作業者1 = siteSheet.getRange("B7").getValue() || "";
        const 作業者2 = siteSheet.getRange("B8").getValue() || "";
        作業者 = 作業者1 && 作業者2 ? `${作業者1}・${作業者2}` : (作業者1 || 作業者2 || "未設定");
      }
    } catch (e) {
      現場名 = "未設定";
      作業者 = "未設定";
    }

    const message = `📋 作業完了通知\n現場名：${現場名}\nシート名：${sheetNameLabel}\n作業項目：${taskName}\n作業者：${作業者}\n完了時刻：${new Date().toLocaleString('ja-JP')}`;
    const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/broadcast", {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + accessToken
      },
      payload: JSON.stringify({ messages: [{ type: "text", text: message }] }),
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code >= 200 && code < 300) {
      return true;
    } else {
      Logger.log("LINE通知失敗：" + response.getContentText());
      return false;
    }
  } catch (error) {
    Logger.log("通知エラー：" + error);
    return false;
  }
}
// --- 月次シートに記録する関数 ---
function recordToMonthlySheet(taskName, timestamp) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const siteSheet = ss.getSheetByName("現場情報");
    const rawDate = siteSheet.getRange("B5").getValue();
    const 作業日 = rawDate ? new Date(rawDate) : new Date();
    const 現場名 = siteSheet.getRange("B4").getValue() || "未設定";
    const 作業者1 = siteSheet.getRange("B7").getValue() || "未設定";
    const 作業者2 = siteSheet.getRange("B8").getValue() || "";
    const 作業者 = 作業者2 ? `${作業者1}・${作業者2}` : 作業者1;

    const yearMonth = Utilities.formatDate(作業日, "Asia/Tokyo", "yyyy-MM");
    let reportSheet = ss.getSheetByName(yearMonth);
    if (!reportSheet) {
      reportSheet = ss.insertSheet(yearMonth);
      reportSheet.appendRow(["作業日", "現場名", "作業者", "作業項目", "完了時刻"]);
    }

    reportSheet.appendRow([作業日, 現場名, 作業者, taskName, timestamp]);
    return true;
  } catch (error) {
    Logger.log("記録エラー：" + error);
    return false;
  }
}
// LINE通知用のトークンを設定する（初回だけ実行）
function setLineToken() {
  const token = "ここに自分のLINEトークンを貼り付けてね";
  PropertiesService.getScriptProperties().setProperty('LINE_TOKEN', token);
}

// トリガーの手動設定（初回だけ実行）
function setupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('onEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
}

// テスト通知を送る（動作確認用）
function testNotification() {
  const result = sendLineNotification("テスト通知", "テスト");
  const msg = result ? "送信成功" : "送信失敗";
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, "テスト結果", 3);
}