// Day3: チェックボックスでLINE通知を飛ばす

// ✅ チェックが入ったらLINE通知＆完了日記録
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;

  // ✅ チェック列（C列 = 3列目）以外なら何もしない
  const checkColumn = 3;
  if (range.getColumn() !== checkColumn) return;

  // ✅ チェックが入ったら（TRUEのときだけ実行）
  if (range.getValue() !== true) return;

  // ✅ 担当者名は B列（2列目）と仮定
  const row = range.getRow();
  const 担当者名 = sheet.getRange(row, 2).getValue();
  const timestamp = new Date();

  // ✅ D列（4列目）に完了時刻を記録
  sheet.getRange(row, 4).setValue(timestamp);

  // ✅ LINE通知を送信（成功ログも表示）
  const result = sendLineNotification(担当者名, timestamp);
  if (result) {
    Logger.log(`✅ ${担当者名} の通知送信＆完了記録を行いました！`);
  } else {
    Logger.log(`❌ ${担当者名} の通知に失敗しました`);
  }
}

function sendLineNotification(担当者名, timestamp) {
  const token = PropertiesService.getScriptProperties().getProperty('LINE_TOKEN');
  if (!token) {
    Logger.log("❌ LINE_TOKENが未設定です");
    return false;
  }

  const formattedTime = Utilities.formatDate(timestamp, "Asia/Tokyo", "HH:mm:ss");
  const message = `📦 完了通知\n担当者: ${担当者名}\n時刻: ${formattedTime}`;

  try {
    const payload = JSON.stringify({
      messages: [{ type: "text", text: message }]
    });

    const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/broadcast", {
      method: "post",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + token
      },
      payload: payload,
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    const body = response.getContentText();

    Logger.log("📤 LINE送信レスポンス: " + body);

    return code >= 200 && code < 300;
  } catch (err) {
    Logger.log("❌ 通知エラー: " + err);
    return false;
  }
}

 // LINE通知用のトークンを設定する関数
 //スクリプトエディタから一度だけ実行してください
 //新しいトークンを取得した場合は、下記のtokenを更新してから再実行してください

function setLineToken() {
  const token = "自分のトークン";
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