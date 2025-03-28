// Day1: GAS初体験スクリプト

function myFirstScript() {
  // 今開いているスプレッドシートを取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // B3セルに文字を書き込む
  sheet.getRange("B3").setValue("お疲れ様です!");

  // 処理完了のログを出力（デバッグ用）
  Logger.log("書き込み完了！");
}