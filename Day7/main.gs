function get工程リスト(工事内容) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("テンプレート管理");
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const target = String(工事内容).trim(); // 空白削除 + 文字列化

  for (let i = 1; i < data.length; i++) {
    const templateName = String(data[i][0]).trim(); // こっちも空白削除
    if (templateName === target) {
      return data[i].slice(1).filter(step => step);
    }
  }

  Logger.log(`工事内容「${target}」に一致するテンプレートが見つかりません`);
  return null;
}