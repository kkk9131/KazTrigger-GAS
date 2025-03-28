// Day2: チェックボックスで日報転記

function onEdit(e) {
  const sheetName = "進捗管理";
  const logSheetName = "作業日報";

  const range = e.range;
  const sheet = e.source.getActiveSheet();

  if (sheet.getName() !== sheetName || range.getColumn() !== 3) return;
  if (range.getValue() !== true) return;

  const row = range.getRow();
  const values = sheet.getRange(row, 1, 1, 3).getValues()[0];

  const today = new Date();
  sheet.getRange(row, 4).setValue(today);

  const logSheet = e.source.getSheetByName(logSheetName);
  const lastRow = logSheet.getLastRow() + 1;
  logSheet.getRange(lastRow, 1, 1, 3).setValues([[values[0], values[1], today]]);
}