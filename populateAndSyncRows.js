function populateAndSyncRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source1 = ss.getSheetByName("ChatGPT 👽");
  const source2 = ss.getSheetByName("Claude 👽");
  const source3 = ss.getSheetByName("Grok 👽");
  const source4 = ss.getSheetByName("X 👽");
  const source5 = ss.getSheetByName("Gemini 👽");
  const source6 = ss.getSheetByName("Perplexity 👽");
  const source7 = ss.getSheetByName("Deepseek 👽");
  const combination = ss.getSheetByName("AI CHATS 👽");

  combination.clearContents(); // Clear existing data in combination

  function getRowsWithY(sheet, sheetName) {
    const range = sheet.getDataRange();
    const values = range.getValues();
    const rowsWithY = [];

    for (let row = 0; row < values.length; row++) {
      if (values[row][4] === "y") {
        // Check if column E is "y"
        // rowsWithY.push([...values[row], sheetName, row + 1]); // Include source sheet name and row
        rowsWithY.push([...values[row]]);
      }
    }
    return rowsWithY;
  }

  const allRowsWithY = [
    ...getRowsWithY(source1, "ChatGPT"),
    ...getRowsWithY(source2, "Claude"),
    ...getRowsWithY(source3, "Grok"),
    ...getRowsWithY(source4, "X"),
    ...getRowsWithY(source5, "Gemini"),
    ...getRowsWithY(source6, "Perplexity"),
    ...getRowsWithY(source7, "Deepseek"),
  ];

  if (allRowsWithY.length > 0) {
    combination
      .getRange(1, 1, allRowsWithY.length, allRowsWithY[0].length)
      .setValues(allRowsWithY);
  }
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "AI CHATS 👽") return;

  const row = e.range.getRow();
  const column = e.range.getColumn();
  const combinationSheet = e.source.getSheetByName("AI CHATS 👽");
  const sheetName = combinationSheet
    .getRange(row, combinationSheet.getLastColumn() - 1)
    .getValue();
  const originalRow = combinationSheet
    .getRange(row, combinationSheet.getLastColumn())
    .getValue();

  const targetSheet = e.source.getSheetByName(sheetName);
  if (targetSheet) {
    targetSheet.getRange(originalRow, column).setValue(e.value);
  }
}
