function checklistAutoCreator() {
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadSheet.getSheetByName("202106 のコピー");
    const lastRow = sheet.getLastRow();
    const maxRow = sheet.getMaxRows();
    const lastColumn = sheet.getLastColumn();
    const maxColumn = sheet.getMaxColumns();

    console.log(lastRow);
    console.log(maxRow);

    if (maxRow - lastRow > 0) {
        sheet.deleteRows(lastRow + 1, maxRow - lastRow);
    }

    console.log(lastColumn);
    console.log(maxColumn);

    if (maxColumn - lastColumn - 1 > 0) {
        sheet.deleteColumns(lastColumn + 1, maxColumn - lastColumn - 1); //チェックボックスを追加する列の確保

    }

/* 参考
 https://qiita.com/yamaotoko4177/items/4474217c18cc864bcc62
*/
    const targetRange = sheet.getRange(1, 1, lastRow, lastRow + 1);

    console.log(targetRange.getBandings());

    if (targetRange.getBandings()[0] != null) {
        console.log("交互の背景色は適用できません");
    } else {
        targetRange.applyRowBanding(SpreadsheetApp.BandingTheme.GREEN);
    }
}
