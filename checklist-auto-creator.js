function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("追加メニュー");
    menu.addItem("シート整形", "checklistAutoCreator");
    menu.addToUi();
}

function checklistAutoCreator() {
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadSheet.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const maxRow = sheet.getMaxRows();
    const lastColumn = sheet.getLastColumn();
    const maxColumn = sheet.getMaxColumns();

    if (maxRow - lastRow > 0) {
        sheet.deleteRows(lastRow + 1, maxRow - lastRow);
    }

    if (maxColumn - lastColumn - 1 > 0) {
        sheet.deleteColumns(lastColumn + 1, maxColumn - lastColumn - 1); //チェックボックスを追加する列の確保
    }

    /* 参考
        https://qiita.com/yamaotoko4177/items/4474217c18cc864bcc62
    */
    const targetRange = sheet.getRange(1, 1, lastRow, lastRow + 1);

    if (targetRange.getBandings()[0] != null) {
        console.log("交互の背景色は適用できません");
    } else {
        targetRange.applyRowBanding(SpreadsheetApp.BandingTheme.GREEN);
    }

    /* 参考
        https://caymezon.com/gas-checkbox/#toc3
    */
    const checkboxColmunsRange = sheet.getRange(1, 7, lastRow);
    checkboxColmunsRange.insertCheckboxes();
}