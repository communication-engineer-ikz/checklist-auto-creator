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

    //CSV ファイルの値をGSS へコピー
        addSheetForCsvFileNotYetImported(spreadSheet);

        //CSV ファイル取り込み
        //GSS へ転記

    //余分なセルの削除
    if (maxRow - lastRow > 0) {
        sheet.deleteRows(lastRow + 1, maxRow - lastRow);
    }

    if (maxColumn - lastColumn - 1 > 0) {
        sheet.deleteColumns(lastColumn + 1, maxColumn - lastColumn - 1); //チェックボックスを追加する列の確保
    }

    //交互の背景色の適用
    /* 参考
        https://qiita.com/yamaotoko4177/items/4474217c18cc864bcc62
    */
    const targetRange = sheet.getRange(1, 1, lastRow, lastRow + 1);

    if (targetRange.getBandings()[0] != null) {
        console.log("交互の背景色は適用できません");
    } else {
        targetRange.applyRowBanding(SpreadsheetApp.BandingTheme.GREEN);
    }

    //一番右の列にチェックボックスを追加する
    /* 参考
        https://caymezon.com/gas-checkbox/#toc3
    */
    const checkboxColmunsRange = sheet.getRange(1, 7, lastRow);
    checkboxColmunsRange.insertCheckboxes();
}

function addSheetForCsvFileNotYetImported(spreadSheet) {

    const sheets = spreadSheet.getSheets();
    const csvFileListNotYetImported = findCsvFileNotYetImported(sheets);

    for (i = 0; i < csvFileListNotYetImported.length ; i++) {
        let newSheet = spreadSheet.insertSheet(sheets.length + i);
        newSheet.setName(csvFileListNotYetImported[i]);
    }
}

function findCsvFileNotYetImported(sheets) {

    const cardDetailsSheetList = [];
    const csvFileListNotYetImported = [];

    if (sheets.length == 0) return;
    for (const sheet of sheets) {
        cardDetailsSheetList.push(sheet.getName());
    }

    /* 参考
        https://moripro.net/gas-drive-get-filename/
    */
    const csvFilesFolderId = getCsvFilesFolderId();
    const files = DriveApp.getFolderById(csvFilesFolderId).getFiles();
    
    while (files.hasNext()) {
        
        const file = files.next();
        const filename = file.getName().replace(".csv", "");
        console.log(filename);

        if (!cardDetailsSheetList.includes(filename)) {
            csvFileListNotYetImported.push(filename);
        }
    }

    return csvFileListNotYetImported;
}