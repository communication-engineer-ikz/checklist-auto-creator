function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("追加メニュー");
    menu.addItem("シート整形", "checklistAutoCreator");
    menu.addToUi();
}

function checklistAutoCreator() {
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = spreadSheet.getActiveSheet();

    const sheets = spreadSheet.getSheets();

    const cardDetailsSheetList = [];

    if (sheets.length == 0) return;
    for (const sheet of sheets) {
        cardDetailsSheetList.push(sheet.getName());
    }

    /* 参考
        https://moripro.net/gas-drive-get-filename/
    */
    const csvFilesFolderId = getCsvFilesFolderId();
    const files = DriveApp.getFolderById(csvFilesFolderId).getFiles();

    const templateSheet = spreadSheet.getSheetByName("template"); //非表示のシート
    
    while (files.hasNext()) {
        
        const file = files.next();
        const filename = file.getName().replace(".csv", "");
        console.log(filename);

        if (!cardDetailsSheetList.includes(filename)) {
            /** 参考
                 * https://qiita.com/chihirot0109/items/d78ec1a6d14783545c32
             */
            let newSheet = spreadSheet.insertSheet(filename, sheets.length + 1, {template: templateSheet}).showSheet(); //複数のシートを追加したときにはシート順はソートされない
            let csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
            newSheet.getRange(1, 1, csvData.length, csvData[1].length).setValues(csvData);
        }
    }


    //CSV ファイルの値をGSS へコピー
        addSheetForCsvFileNotYetImported(spreadSheet);

        //CSV ファイル取り込み
        //GSS へ転記

    //余分なセルの削除
    const lastRow = activeSheet.getLastRow();
    const maxRow = activeSheet.getMaxRows();
    const lastColumn = activeSheet.getLastColumn();
    const maxColumn = activeSheet.getMaxColumns();

    if (maxRow - lastRow > 0) {
        activeSheet.deleteRows(lastRow + 1, maxRow - lastRow);
    }

    if (maxColumn - lastColumn - 1 > 0) {
        activeSheet.deleteColumns(lastColumn + 1, maxColumn - lastColumn - 1); //チェックボックスを追加する列の確保
    }

    //交互の背景色の適用
    /* 参考
        https://qiita.com/yamaotoko4177/items/4474217c18cc864bcc62
    */
    const targetRange = activeSheet.getRange(1, 1, lastRow, lastRow + 1);

    if (targetRange.getBandings()[0] != null) {
        console.log("交互の背景色は適用できません");
    } else {
        targetRange.applyRowBanding(SpreadsheetApp.BandingTheme.GREEN);
    }

    //一番右の列にチェックボックスを追加する
    /* 参考
        https://caymezon.com/gas-checkbox/#toc3
    */
    const checkboxColmunsRange = activeSheet.getRange(1, 7, lastRow);
    checkboxColmunsRange.insertCheckboxes();
}

function addSheetForCsvFileNotYetImported(spreadSheet) {

    const sheets = spreadSheet.getSheets();
    const templateSheet = spreadSheet.getSheetByName("template"); //非表示のシート
    const csvFileListNotYetImported = findCsvFileNotYetImported(sheets);

    for (i = 0; i < csvFileListNotYetImported.length ; i++) {
        spreadSheet.insertSheet(csvFileListNotYetImported[i], sheets.length + i, {template: templateSheet}).showSheet();
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

    return csvFileListNotYetImported.sort();
}