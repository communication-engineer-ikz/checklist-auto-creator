function onOpen() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("CheckList AutoCreator");
    menu.addItem("CSV取込", "checklistAutoCreator");
    menu.addItem("Repository Info.", "displayRepositoryInfo");
    menu.addToUi();
}

function checklistAutoCreator() {
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = spreadSheet.getActiveSheet();

    const csvFilesFolderId = getCsvFilesFolderId();

    uploadCsvFileDataFromGDriveFolder(spreadSheet, csvFilesFolderId);
}

function uploadCsvFileDataFromGDriveFolder(spreadSheet, csvFilesFolderId) {

    const templateSheet = spreadSheet.getSheetByName("template"); //非表示のシート

    const sheets = spreadSheet.getSheets();
    if (sheets.length == 0) return;

    const cardDetailsSheetList = [];
    for (const sheet of sheets) {
        cardDetailsSheetList.push(sheet.getName());
    }

    /* 参考
        https://moripro.net/gas-drive-get-filename/
    */
    const files = DriveApp.getFolderById(csvFilesFolderId).getFiles();
    while (files.hasNext()) {
        
        const file = files.next();
        const filename = file.getName().replace(".csv", "");
        console.log(filename);

        if (!cardDetailsSheetList.includes(filename)) {
            /** 参考
                 * https://qiita.com/chihirot0109/items/d78ec1a6d14783545c32
             */
            const newSheet = spreadSheet.insertSheet(filename, sheets.length + 1, {template: templateSheet}).showSheet(); //複数のシートを追加したときにはシート順はソートされない
            const csvData = Utilities.parseCsv(file.getBlob().getDataAsString("sjis"));
            newSheet.getRange(1, 1, csvData.length, csvData[1].length).setValues(csvData);
        }
    }
}

function displayRepositoryInfo() {
    return Browser.msgBox("Repository Info.",
        "communication-engineer-ikz / checklist-auto-creator" + "\\n" + "https://github.com/communication-engineer-ikz/checklist-auto-creator",
        Browser.Buttons.OK);
}