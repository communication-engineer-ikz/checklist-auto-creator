function addSheetForCsvFileNotYetImportedTest() {
    addSheetForCsvFileNotYetImported(SpreadsheetApp.openById(getCardDetailsSpreadSheetId()));
}

function findCsvFileNotYetImportedTest() {
    findCsvFileNotYetImported(SpreadsheetApp.openById(getCardDetailsSpreadSheetId()));
}