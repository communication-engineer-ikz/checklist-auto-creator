function findCsvFileNotYetImportedTest() {
    findCsvFileNotYetImported(SpreadsheetApp.openById(getCardDetailsSpreadSheetId()));
}