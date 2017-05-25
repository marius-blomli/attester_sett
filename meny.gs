function addVandelskontrollMeny() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Vandelskontroll')
  .addItem('Match registrerte attester', 'matchAttester')
  .addItem('Send eposter', 'sendEposter')
  .addToUi();
}
