function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Vandelskontroll')
  .addItem('Match registrerte attester', 'matchAttester')
  .addItem('Send eposter', 'sendEposter')
  .addToUi();
}

function matchAttester() {
  var ui = SpreadsheetApp.getUi();

  var resultat = ui.prompt(
      'Angi trenerliste',
      'Hva heter dokumentet med trenere?',
      ui.ButtonSet.OK_CANCEL);
  var valgtKnapp = resultat.getSelectedButton();
  if (valgtKnapp === ui.Button.OK) {
    var trenerFilNavn = resultat.getResponseText();
    ui.alert('Åpner fil ' + trenerFilNavn + '.');
    var trenere = hentAlleTrenere(trenerFilNavn);
    var attesterSett = hentAlleAttesterSett();
    var trenerStatuser = trenerStatuser(trenere, attesterSett);
    var sheet = lagNyttArk();
    for (index = 0; index = trenerStatuser.length; index++) {
      sheet.appendRow(trenerStatuser[index]);
    }
  }
}

function sendEposter() {
  SpreadsheetApp.getUi()
    .alert('Send eposter er dessverre ikke implementert...');
}

function hentAlleTrenere(trenerFilNavn) {
  trenerFilNavn = 'coachReport_20170510.xls'; // DEBUG
  var coachesFileIterator = DriveApp.getFilesByName(trenerFilNavn);
  var hasNext = coachesFileIterator.hasNext();
  if (!coachesFileIterator.hasNext()) {
    return "Can't find file " + coachesListName;
  } else {
    var coaches = SpreadsheetApp.open(coachesFileIterator.next());
    var sheet = coaches.getActiveSheet();
    var range = sheet.getRange(1, 13, sheet.getLastRow(), 2);
    var values = range.getValues();
    return "File found  " + coachesListName;
  }
}

function hentAlleAttesterSett() {
  var attester_sett_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('attester_sett');
  var attestRange = attester_sett_sheet.getRange(1, 1, attester_sett_sheet.getLastRow(), 5);
  var attestArray = attestRange.getValues();
  return attestArray;
}

function trenerStatuser(trenerArray) {
  return [];
}

function lagNyttArk() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = nyttArkNavn();
  var sheet = ss.insertSheet(sheetName);
  return sheet;
}

function nyttArkNavn() {
  var dato = new Date();
  var aarString = dato.getFullYear();
  var maanedString = ('0' + (dato.getMonth() + 1)).slice(-2);
  var dagString = ('0' + dato.getDate()).slice(-2)
  var baseSheetName = aarString + '-' + maanedString + "-" + dagString;
  var counter = 2;
  var sheetName = baseSheetName;
  while (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)) {
    sheetName = baseSheetName + ' - v';
    if (counter < 10) {
      sheetName += '0' + counter;
    } else {
      sheetName += counter;
    }
    counter += 1;
  }
  return sheetName;
}
