var SETTINGS = getData("Ustawienia", "B:B")
var SETTINGS_CHECKPOINT_1 = 1
var SETTINGS_CHECKPOINT_2 = 2
var SETTINGS_CHECKPOINT_3 = 3
var SETTINGS_CHECKPOINT_4 = 4
var SETTINGS_MAIL_YELLOW_GRANT = 5
var SETTINGS_MAIL_RED_GRANT = 6
var SETTINGS_MAIL_YELLOW_PICK_UP = 7
var SETTINGS_MAIL_RED_PICK_UP = 8
var SETTINGS_CHECKPOINT_TITLE_1 = 9
var SETTINGS_CHECKPOINT_TITLE_2 = 10
var SETTINGS_CHECKPOINT_TITLE_3 = 11
var SETTINGS_CHECKPOINT_TITLE_4 = 12
var SETTINGS_CHECKPOINT_TITLE_5 = 13
var SETTINGS_DEFAULT_MAIL = 14
var SETTINGS_MAIL_RED1_GRANT = 15
var SETTINGS_MAIL_RED1_PICK_UP = 16

var LEVEL_YELLOW_CAR_TO_GET_RED = 2

var CHECKPOINT_TITLE = [
  "",
  SETTINGS[SETTINGS_CHECKPOINT_TITLE_1][0],
  SETTINGS[SETTINGS_CHECKPOINT_TITLE_2][0],
  SETTINGS[SETTINGS_CHECKPOINT_TITLE_3][0],
  SETTINGS[SETTINGS_CHECKPOINT_TITLE_4][0],
  SETTINGS[SETTINGS_CHECKPOINT_TITLE_5][0],
]

var YELLOW = "yellow"
var RED = "red"

var TEST = true
var WRITE = true

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Dodatkowe opcje')
    .addItem('Sprawdź terminy i NIE wysyłaj maili (TEST)', 'runTest')
    .addItem('Sprawdź terminy i wyślij maile', 'runMain')
    .addToUi();
}

function testValue(value, expected) {
  if (value != expected) {
    console.error("value: " + value + " is not " + expected)
  }
}

function testGetRowsData() {
  Logger.log(getRowsData().length)
  Logger.log(getRowsData()[9])
}

function testCharToInt() {
  testValue(charToInt('A'), 0)
  testValue(charToInt('B'), 1)
  testValue(charToInt('C'), 2)
  console.log("end test")
}

function testSettings() {
  console.log(SETTINGS);
}


function testEmpty() {
  var empty = ""
  //if something is fill
  if (empty) { console.log("not empty") } else { console.log("empty") } // empty 
  if (!empty) { console.log("empty") } else { console.log("not empty") } // empty

  var something = "s"
  if (something) { console.log("not empty") } else { console.log("empty") } // not empty
  if (!something) { console.log("empty") } else { console.log("not empty") } // not empty
}

function testCheckData() {
  console.log(checkData())
}

function testGetRecipients() {
  var list = [];
  list.push("first")
  var str = "a,b,c"
  var out = list.concat(str.split(","));
  console.log(out)
  console.log(out.join(","))
}


function runMain() {
  TEST = false
  main2()
}

function runTest() {
  TEST = true
  WRITE = false
  main2()
}


function testMain2() {
  console.log("#1 - save")
  checkCard(true, "", 7, 0, YELLOW)
  console.log("#2")
  checkCard(true, new Date(), 7, 0, YELLOW)
  console.log("#3 - delete")
  checkCard("", new Date(), 7, 0, YELLOW)
  console.log("#4")
  checkCard("", "test", 7, 0, YELLOW)

  console.log("#1 - save")
  checkCard(true, "", 7, 0, RED)
  console.log("#2")
  checkCard(true, new Date(), 7, 0, RED)
  console.log("#3 - delete")
  checkCard("", new Date(), 7, 0, RED)
  console.log("#4")
  checkCard("", "test", 7, 0, RED)
  console.log("#5")
  checkCard(1, "", 7, 1, RED)
  console.log("#6 - save")
  checkCard(-1, "", 7, 1, RED)
  console.log("#7 - delete")
  checkCard("", new Date(), 7, 1, RED)

  console.log("#1")
  checkMoreThreeCard(true, "", 7)
  console.log("#2")
  checkMoreThreeCard(1, "", 7)
  console.log("#3")
  checkMoreThreeCard(2, "", 7)
  console.log("#4")
  checkMoreThreeCard(3, "", 7)
  console.log("#5 - save")
  checkMoreThreeCard(4, "", 7)
  console.log("#6 - delete")
  checkMoreThreeCard(3, new Date(), 7)


}

function main2() {
  // TEST = true
  // WRITE = false

  var summary = getSummaryData()
  var saveData = getSaveData(summary.length)

  if (summary.length != saveData.length) {
    console.error("Ilość wyliczonych kartek nie zgadza się z zapisanymi (" + summary.length + " " + saveData.length + ")")
    return;
  }

  for (var i = 0; i < summary.length; i++) {

    var rowSummary = summary[i];
    var rowSave = saveData[i];
    var id = rowSummary.id

    if (!rowSummary.canceled) {
      console.log("Canceled")
      continue
    }
    for (var c = 0; c < rowSummary.term_yellow.length; c++) {
      var yellow = rowSummary.term_yellow[c]
      var saveYellow = rowSave.term_yellow[c]
      checkCard(yellow, saveYellow, id, c, YELLOW, rowSummary.three_yellow, rowSummary.term_to_fix[c])
    }
    for (var c = 0; c < rowSummary.term_red.length; c++) {
      var red = rowSummary.term_red[c]
      var saveRed = rowSave.term_red[c]
      checkCard(red, saveRed, id, c, RED, rowSummary.three_yellow, null)
    }
    checkMoreThreeCard(rowSummary.three_yellow, rowSave.three_yellow, id)
  }
}

function checkCard(actual, saveData, id, idCard, color, numberYellowCard, dayToFix) {
  if (
    (typeof actual == "boolean" && actual == true) ||
    (typeof actual == "number" && actual < 0)
  ) {
    if (typeof saveData.getMonth !== 'function') {
      saveCard2(id, idCard, color)
      sendMailForCard2(id, idCard, color, numberYellowCard, dayToFix)
    }
  } else {
    if (typeof saveData.getMonth === 'function') {
      deleteCard2(id, idCard, color)
      sendMailForTakeBack2(id, idCard, color)
    }
  }
}

function checkMoreThreeCard(actual, saveData, id) {
  if (
    (typeof actual == "number" && actual > 3)
  ) {
    if (typeof saveData.getMonth !== 'function') {
      saveRedCard2(id)
      sendMailForRedCard2(id)
    }
  } else {
    if (typeof saveData.getMonth === 'function') {
      deleteRedCard2(id)
      sendMailForTakeBackRed2(id)
    }
  }
}

function saveCard2(id, idCard, color) {
  console.log("save " + color + " " + id + " " + (idCard + 1))
  if (WRITE) {
    saveValue('stan kartek2', getLocationForCard(id, idCard, color), new Date())
  }
}

function saveRedCard2(id) {
  console.log("saveRED " + id)
  if (WRITE) {
    saveValue('stan kartek2', getLocationForRedCard(id), new Date())
  }
}

function deleteCard2(id, idCard, color) {
  console.log("delete " + color + " " + id + " " + (idCard + 1))
  if (WRITE) {
    saveValue('stan kartek2', getLocationForCard(id, idCard, color), "")
  }
}

function deleteRedCard2(id) {
  console.log("deleteRed " + id)
  if (WRITE) {
    saveValue('stan kartek2', getLocationForRedCard(id), "")
  }
}

function getLocationForCard(id, idCard, color) {
  if (color == YELLOW) {
    return intToChar(idCard + 1) + (id + 1)
  } else if (color == RED) {
    return intToChar(idCard + 6) + (id + 1)
  } else {
    return -1
  }
}

function getLocationForRedCard(id) {
  return "L" + (id + 1)
}

function sendMailForCard2(id, idCard, color, numberYellowCard, dayToFix) {
  var level = idCard + 1;
  var mailData = getMailData(id)
  var numberYellowCardText = ""
  if (numberYellowCard > 0) {
    numberYellowCardText += numberYellowCard
  }
  var checkPointName = CHECKPOINT_TITLE[level]
  var subject = "HAZ 2023 " + mailData.komendant + "/" + mailData.kwatermitrz + " [" + numberYellowCardText + " " + colorToText(color) + " kartka - termin nr " + level + "]"
  var message = ""
  if (color == YELLOW) {
    var dtf = Utilities.formatDate(dayToFix, "GMT+1", "yyyy-MM-dd");
    message = getMessageForYellow(SETTINGS_MAIL_YELLOW_GRANT, mailData.komendant, mailData.kwatermitrz, numberYellowCard, checkPointName, dtf)
  } else {
    message = getMessageForYellow(SETTINGS_MAIL_RED1_GRANT, mailData.komendant, mailData.kwatermitrz, numberYellowCard, checkPointName, "")
  }

  sendMail(mailData.recipeint, subject, message)
  internalLog([new Date(), id, "przyznanie", colorToText(color), (idCard + 1), message, mailData.recipeint])
}

function sendMailForTakeBack2(id, idCard, color) {
  var level = idCard + 1;
  var mailData = getMailData(id)
  var subject = "HAZ 2023 " + mailData.komendant + "/" + mailData.kwatermitrz + " unieważnienie [Żółta kartka - termin nr " + level + "]"
  var message = ""
  if (color == YELLOW) {
    var message = getMessage(SETTINGS_MAIL_YELLOW_PICK_UP, mailData.komendant, mailData.kwatermitrz)
  } else {
    var message = getMessage(SETTINGS_MAIL_RED1_PICK_UP, mailData.komendant, mailData.kwatermitrz)
  }

  sendMail(mailData.recipeint, subject, message)
  internalLog([new Date(), id, "odebranie", colorToText(color), (idCard + 1), message, mailData.recipeint])
}

function sendMailForRedCard2(id, idCard) {
  var mailData = getMailData(id)
  var subject = "HAZ 2023 " + mailData.komendant + "/" + mailData.kwatermitrz + " [Czerwona kartka]"
  var message = getMessage(SETTINGS_MAIL_RED_GRANT, kom, kwa)
  sendMail(mailData.recipeint, subject, message)
  internalLog([new Date(), id, "przyznanie", colorToText(color), (idCard + 1), message, mailData.recipeint])
}

function sendMailForTakeBackRed2(id, idCard) {
  var mailData = getMailData(id)
  var subject = "HAZ 2023 " + mailData.komendant + "/" + mailData.kwatermitrz + " unieważnienie [Czerwona kartka]"
  var message = getMessage(SETTINGS_MAIL_RED_PICK_UP, mailData.komendant, mailData.kwatermitrz)
  sendMail(mailData.recipeint, subject, message)
  internalLog([new Date(), id, "przyznanie", colorToText(color), (idCard + 1), message, mailData.recipeint])
}

function colorToText(color) {
  if (color == YELLOW) {
    return "żółta"
  } else if (color == RED) {
    return "czerwona"
  } else {
    ""
  }
}

function getMailData(id) {
  var dataToReceipient = dataToReceipientRaw(id);
  var recipeint = getRecipients(dataToReceipient)
  if (!recipeint) {
    console.error("Błąd z pobraniem adresu email")
  }
  return {
    recipeint: recipeint,
    komendant: dataToReceipient[0],
    kwatermitrz: dataToReceipient[2],
  };
}

function getSummaryData() {
  var data = getData("Podsumowanie Kartek", "A3:" + getLastRow("Podsumowanie Kartek"))
  var output = []

  var id = charToInt("A")
  var canceled = charToInt("y")

  var yellow1 = charToInt("AB")
  var yellow2 = charToInt("AF")
  var yellow3 = charToInt("AJ")
  var yellow4 = charToInt("AN")
  var yellow5 = charToInt("AR")

  var red1 = charToInt("AC")
  var red2 = charToInt("AG")
  var red3 = charToInt("AK")
  var red4 = charToInt("AO")
  var red5 = charToInt("AS")

  var fix1 = charToInt("AD")
  var fix2 = charToInt("AH")
  var fix3 = charToInt("AL")
  var fix4 = charToInt("AP")
  var fix5 = charToInt("AP")

  var threeYellow = charToInt("AT")

  for (var i = 0; i < data.length; i++) {
    var dataRow = data[i]
    var row = {
      id: dataRow[id],
      canceled: canceled,
      term_yellow: [dataRow[yellow1], dataRow[yellow2], dataRow[yellow3], dataRow[yellow4], dataRow[yellow5]],
      term_red: [dataRow[red1], dataRow[red2], dataRow[red3], dataRow[red4], dataRow[red5]],
      three_yellow: dataRow[threeYellow],
      term_to_fix: [dataRow[fix1], dataRow[fix2], dataRow[fix3], dataRow[fix4], dataRow[fix5]],
    }
    output.push(row)
  }
  return output;
}

function getSaveData(rowsLength) {
  var yellow1 = charToInt("B")
  var yellow2 = charToInt("C")
  var yellow3 = charToInt("D")
  var yellow4 = charToInt("E")
  var yellow5 = charToInt("F")

  var red1 = charToInt("G")
  var red2 = charToInt("H")
  var red3 = charToInt("I")
  var red4 = charToInt("J")
  var red5 = charToInt("K")

  var threeYellow = charToInt("L")

  var data = getData("stan kartek2", "A2:L" + (rowsLength + 1));
  var output = []
  for (var i = 0; i < data.length; i++) {
    var dataRow = data[i]
    var row = {
      term_yellow: [dataRow[yellow1], dataRow[yellow2], dataRow[yellow3], dataRow[yellow4], dataRow[yellow5]],
      term_red: [dataRow[red1], dataRow[red2], dataRow[red3], dataRow[red4], dataRow[red5]],
      three_yellow: dataRow[threeYellow]
    }
    output.push(row)
  }
  return output;
}



//// OLD VERSION
function main() {
  var yellowCardsForTrips = checkData()
  var yellowCardState = getData("stan kartek", "A2:F" + (yellowCardsForTrips.length + 1));

  if (yellowCardsForTrips.length != yellowCardState.length) {
    console.error("Ilość wyliczonych kartek nie zgadza się z zapisanymi (" + yellowCardsForTrips.length + " " + yellowCardState.length + ")")
    return;
  }

  for (var i = 0; i < yellowCardsForTrips.length; i++) {
    // var i = 6 //test

    var sumYellowCard = 0
    var id = yellowCardsForTrips[i][0]
    var row = yellowCardsForTrips[i][1]
    var numberYellowCard = 1
    for (var c = 1; c <= 4; c++) {
      if (checkIfTripIsCanceled(row)) {
        continue
      }

      var calculateCard = yellowCardsForTrips[i][c + 1]
      var saveData = yellowCardState[i][c]

      if (!saveData && calculateCard.length > 0 && sumYellowCard <= LEVEL_YELLOW_CAR_TO_GET_RED) {
        sendYellowCard(id, row, c, calculateCard, numberYellowCard)
      } else if (saveData && calculateCard.length == 0) {
        takeBackYellowCard(id, row, c)
      }

      if (calculateCard.length > 0) {
        sumYellowCard++;
        numberYellowCard++
      }
    }
    var saveRedData = yellowCardState[i][5]
    if (sumYellowCard > LEVEL_YELLOW_CAR_TO_GET_RED) {
      sendRedCard(id, row)
    } else if (saveRedData && sumYellowCard <= LEVEL_YELLOW_CAR_TO_GET_RED) {
      takeBackRedCard(id, row)
    }
  }
}

function sendYellowCard(id, row, level, calculateCard, numberYellowCard) {
  var message = sendMailSendYellowCard(id, level, calculateCard, numberYellowCard, getDateForNextLevel(row, level))
  var dataToReceipient = dataToReceipientRaw(id);
  var recipeint = getRecipients(dataToReceipient);
  internalLog([new Date(), row[0], "przyznanie", "żółta", level, message, recipeint])
  saveValue('stan kartek', intToChar(level) + (row[0] + 1), new Date())
}

function takeBackYellowCard(id, row, level) {
  var message = sendMailTakeBackYellowCard(id, level)
  var dataToReceipient = dataToReceipientRaw(id);
  var recipeint = getRecipients(dataToReceipient);
  internalLog([new Date(), row[0], "odebranie", "żółta", level, message, recipeint])
  saveValue('stan kartek', intToChar(level) + (row[0] + 1), "")
}

function sendRedCard(id, row) {
  var message = sendMailSendRedCard(id)
  var dataToReceipient = dataToReceipientRaw(id);
  var recipeint = getRecipients(dataToReceipient);
  internalLog([new Date(), row[0], "przyznanie", "czerwona", "", message, recipeint])
  saveValue('stan kartek', "F" + (row[0] + 1), new Date())
}

function takeBackRedCard(id, row) {
  var message = sendMailTakeBackRedCard(id)
  var dataToReceipient = dataToReceipientRaw(id);
  var recipeint = getRecipients(dataToReceipient);
  internalLog([new Date(), row[0], "odebranie", "czerwona", "", message, recipeint])
  saveValue('stan kartek', "F" + (row[0] + 1), "")
}


function getDateForNextLevel(row, level) {
  var columnsForCheckpoint = getColumnNumberForCheckpoint(level + 1)
  var currentDate = null

  for (var c = 0; c < columnsForCheckpoint.length; c++) {
    var index = columnsForCheckpoint[c]
    var date = new Date(row[index])
    date.setHours(23, 59, 59);

    if (currentDate == null || date < currentDate) {
      currentDate = date
    }
  }
  return Utilities.formatDate(currentDate, "GMT+1", "yyyy-MM-dd");
}

function sendMailSendYellowCard(id, level, calculateCard, numberYellowCard, dayToFix) {
  var dataToReceipient = dataToReceipientRaw(id);
  var recipeint = getRecipients(dataToReceipient)
  if (!recipeint) {
    console.error("Błąd z pobraniem adresu email")
  }
  var kom = dataToReceipient[0];
  var kwa = dataToReceipient[2];
  var checkPointName = CHECKPOINT_TITLE[level + 1]

  var subject = "HAZ 2023 " + kom + "/" + kwa + " [" + numberYellowCard + " żółta kartka - termin nr " + level + "]"
  var message = getMessageForYellow(SETTINGS_MAIL_YELLOW_GRANT, kom, kwa, numberYellowCard, checkPointName, dayToFix)
  sendMail(recipeint, subject, message)
  return message
}

function sendMailTakeBackYellowCard(id, level) {
  var dataToReceipient = dataToReceipientRaw(id);
  var recipeint = getRecipients(dataToReceipient)
  if (!recipeint) {
    console.error("Błąd z pobraniem adresu email")
  }
  var kom = dataToReceipient[0];
  var kwa = dataToReceipient[2];
  var subject = "HAZ 2023 " + kom + "/" + kwa + " unieważnienie [Żółta kartka - termin nr " + level + "]"
  var message = getMessage(SETTINGS_MAIL_YELLOW_PICK_UP, kom, kwa)
  sendMail(recipeint, subject, message)
  return message;
}

function sendMailSendRedCard(id) {
  var dataToReceipient = dataToReceipientRaw(id);
  var recipeint = getRecipients(dataToReceipient)
  if (!recipeint) {
    console.error("Błąd z pobraniem adresu email")
  }
  var kom = dataToReceipient[0];
  var kwa = dataToReceipient[2];
  var subject = "HAZ 2023 " + kom + "/" + kwa + " [Czerwona kartka]"
  var message = getMessage(SETTINGS_MAIL_RED_GRANT, kom, kwa)
  sendMail(recipeint, subject, message)
  return message;
}

function sendMailTakeBackRedCard(id) {
  var dataToReceipient = dataToReceipientRaw(id);
  var recipeint = getRecipients(dataToReceipient)
  if (!recipeint) {
    console.error("Błąd z pobraniem adresu email")
  }
  var kom = dataToReceipient[0];
  var kwa = dataToReceipient[2];
  var subject = "HAZ 2023 " + kom + "/" + kwa + " unieważnienie [Czerwona kartka]"
  var message = getMessage(SETTINGS_MAIL_RED_PICK_UP, kom, kwa)
  sendMail(recipeint, subject, message)
  return message
}

function dataToReceipientRaw(id) {
  return getData("Dane aktualne", "A" + (id + 1) + ":F" + (id + 1))[0];
}

function getRecipients(dataToReceipient) {
  var receipientList = []

  if (dataToReceipient[1] != "") {
    receipientList.push(dataToReceipient[1])
  }
  if (dataToReceipient[3] != "") {
    receipientList.push(dataToReceipient[3])
  }
  if (dataToReceipient[4] != "") {
    receipientList.push(dataToReceipient[4])
  }
  if (dataToReceipient[5] != "") {
    receipientList = receipientList.concat(dataToReceipient[5].split(","));
  }
  var defaultMail = SETTINGS[SETTINGS_DEFAULT_MAIL][0]
  if (defaultMail != "") {
    receipientList = receipientList.concat(defaultMail.split(","));
  }
  return receipientList.join(",")
}

function getMessageForYellow(settings, kom, kwa, numberYellowCard, checkPointName, dayToFix) {
  return SETTINGS[settings][0].replace("$kom", kom).replace("$kwa", kwa).replace("$nrCard", numberYellowCard).replace("$checkPointName", checkPointName).replace("$dayToFix", dayToFix);
}

function getMessage(settings, kom, kwa) {
  return SETTINGS[settings][0].replace("$kom", kom).replace("$kwa", kwa);
}

function internalLog(row) {
  if (TEST) {
    saveValues("logi-test", [row])
  } else {
    saveValues("logi", [row])
  }
}


function checkData() {
  var data = getRowsData()
  var columnsForCheckpoint1 = getColumnNumberForCheckpoint(1)
  var columnsForCheckpoint2 = getColumnNumberForCheckpoint(2)
  var columnsForCheckpoint3 = getColumnNumberForCheckpoint(3)
  var columnsForCheckpoint4 = getColumnNumberForCheckpoint(4)
  var output = []
  for (var i = 0; i < data.length; i++) {
    var row = data[i]
    var yellowCard1 = checkRowForCheckPoint(row, columnsForCheckpoint1)
    var yellowCard2 = checkRowForCheckPoint(row, columnsForCheckpoint2)
    var yellowCard3 = checkRowForCheckPoint(row, columnsForCheckpoint3)
    var yellowCard4 = checkRowForCheckPoint(row, columnsForCheckpoint4)
    output.push([i, row, yellowCard1, yellowCard2, yellowCard3, yellowCard4])
  }
  return output;
}

//return number of checkPoint that is not pass
function checkRowForCheckPoint(row, columnsForCheckpoint) {
  var currentDate = new Date()
  var output = []
  for (var c = 0; c < columnsForCheckpoint.length; c++) {
    var indexPlan = columnsForCheckpoint[c]
    var indexReal = columnsForCheckpoint[c] + 1
    var planDate = new Date(row[indexPlan]).setHours(23, 59, 59);
    var realDate = new Date(row[indexReal])
    if (planDate < currentDate) {
      if (!row[indexReal]) { // termin nie uzupełniony
        output.push(c)
      } else if (realDate > planDate) { //termin uzupełniony ale minął
        output.push(c)
      }
    }
  }
  return output
}


function getColumnNumberForCheckpoint(checkpoint) {
  var level = 0
  if (checkpoint == 1) { level = SETTINGS_CHECKPOINT_1 }
  if (checkpoint == 2) { level = SETTINGS_CHECKPOINT_2 }
  if (checkpoint == 3) { level = SETTINGS_CHECKPOINT_3 }
  if (checkpoint == 4) { level = SETTINGS_CHECKPOINT_4 }

  var columnsForCheckpointStr = SETTINGS[level][0]
  return columnsForCheckpointStr.split(",").map(function (value) { return charToInt(value) });
}


function getRowsData() {
  return getData("Dane", "A3:" + getLastRow("Dane"))
}

function checkIfTripIsCanceled(value) {
  if (value[charToInt('Y')]) {
    return true
  } else {
    return false
  }
}

function sendMail(recipient, subject, body) {
  if (!TEST) {
    MailApp.sendEmail({
      to: recipient,
      // to: "pawel.marczuk@zhr.pl",
      subject: subject,
      htmlBody: body.replace(/\n/g, '<br>')
    });
  }
}

function charToInt(str) {
  var value = 0;
  for (var i = 0; i < str.length; i++) {
    var char = str.charAt(i);
    var v = char.charCodeAt(0) - 'A'.charCodeAt(0)
    value += base(str.length - i, v)
  }
  return value
}

function base(position, value) {
  if (position == 1) {
    return value
  } else {
    return (value + 1) * 26
  }
}

function intToChar(int) {
  const code = 'A'.charCodeAt(0);
  return String.fromCharCode(code + int);
}

function getLastRow(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  return sheet.getLastRow();
}

function getLastCol(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  return sheet.getLastColumn();
}

function getData(sheetName, range) {
  return SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(range).getValues()
}

function getValue(sheetName, range) {
  return SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(range).getValue()
}

function saveValue(sheetName, range, value) {
  SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(range).setValue(value);
}

function saveValues(sheetName, data) {
  var lastRow = SpreadsheetApp.getActive().getSheetByName(sheetName).getLastRow();
  SpreadsheetApp.getActive().getSheetByName(sheetName).getRange(lastRow + 1, 1, data.length, data[0].length).setValues(data);
}
