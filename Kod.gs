var SETTINGS = getData("Ustawienia", "B:B")
var SETTINGS_MAIL_YELLOW_GRANT = 1
var SETTINGS_MAIL_YELLOW_PICK_UP = 2
var SETTINGS_MAIL_RED_GRANT = 3
var SETTINGS_MAIL_RED_PICK_UP = 4
var SETTINGS_MAIL_RED1_GRANT = 5
var SETTINGS_MAIL_RED1_PICK_UP = 6
var SETTINGS_CHECKPOINT_TITLE_1 = 7
var SETTINGS_CHECKPOINT_TITLE_2 = 8
var SETTINGS_CHECKPOINT_TITLE_3 = 9
var SETTINGS_CHECKPOINT_TITLE_4 = 10
var SETTINGS_CHECKPOINT_TITLE_5 = 11
var SETTINGS_DEFAULT_MAIL = 12

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

function runMain() {
  TEST = false
  main()
}

function runTest() {
  TEST = true
  WRITE = false
  main()
}

function main() {
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

    if (typeof id != "number") {
      console.error("ID is not a number")
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
      saveCard(id, idCard, color)
      sendMailForCard(id, idCard, color, numberYellowCard, dayToFix)
    }
  } else {
    if (typeof saveData.getMonth === 'function') {
      deleteCard(id, idCard, color)
      sendMailForTakeBack(id, idCard, color)
    }
  }
}

function checkMoreThreeCard(actual, saveData, id) {
  if (
    (typeof actual == "number" && actual > 3)
  ) {
    if (typeof saveData.getMonth !== 'function') {
      saveRedCard(id)
      sendMailForRedCard(id)
    }
  } else {
    if (typeof saveData.getMonth === 'function') {
      deleteRedCard(id)
      sendMailForTakeBackRed(id)
    }
  }
}

function saveCard(id, idCard, color) {
  console.log("save " + color + " " + id + " " + (idCard + 1))
  if (WRITE) {
    saveValue('stan kartek', getLocationForCard(id, idCard, color), new Date())
  }
}

function saveRedCard(id) {
  console.log("saveRED " + id)
  if (WRITE) {
    saveValue('stan kartek', getLocationForRedCard(id), new Date())
  }
}

function deleteCard(id, idCard, color) {
  console.log("delete " + color + " " + id + " " + (idCard + 1))
  if (WRITE) {
    saveValue('stan kartek', getLocationForCard(id, idCard, color), "")
  }
}

function deleteRedCard(id) {
  console.log("deleteRed " + id)
  if (WRITE) {
    saveValue('stan kartek', getLocationForRedCard(id), "")
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

function sendMailForCard(id, idCard, color, numberYellowCard, dayToFix) {
  var level = idCard + 1;
  var mailData = getMailData(id)
  var checkPointName = CHECKPOINT_TITLE[level]
  var subject = ""
  var message = ""
  if (color == YELLOW) {
    var dtf = Utilities.formatDate(dayToFix, "GMT+1", "yyyy-MM-dd");
    message = getMessageForYellow(SETTINGS_MAIL_YELLOW_GRANT, mailData.komendant, mailData.kwatermitrz, numberYellowCard, checkPointName, dtf)
    subject = "HAZ 2023 " + mailData.komendant + "/" + mailData.kwatermitrz + " [żółta kartka - termin nr " + level + "]"
  } else {
    message = getMessageForYellow(SETTINGS_MAIL_RED1_GRANT, mailData.komendant, mailData.kwatermitrz, numberYellowCard, checkPointName, "")
    subject = "HAZ 2023 " + mailData.komendant + "/" + mailData.kwatermitrz + " [czerwona kartka - termin nr " + level + "]"
  }

  sendMail(mailData.recipeint, subject, message)
  internalLog([new Date(), id, "przyznanie", colorToText(color), (idCard + 1), message, mailData.recipeint])
}

function sendMailForTakeBack(id, idCard, color) {
  var level = idCard + 1;
  var mailData = getMailData(id)
  var subject = ""
  var message = ""
  if (color == YELLOW) {
    var message = getMessage(SETTINGS_MAIL_YELLOW_PICK_UP, mailData.komendant, mailData.kwatermitrz)
    subject = "HAZ 2023 " + mailData.komendant + "/" + mailData.kwatermitrz + " [żółta kartka - termin nr " + level + "]"
  } else {
    var message = getMessage(SETTINGS_MAIL_RED1_PICK_UP, mailData.komendant, mailData.kwatermitrz)
    subject = "HAZ 2023 " + mailData.komendant + "/" + mailData.kwatermitrz + " [czerwona kartka - termin nr " + level + "]"
  }

  sendMail(mailData.recipeint, subject, message)
  internalLog([new Date(), id, "odebranie", colorToText(color), (idCard + 1), message, mailData.recipeint])
}

function sendMailForRedCard(id, idCard) {
  var mailData = getMailData(id)
  var subject = "HAZ 2023 " + mailData.komendant + "/" + mailData.kwatermitrz + " [Czerwona kartka]"
  var message = getMessage(SETTINGS_MAIL_RED_GRANT, mailData.komendant, mailData.kwatermitrz)
  sendMail(mailData.recipeint, subject, message)
  internalLog([new Date(), id, "przyznanie", colorToText(RED), (idCard + 1), message, mailData.recipeint])
}

function sendMailForTakeBackRed(id, idCard) {
  var mailData = getMailData(id)
  var subject = "HAZ 2023 " + mailData.komendant + "/" + mailData.kwatermitrz + " unieważnienie [Czerwona kartka]"
  var message = getMessage(SETTINGS_MAIL_RED_PICK_UP, mailData.komendant, mailData.kwatermitrz)
  sendMail(mailData.recipeint, subject, message)
  internalLog([new Date(), id, "przyznanie", colorToText(RED), (idCard + 1), message, mailData.recipeint])
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

    if (typeof dataRow[id] != "number") { continue }

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

  var data = getData("stan kartek", "A2:L" + (rowsLength + 1));
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

function dataToReceipientRaw(id) {
  if (typeof id != "number")
    return ["", "", "", "", "", ""];

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
  return SETTINGS[settings][0].replace("$kom", kom).replace("$kwa", kwa).replace("$allYellowCard", numberYellowCard).replace("$checkPointName", checkPointName).replace("$dayToFix", dayToFix);
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
