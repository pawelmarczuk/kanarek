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
var SETTINGS_MAIL_ORANGE_GRANT = 13
var SETTINGS_MAIL_GREEN_GRANT = 14

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
var ORANGE = "orange"
var GREEN = "green"

var DAYS_FOR_NOTIFICATION = 3
var DAYS_FOR_GREENCARD_TO_CHECK_IN_PAST = 7

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

function runTestOneRow() {
  TEST = true
  WRITE = false

  var row = 1;

  var summary = getSummaryData()
  var saveData = getSaveData(summary.length)

  var rowSummary = summary[row];
  var rowSave = saveData[row];
  var id = rowSummary.id

  checkRow(rowSummary, rowSave, id)
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
    checkRow(rowSummary, rowSave, id)
  }
}

function checkRow(rowSummary, rowSave, id) {
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
  //notification
  for (var c = 0; c < rowSummary.notification.length; c++) {
    var orange = rowSummary.notification[c]
    var save = rowSave.term_oragne[c]
    checkOrangeCard(orange, save, id, c)
  }
  //green
  for (var c = 0; c < rowSummary.notification.length; c++) {
    var green = rowSummary.notification[c]
    var save = rowSave.term_green[c]
    checkGreenCard(green, rowSummary.term_to_fix[c], save, id, c)
  }
  checkMoreThreeCard(rowSummary.three_yellow, rowSave.three_yellow, id)
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

function checkOrangeCard(actual, saveData, id, idCard) {
  var term_plan = actual[0]
  var term_real = actual[1]

  if (
    term_real == "" && typeof saveData.getMonth !== 'function'
  ) {
    //check if there is X dat to send notification
    var diffTime = calcDiffTimeInDays(new Date(), term_plan)
    if (diffTime >= 0 && diffTime <= DAYS_FOR_NOTIFICATION) {
      if (sendMailForOrangeCard(id, idCard, term_plan)) {
        saveOrangeCard(id, idCard)
      }
    }
  }
}

function checkGreenCard(actual, last_plan_term, saveData, id, idCard) {
  var term_plan = actual[0]
  var term_real = actual[1]

  if (
    typeof term_real.getMonth === 'function' && typeof saveData.getMonth !== 'function'
  ) {
    var diffTime = calcDiffTimeInDays(term_real, last_plan_term)
    if (diffTime >= 0) {
      var diffTimeToCurrentDate = calcDiffTimeInDays(term_real, new Date())
      if (diffTimeToCurrentDate >= 0 && diffTimeToCurrentDate <= DAYS_FOR_GREENCARD_TO_CHECK_IN_PAST) {
        if (sendMailForGreenCard(id, idCard)) {
          saveGreenCard(id, idCard)
        }
      }
    }
  }
}

function calcDiffTimeInDays(term1, term2) {
  var before = createnewDateWithoutHours(term1)
  var after = createnewDateWithoutHours(term2)
  let difference = after.getTime() - before.getTime();
  return Math.ceil(difference / (1000 * 3600 * 24));
}

function createnewDateWithoutHours(original) {
  var date = new Date(original)
  date.setHours(0)
  date.setMinutes(0)
  date.setSeconds(0)
  return date;
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

function saveOrangeCard(id, idCard) {
  console.log("save orange " + id + " " + (idCard + 1))
  if (WRITE) {
    saveValue('stan kartek', getLocationForCard(id, idCard, ORANGE), new Date())
  }
}

function saveGreenCard(id, idCard) {
  console.log("save green " + id + " " + (idCard + 1))
  if (WRITE) {
    saveValue('stan kartek', getLocationForCard(id, idCard, GREEN), new Date())
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
  } else if (color == ORANGE) {
    return intToChar(idCard + 12) + (id + 1)
  } else if (color == GREEN) {
    return intToChar(idCard + 17) + (id + 1)
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

function sendMailForOrangeCard(id, idCard, termPlan) {
  var level = idCard + 1;
  var mailData = getMailData(id)
  var checkPointName = CHECKPOINT_TITLE[level]
  var tp = Utilities.formatDate(termPlan, "GMT+1", "yyyy-MM-dd");

  var subject = "HAZ 2023 " + mailData.komendant + "/" + mailData.kwatermitrz + " [przypomnienie - termin nr " + level + "]"
  var message = getMessageForOrange(SETTINGS_MAIL_ORANGE_GRANT, mailData.komendant, mailData.kwatermitrz, checkPointName, tp)

  if (message != "") {
    sendMail(mailData.recipeint, subject, message)
    internalLog([new Date(), id, "przyznanie", colorToText(ORANGE), (idCard + 1), message, mailData.recipeint])
    return true
  } else {
    console.error("pusta wiadomość")
    internalLog([new Date(), id, "BŁĄD", colorToText(ORANGE), (idCard + 1), "", ""])
    return false
  }
}

function sendMailForGreenCard(id, idCard) {
  var level = idCard + 1;
  var mailData = getMailData(id)
  var checkPointName = CHECKPOINT_TITLE[level]

  var subject = "HAZ 2023 " + mailData.komendant + "/" + mailData.kwatermitrz + " [zielona kartka - termin nr " + level + "]"
  var message = getMessageForGreen(SETTINGS_MAIL_GREEN_GRANT, mailData.komendant, mailData.kwatermitrz, checkPointName)

  if (message != "") {
    sendMail(mailData.recipeint, subject, message)
    internalLog([new Date(), id, "przyznanie", colorToText(GREEN), (idCard + 1), message, mailData.recipeint])
    return true
  } else {
    console.error("pusta wiadomość")
    internalLog([new Date(), id, "BŁĄD", colorToText(GREEN), (idCard + 1), "", ""])
    return false
  }
}

function colorToText(color) {
  if (color == YELLOW) {
    return "żółta"
  } else if (color == RED) {
    return "czerwona"
  } else if (color == ORANGE) {
    return "pomarańczowa"
  } else if (color == GREEN) {
    return "zielona"
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

  var noti1A = charToInt("Z")
  var noti1B = charToInt("AA")
  var noti2A = charToInt("AD")
  var noti2B = charToInt("AE")
  var noti3A = charToInt("AH")
  var noti3B = charToInt("AI")
  var noti4A = charToInt("AL")
  var noti4B = charToInt("AM")
  var noti5A = charToInt("AP")
  var noti5B = charToInt("AQ")

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
      notification: [
        [dataRow[noti1A], dataRow[noti1B]],
        [dataRow[noti2A], dataRow[noti2B]],
        [dataRow[noti3A], dataRow[noti3B]],
        [dataRow[noti4A], dataRow[noti4B]],
        [dataRow[noti5A], dataRow[noti5B]],
      ]
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

  var oragne1 = charToInt("M")
  var oragne2 = charToInt("N")
  var oragne3 = charToInt("O")
  var oragne4 = charToInt("P")
  var oragne5 = charToInt("Q")

  var green1 = charToInt("R")
  var green2 = charToInt("S")
  var green3 = charToInt("T")
  var green4 = charToInt("U")
  var green5 = charToInt("V")

  var data = getData("stan kartek", "A2:V" + (rowsLength + 1));
  var output = []
  for (var i = 0; i < data.length; i++) {
    var dataRow = data[i]
    var row = {
      term_yellow: [dataRow[yellow1], dataRow[yellow2], dataRow[yellow3], dataRow[yellow4], dataRow[yellow5]],
      term_red: [dataRow[red1], dataRow[red2], dataRow[red3], dataRow[red4], dataRow[red5]],
      three_yellow: dataRow[threeYellow],
      term_oragne: [dataRow[oragne1], dataRow[oragne2], dataRow[oragne3], dataRow[oragne4], dataRow[oragne5]],
      term_green: [dataRow[green1], dataRow[green2], dataRow[green3], dataRow[green4], dataRow[green5]]
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

function getMessageForOrange(settings, kom, kwa, checkPointName, termPlan) {
  return SETTINGS[settings][0].replace("$kom", kom).replace("$kwa", kwa).replace("$checkPointName", checkPointName).replace("$date", termPlan);
}

function getMessageForGreen(settings, kom, kwa, checkPointName) {
  return SETTINGS[settings][0].replace("$kom", kom).replace("$kwa", kwa).replace("$checkPointName", checkPointName)
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
