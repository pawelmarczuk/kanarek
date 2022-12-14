
function testValue(value, expected) {
  if (value != expected) {
    console.error("value: " + value + " is not " + expected)
  }
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

function testGetRecipients() {
  var list = [];
  list.push("first")
  var str = "a,b,c"
  var out = list.concat(str.split(","));
  console.log(out)
  console.log(out.join(","))
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

function testSummary() {
  var data = getData("Dane", "A3:5")
  console.log(data)
}