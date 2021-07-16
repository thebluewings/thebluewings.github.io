function doGet(e) {
  var type = e.parameter.type
  if (type == "state") {
    var cellphone = e.parameter.cellphone
    var state = getRandomString()
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("State")
    var lastRow = sheet.getLastRow()
    for (var i = 1; i<=lastRow; i++) {
      if (sheet.getRange(i, 1).getDisplayValue() == cellphone) {
        sheet.getRange(i, 2).setValue(state)
        return ContentService.createTextOutput(state)
      }
    }
    sheet.getRange(lastRow + 1, 1).setValue("'"+cellphone)
    sheet.getRange(lastRow + 1, 2).setValue(state)
    return ContentService.createTextOutput(state)
  }else if (type == "send") {
    var msg = e.parameter.msg
    var cellphone = e.parameter.cellphone
    var token = ""
    var tokenSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AccessToken")
    var lastRow = tokenSheet.getLastRow()
    for (var i = 1; i<=lastRow; i++) {
      if (tokenSheet.getRange(i, 1).getDisplayValue() == cellphone) {
        token = tokenSheet.getRange(i, 2).getDisplayValue()
      }
    }
    if (token == "") {
      return ContentService.createTextOutput("Can't find token")
    }
    var data = {
      "message" : msg
    }
    UrlFetchApp.fetch("https://notify-api.line.me/api/notify", {
      "headers": {
        "Content-Type" : "application/x-www-form-urlencoded",
        "Authorization" : "Bearer " + token
      },
      "method": "post",
      "payload": data
    })
    return ContentService.createTextOutput("訊息已發送")
  }
  return ContentService.createTextOutput("")
}

function getRandomString() {
  var sample = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
  var result = ""
  for (var i = 0; i<16; i++) {
    var index = Math.ceil(Math.random() * 100000) % sample.length
    result = result + sample[index]
  }
  return result
}

function doPost(e) {
  var code = e.parameter.code
  var state = e.parameter.state
  var cellphone = ""
  var stateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("State")
  var lastRow = stateSheet.getLastRow()
  for (var i = 1; i<=lastRow; i++) {
    if (stateSheet.getRange(i, 2).getDisplayValue() == state) {
      cellphone = stateSheet.getRange(i, 1).getDisplayValue()
      stateSheet.deleteRow(i)
    }
  }
  if (cellphone == "") {
    return ContentService.createTextOutput("Can't find cellphone")
  }
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Callback")
  var url = sheet.getRange("A1").getDisplayValue()
  var api = "https://notify-bot.line.me/oauth/token"
  
  var formData = {
    'grant_type': 'authorization_code',
    'code': code,
    'redirect_uri': url,
    'client_id': 'nQTAPEeeT6MPcZA15ghfV3',
    'client_secret': 'UFOITvEiSNc9bDPCxh8yDVz36HPalDxxM8F99EtiVTO'
  }
  var response = UrlFetchApp.fetch(api, {
    "headers": {
      "Content-Type" : "application/x-www-form-urlencoded"
    },
    "method": "post",
    "payload": formData
  })
  var json = JSON.parse(response.getContentText())
  var token = json.access_token
  
  var tokenSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AccessToken")
  var lastRow = tokenSheet.getLastRow()
  for (var i = 1; i<=lastRow; i++) {
    if (tokenSheet.getRange(i, 1).getDisplayValue() == cellphone) {
      tokenSheet.getRange(i, 2).setValue(token)
      return ContentService.createTextOutput("綁定完成")
    }
  }
  tokenSheet.getRange(lastRow + 1, 1).setValue("'"+cellphone)
  tokenSheet.getRange(lastRow + 1, 2).setValue(state)
  
  return ContentService.createTextOutput("綁定完成")
}
