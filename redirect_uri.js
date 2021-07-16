// LINE Notify Redirect_Uri：https://script.google.com/macros/s/AKfycbzdYblEGSFew9yw7w_vKOnzWuMwucDcePF45VhMOpDKJhMrOApW8M-xRh7FU4Ivclf0/exec
function doGet(e) {
  //這裡的都是不一定要用Get的功能，因為用Google AppScript無法準備多支介接只好出此下策，正常應該分成好幾支API處理
  var type = e.parameter.type
  if (type == "state") {
    //隨機產生一組state字串並與傳入的cellphone綁定寫到State資料表
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
    //根據傳入的cellphone找出AccessToken並發送訊息
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
  //↓↓利用state取出存在State表中的cellphone↓↓
  var stateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("State")
  var lastRow = stateSheet.getLastRow()
  for (var i = 1; i<=lastRow; i++) {
    if (stateSheet.getRange(i, 2).getDisplayValue() == state) {
      //↓↓透過state找出對應的cellphone
      cellphone = stateSheet.getRange(i, 1).getDisplayValue()
      //↓↓刪掉該行state的資料↓↓
      stateSheet.deleteRow(i)
    }
  }
  if (cellphone == "") {
    //找不到cellphone
    return ContentService.createTextOutput("Can't find cellphone")
  }
  //↓↓取出redirect_uri，因為無法直接寫在程式中，只好另外讀取↓↓
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Callback")
  var url = sheet.getRange("A1").getDisplayValue()
  //↓↓呼叫取得AccessToken的LINE Notify API
  var api = "https://notify-bot.line.me/oauth/token"
  var formData = {
    'grant_type': 'authorization_code',
    'code': code, //LINE API回傳的code，用來取得AccessToken，無法論成與否都只能使用一次
    'redirect_uri': url, //如果這邊帶的redirect_uri與後台設定的不一致會失敗
    'client_id': 'nQTAPEeeT6MPcZA15ghfV3', //註冊服務時會取得
    'client_secret': 'UFOITvEiSNc9bDPCxh8yDVz36HPalDxxM8F99EtiVTO' //註冊服務時取得
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
  //↓↓將AccessToken與手機寫回資料表，以便透過手機號碼取得AccessToken↓↓
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
  //回傳response
  return ContentService.createTextOutput("綁定完成")
}
