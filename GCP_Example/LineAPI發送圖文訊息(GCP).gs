// 本腳本可以抓存放於雲端硬碟中的試算表特定的Sheet中的欄位資料列出來，並根據過去指定時間內是否有資料異動，
// 若有異動才會觸發資料彙整並且指定表格範圍產生圖片檔(需指定driver folder ID)，再藉由已知的公開連結範圍作為LINE所需正確格式內容，才可正常發送。
//範例會是這樣

//試算表: Sheet名稱
//連結: Sheet所在包含gid連結
//欄位1欄位內容
//欄位2欄位內容
//欄位3欄位內容(Hyperlink)

//欄位4欄位內容
//欄位5欄位內容(有做自動進位)

//最後編輯時間: "抓取該試算表檔案的修改時間"

function checkUpdates() {
  var spreadsheetId = ''; // 替換為你的試算表檔案ID
  var sheetName = ''; // 替換為你的試算表名稱
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);

  if (!sheet) {
    Logger.log('找不到工作表: ' + sheetName);
    return;
  }

  var range = sheet.getRange('A1:H9'); // 取得範圍A1:H9的資料
  var data = range.getValues(); // 取得範圍內的資料
  var imageUrl = generateTableImage(data, sheetName); // 生成範圍內資料的圖片並取得圖片URL
  
  // 等待圖片上傳完畢
  Utilities.sleep(1000); // 等待1秒鐘，可以根據需要調整時間

  var file = DriveApp.getFileById(spreadsheetId);
  var lastEdited = file.getLastUpdated(); // 取得最後編輯時間
  var formattedLastEdited = Utilities.formatDate(lastEdited, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"); // 格式化編輯時間
  var sheetUrl = file.getUrl() + "#gid=" + sheet.getSheetId(); // 生成工作表的URL
  
  var rangeG = sheet.getRange('G2:G7'); // G欄位範圍
  var valuesG = rangeG.getValues(); // 取得G欄位範圍內的資料
  
  var rangeH = sheet.getRange('H2:H7'); // H欄位範圍
  var valuesH = rangeH.getValues(); // 取得H欄位範圍內的資料
  var urlsH = getHyperlinks(rangeH); // 取得H欄位範圍內的超連結
  
  var dateRange1 = sheet.getRange('A2:A7'); // A2到A7範圍
  var dateRange2 = sheet.getRange('A9');    // A9範圍
  var dateValues1 = dateRange1.getValues(); // 取得A2到A7範圍內的資料
  var dateValues2 = dateRange2.getValues(); // 取得A9範圍內的資料
  var dateValues = dateValues1.concat(dateValues2); // 合併日期資料
  
  var columnHeaderA1 = sheet.getRange('A1').getValue(); // 取得A1的欄位名稱
  
  var columnHeaderG = sheet.getRange('G1').getValue(); // 取得G1的欄位名稱
  
  var columnHeaderH = sheet.getRange('H1').getValue(); // 取得H1的欄位名稱

  var f8Header = sheet.getRange('F8').getValue(); // 取得F8的欄位名稱
  var f9Value = sheet.getRange('F9').getValue(); // 取得F9的值
  var g8Header = sheet.getRange('G8').getValue(); // 取得G8的欄位名稱
  var g9Value = sheet.getRange('G9').getValue(); // 取得G9的值
  var newg9Value = Math.round(g9Value); // 將G9的值四捨五入
  
  var scriptProperties = PropertiesService.getScriptProperties();
  var savedValuesG = JSON.parse(scriptProperties.getProperty('SAVED_VALUES_G') || '[]'); // 取得已保存的G欄位值
  var savedValuesH = JSON.parse(scriptProperties.getProperty('SAVED_VALUES_H') || '[]'); // 取得已保存的H欄位值
  var savedUrlsH = JSON.parse(scriptProperties.getProperty('SAVED_URLS_H') || '[]'); // 取得已保存的H欄位超連結

  var notifications = []; // 用於儲存通知訊息
  
  if (savedValuesG.length === 0 || savedValuesH.length === 0 || savedUrlsH.length === 0) {
    // 如果沒有保存的值，則保存初始值
    scriptProperties.setProperty('SAVED_VALUES_G', JSON.stringify(valuesG));
    scriptProperties.setProperty('SAVED_VALUES_H', JSON.stringify(valuesH));
    scriptProperties.setProperty('SAVED_URLS_H', JSON.stringify(urlsH));
  } else {
    for (var i = 0; i < valuesG.length; i++) { // 檢查變更
      var newValueG = Math.round(valuesG[i][0]); // 將G欄位的值四捨五入
      var newValueH = valuesH[i][0];
      var newUrlH = urlsH[i];

      if (valuesG[i][0] != savedValuesG[i][0] || valuesH[i][0] != savedValuesH[i][0] || urlsH[i] != savedUrlsH[i]) { // 檢查新值與原值是否不同
        var message = "\n" + "試算表: " + sheet.getName() + "\n" +
                      "連結: " + sheetUrl + "\n" +
                      "【" + columnHeaderA1 + "】" + dateValues[i][0] + "\n" +
                      "【" + columnHeaderG + "】" + newValueG;
        if (valuesH[i][0] != savedValuesH[i][0] || urlsH[i] != savedUrlsH[i]) {
          var messageH = "【" + columnHeaderH + "】" + (newUrlH ? newUrlH : newValueH); // 顯示超連結或文字
          message += "\n" + messageH;
        }
        message += "\n\n" + "【" + f8Header + "】" + f9Value;
        message += "\n" + "【" + g8Header + "】" + newg9Value;
        message += "\n\n最後編輯時間: " + formattedLastEdited;
        notifications.push(message);
      }
    }
    var messagetable = "快照：";
    // 這邊是如果上面有監測到資料異動時，並且將異動的資料存入清單後將自動觸發發送條件
    if (notifications.length > 0) {
      Logger.log('發送通知：' + notifications.join("\n\n"));
      // sendLineNotification(notifications.join("\n\n")); // 這個是用LINE Notify帳號發送而非API
      sendLineGroupMessage(notifications.join("\n\n")); // 發送通知到LINE群組
      sendLineGroupImage(messagetable, imageUrl); // 發送快照圖片到LINE群組
  
      // 更新保存的值
      scriptProperties.setProperty('SAVED_VALUES_G', JSON.stringify(valuesG));
      scriptProperties.setProperty('SAVED_VALUES_H', JSON.stringify(valuesH));
      scriptProperties.setProperty('SAVED_URLS_H', JSON.stringify(urlsH));
    } else {
      Logger.log('沒有符合條件的變更');
    }
  }
}

function getHyperlinks(range) {
  var richTextValues = range.getRichTextValues();
  return richTextValues.map(row => {
    return row.map(cell => {
      var url = cell.getLinkUrl();
      return url ? url : "";
    });
  }).flat();
}

//這個函數的用法是利用LINE Notify 去發通知，因為後來想說用API比較方便也能夠個人化
/**
function sendLineNotification(message) {
  // var token = ''; // 替換為你的 Line Notify Token 權杖
  
  var options = {
    'method': 'post',
    'headers': {
      'Authorization': 'Bearer ' + token
    },
    'payload': {
      'message': message
    }
  };
  Logger.log("message" + message);
  UrlFetchApp.fetch('https://notify-api.line.me/api/notify', options);
}
/** */


function onEdit(e) {
  var range = e.range;
  var sheet = e.source.getActiveSheet();
  var editedColumn = range.getColumn();
  
  // 假設我們監控的是第7欄（G欄）
  if (editedColumn == 7 && range.getRow() <= 9) {
    var user = Session.getActiveUser().getEmail();
    var time = new Date();
    var formattedTime = Utilities.formatDate(time, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
    Logger.log('記錄編輯：欄位 G' + range.getRow() + '，編輯者：' + user + '，時間：' + formattedTime);
    range.setNote("編輯者: " + user + "\n時間: " + formattedTime);
  }
}

function generateTableImage(data, sheetName) {
  var dataTable = Charts.newDataTable();
  
  // 假設第一行是表頭
  var headers = data[0];
  headers.forEach(header => {
    dataTable.addColumn(Charts.ColumnType.STRING, header);
  });
  
  // 資料處理
  for (var i = 1; i < data.length; i++) {
    var row = data[i].map(String); // 將所有資料轉換為字串類型
    dataTable.addRow(row);
  }

  var chartBuilder = Charts.newTableChart()
    .setDimensions(800, 400)
    .setDataTable(dataTable);
  var chart = chartBuilder.build();
  var image = chart.getAs('image/png');
  
  // 取得當前時間戳用於產生圖片上傳時是帶入Sheet名+時間.png
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  var fileName = sheetName + "_" + timestamp + ".png";
  
  // 取得資料夾（用您的資料夾 ID 替換 'yourfolderid'）
  var folder = DriveApp.getFolderById('yourfolderid');
  var file = folder.createFile(image).setName(fileName);

  // 設定共用權限
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  // 取得文件 ID 並生成正確的下載連結
  var fileId = file.getId();
  var imageUrl = "https://drive.usercontent.google.com/download?id=" + fileId + "&export=view";
  return imageUrl; // 返回正確的公開連結
}

function sendLineGroupImage(messagetable, imageUrl) {
  var accessToken = ''; //這邊替換成你的LINE API BOT Channel access token
  var userId = '';//這邊是你要發給目標對象之ID (用戶或群組)

  var url = 'https://api.line.me/v2/bot/message/push';
  var payload = {
    "to": userId,
    "messages": [
      {
        "type": "text",
        "text": messagetable
      },
      {
        "type": "image",
        "originalContentUrl": imageUrl,
        "previewImageUrl": imageUrl
      }
    ]
  };

  var options = {
    'method': 'post',
    'headers': {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + accessToken
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  var response = UrlFetchApp.fetch(url, options);
}

function sendLineGroupMessage(message) {
  var accessToken = ''; // 替換為你的 Channel Access Token 
  var userId = ''; // 替換為你的用戶 ID或群組ID

  var url = 'https://api.line.me/v2/bot/message/push';
  var payload = {
    "to": userId,
    "messages": [
      {
        "type": "text",
        "text": message
      }
    ]
  };

  var options = {
    'method': 'post',
    'headers': {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + accessToken
    },
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  var response = UrlFetchApp.fetch(url, options);
}
