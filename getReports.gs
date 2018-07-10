function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Load Deals by Originator data', functionName: 'dealsByOriginator'},
    {name: 'Load Total Activities data', functionName: 'totalActivities'}
  ];
  spreadsheet.addMenu('Salesforce', menuItems);
}

function dealsByOriginator() {
  var reportId = '00O1N000008ny0P';
  var sheetName = 'Deal Data'; 
  
  var sfService = getSfService();
  var userProps = PropertiesService.getUserProperties();
  var props = userProps.getProperties();
  var name = getSfService().serviceName_;
  var obj = JSON.parse(props['oauth2.' + name]);
  var instanceUrl = obj.instance_url;
  var queryUrl = instanceUrl + "/services/data/v35.0/analytics/reports/" + reportId + "?includeDetails=true";  // Actual request for report Data
  var response = UrlFetchApp.fetch(queryUrl, { method : "POST", headers : { "Authorization" : "OAuth "+ sfService.getAccessToken()} });
  var contentText = response.getContentText();
  var queryResult = JSON.parse(contentText);
  
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  
  var answer = queryResult.factMap["T!T"].rows;  // assumes tabular report
  var headers = queryResult.reportExtendedMetadata.detailColumnInfo;
  var headname = queryResult.reportMetadata.detailColumns;
  var myArray = [];
  var tempArray = [];
  for (i = 0 ; i < headname.length ; i++) {
    tempArray.push(headers[headname[i]].label);
  }
  myArray.push(tempArray);
  
  for (i = 0 ; i < answer.length ; i++ ) {
    var tempArray = [];
    function getData(element,index,array) {
      tempArray.push(array[index].label)
    }
    answer[i].dataCells.forEach(getData);
    myArray.push(tempArray);
  }
  
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow !== 0 || lastColumn !== 0) {
    sheet.getRange(1,1,lastRow, myArray[0].length).clearContent();
  }
  sheet.getRange(1,1, myArray.length, myArray[0].length).setValues(myArray);
}

function totalActivities() {
  var reportId = '00O1N000008ny0U';
  var sheetName = 'Activity Data'; 
  
  var sfService = getSfService();
  var userProps = PropertiesService.getUserProperties();
  var props = userProps.getProperties();
  var name = getSfService().serviceName_;
  var obj = JSON.parse(props['oauth2.' + name]);
  var instanceUrl = obj.instance_url;
  var queryUrl = instanceUrl + "/services/data/v35.0/analytics/reports/" + reportId + "?includeDetails=true";  // Actual request for report Data
  var response = UrlFetchApp.fetch(queryUrl, { method : "POST", headers : { "Authorization" : "OAuth "+ sfService.getAccessToken()} });
  var contentText = response.getContentText();
  var queryResult = JSON.parse(contentText);
  
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  
  var answer = queryResult.factMap["T!T"].rows;  // assumes tabular report
  var headers = queryResult.reportExtendedMetadata.detailColumnInfo;
  var headname = queryResult.reportMetadata.detailColumns;
  var myArray = [];
  var tempArray = [];
  for (i = 0 ; i < headname.length ; i++) {
    tempArray.push(headers[headname[i]].label);
  }
  myArray.push(tempArray);
  
  for (i = 0 ; i < answer.length ; i++ ) {
    var tempArray = [];
    function getData(element,index,array) {
      tempArray.push(array[index].label)
    }
    answer[i].dataCells.forEach(getData);
    myArray.push(tempArray);
  }
  
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow !== 0 || lastColumn !== 0) {
    sheet.getRange(1,1,lastRow, myArray[0].length).clearContent();
  }
  sheet.getRange(1,1, myArray.length, myArray[0].length).setValues(myArray);
}
