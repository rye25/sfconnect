function getReport() {
  
  var reportId = '00O1N000008nFReUAM'; //SF Contacts Report ID
  
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Testing');
  
  var sfService = getSfService();
  var userProps = PropertiesService.getUserProperties();
  var props = userProps.getProperties();
  var name = getSfService().serviceName_;
  var obj = JSON.parse(props['oauth2.' + name]);
  var instanceUrl = obj.instance_url;
  var queryUrl = instanceUrl + "/services/data/v35.0/analytics/reports/" + reportId + "/instances";  // Actual request for report Data
  
  var options = {
    
  }; 
  
  var response = UrlFetchApp.fetch(queryUrl, options);


}