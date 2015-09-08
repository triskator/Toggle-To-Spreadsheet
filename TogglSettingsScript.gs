
/**
*Load stored settings for settings popup
*/
function getToggleSettings() {
  var settings = new Array();
  var apiKey, useragent, workspace;
  apiKey =  userProperties.getProperty("TOGGL_API_KEY");
  useragent = userProperties.getProperty("USER_AGENT");
  workspace = userProperties.getProperty("WORKSPACE_ID");
  
  settings.push({"name":"toggl_api_key", "value": apiKey});
  settings.push({"name":"user_agent", "value": useragent});
  settings.push({"name":"workspace_id", "value": workspace});
  
  return settings;
}

/**
*Save settings from popup and test connection if data is valid.
*/
function setToggleSettings(values){
  for (var i = 0; i < values.length; i++) {
    var id = values[i]["name"];
    id = id.toUpperCase();
    var value = values[i]["value"];
    userProperties.setProperty(id, value);
  }
  tryConnection();
}

/**
* Tries to get some data from Toggl. result should have some data, if it's null, 
* it didn't connect at all, so it's failed.
* According response it will disable/enable sidebar.
*/
function tryConnection(){
  var ui = SpreadsheetApp.getUi();
  var start, workspace, useragent, now = new Date();
  now.setHours(0);
  now.setMinutes(0);
  now.setSeconds(0);
  useragent = userProperties.getProperty("USER_AGENT");
  workspace = userProperties.getProperty("WORKSPACE_ID");
  
  var data = callToggl("https://toggl.com/reports/api/v2/details", {
    "user_agent": useragent,
    "workspace_id": workspace, 
    "since": toYMD(now),
    "until": toYMD(now),
    "page" : 1
  });
 
  if(data !== null){
    unlockSidebar();
    ui.alert("Spojení proběhlo úspěšně."); 
    showSidebar();
  } else {
    lockSidebar();
    ui.alert('Spojení se nezdařilo, zkontrolujte nastavení aplikace.'); 
    setTogglApi();
  }
}

/**
* Empty function, it has to call something for "callback" to close the sidebar.
* Right now only way to close sidebar is to call init of sidebar and destruction inside it right away it's created.
* So empty sidebar is created and then closed.
*/
function justCloseDamnSidebar() {

}

/**
* Function to display Settings popup/modal dialog.
*/
function setTogglApi() {
  var html = HtmlService.createHtmlOutputFromFile('TogglSettingsDialog')
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setWidth(400)
  .setHeight(300);
   html.append(HtmlService.createHtmlOutputFromFile('Stylesheet').getContent());
  //append JS file
  html.append(HtmlService.createHtmlOutputFromFile('TogglSettingsJS').getContent());
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showModalDialog(html, 'Toggl Settings');
}

/**
* Will replace dialog with en empty one and then closes itself. Also removes toggl from menu.
*/
function lockSidebar(){
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem('Set Toggle API', 'setTogglApi')
  .addToUi();
 var ui = HtmlService.createTemplateFromFile('CloseSidebar')
    .evaluate()
    .setTitle(SIDEBAR_TITLE);
    SpreadsheetApp.getUi().showSidebar(ui);
}

/**
* Opens toggl sidebar and enables items within menu.
*/
function unlockSidebar(){
  SpreadsheetApp.getUi()
  .createAddonMenu()
  .addItem('Show sidebar', 'showSidebar')
  .addItem('Set Toggle Settings', 'setTogglApi')
  .addToUi();
}