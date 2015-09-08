/**
* mar@OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
*/

var SIDEBAR_TITLE = "Toggl výkazy";
var userProperties = PropertiesService.getUserProperties();
var scriptProperties = PropertiesService.getScriptProperties();

scriptProperties.setProperty("VERSION", "1.3");

var page = 1;
var total_pages = page;
var last_position = 0;

function getAuthDigest(){
  return  "Basic " + Utilities.base64Encode(userProperties.getProperty("TOGGL_API_KEY") + ":api_token");
}

function toYMD(date) {
  var year, month, day;
  year = String(date.getFullYear());
  month = String(date.getMonth() + 1);
  if (month.length == 1) {
    month = "0" + month;
  }
  day = String(date.getDate());
  if (day.length == 1) {
    day = "0" + day;
  }
  return year + "-" + month + "-" + day;
}

function toQueryString(obj) {
  var parts = [];
  for (var i in obj) {
    if (obj.hasOwnProperty(i)) {
      parts.push(encodeURIComponent(i) + "=" + encodeURIComponent(obj[i]));
    }
  }
  return parts.join("&");
}

/**
* Call Toggl API
*
* @param {String} url
* @param {Object} params
* @return {Object}
*/
function callToggl(url, params) {
  var data, json, response;
  var ui = SpreadsheetApp.getUi();
  if (params) {
    url = url + "?" + toQueryString(params);
  }
  try
  {
    response = UrlFetchApp.fetch(url, {
      method: "GET",
      headers: {
        "Authorization": getAuthDigest()
      },
      contentType: "application/json"
    });
    json = response.getContentText();    
    data = JSON.parse(json);
    //ui.alert(JSON.stringify(data));
  }
  catch(err)
  {
    //ui.alert(err);
    Logger.log(err);
    //Error happend, probably 403 because of wrong api token
    return null;
  }
  return data;
}

/**
* Adds a custom menu with items to show the sidebar and dialog.
*
* @param {Object} e The event parameter for a simple onOpen trigger.
*/
function onOpen(e) {
  tryConnection();
}

/**
* Runs when the add-on is installed; calls onOpen() to ensure menu creation and
* any other initializion work is done immediately.
*
* @param {Object} e The event parameter for a simple onInstall trigger.
*/
function onInstall(e) {
  onOpen(e);
}

/**
* Opens a sidebar. The sidebar structure is described in the Sidebar.html
* project file.
*/
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
  .evaluate()
  .setTitle(SIDEBAR_TITLE);
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
* Retrieve list of clients from Toggl (cached)
*
* @return {Array} list of clients
*/
function getClients() {
  var cache, cached, data, cacheName, useragent, workspace, version;
  useragent = userProperties.getProperty("USER_AGENT");
  workspace = userProperties.getProperty("WORKSPACE_ID");
  version = scriptProperties.getProperty("VERSION");
  cache = CacheService.getScriptCache();
  cacheName = "toggl-clients" + getAuthDigest() + useragent + workspace + version; //if someone swith accounts, use that kind of cache name
  cached = cache.get(cacheName);
  
  if (cached != null && "" !== cached ) {
    return JSON.parse(cached);
  }
  
  data = callToggl("https://www.toggl.com/api/v8/clients");
  
  cache.put(cacheName, JSON.stringify(data), 60*60*24); // 24 hours cache
  
  return data;
}

function getTimeEntries( clientId, timeRange, clientName, toSheets ) {
  var start, end, day, now = new Date(), monday, useragent, workspace;
  now.setHours(0);
  now.setMinutes(0);
  now.setSeconds(0);
  
  var ui = SpreadsheetApp.getUi();
  
  
  if ( 'lastweek' == timeRange ) {
    day = now.getDay() || 7;
    if ( day !== 1 ) {
      now.setHours(-24 * (day - 1));
    }
    
    start = new Date(now.setDate(now.getDate()-7));
    end = new Date(now.setDate(now.getDate()+6));
  }
  
  
  if ( 'lastmonth' == timeRange ) {
    start = new Date(now.setDate(1));
    start.setMonth(start.getMonth()-1);
    end = new Date(now.setDate(0));
  }
  
  if ( 'thisweek' == timeRange ) {
    day = now.getDay() || 7;
    if ( day !== 1 ) {
      now.setHours(-24 * (day - 1));
    }
    start = new Date(now.setDate(now.getDate()));
    end = new Date(now.setDate(now.getDate()+6));
  }
  
  if ( 'thismonth' == timeRange ) {
    start = new Date(now.setDate(1));
    start.setMonth(start.getMonth());
    
    end = new Date(now);
    end.setMonth(end.getMonth()+1);
    end.setDate(0);
  }
  
  
  useragent = userProperties.getProperty("USER_AGENT");
  workspace = userProperties.getProperty("WORKSPACE_ID");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //remove old sheets
  var sheets = ss.getSheets();
  for(var i = 1; i < sheets.length; i++){
    ss.deleteSheet(sheets[i]);
  }
  var ui = SpreadsheetApp.getUi();
  var project_ids;
  var onActive = true;
  //if selected divide projects to sheets, we will separate them
  
  if(toSheets){
    
    var projects = getProjects(clientId);
    if( projects != null){
      for(var i = 0; i < projects.length; i++){ 
        var sheetName = clientName + " - " + projects[i].name ;
        project_ids = projects[i].id;
        
        var result = fillSheet(sheetName, useragent, workspace, clientId, start, end, project_ids, onActive);
        if(result){
          onActive = false;
        }
      }
    } else {
      var sheetName = clientName;
      project_ids = 0;
      var result = fillSheet(sheetName, useragent, workspace, clientId, start, end, project_ids, onActive);
      if(result){
        onActive = false;
      }
    }
  } else {
    var sheetName = clientName;
    project_ids = 0;
    var result = fillSheet(sheetName, useragent, workspace, clientId, start, end, project_ids, onActive);
    if(result){
      onActive = false;
    }
  }
  
  //it shouldn't be active... if is, no results found
  if(onActive){
    var sheet = ss.getActiveSheet();
    sheet.clear();
    sheet.setName(clientName);
  }
  
  //return data;
}

/**
* Display data in the sheet.
*/
function fillSheet(sheetName, useragent, workspace, clientId, start, end, project_ids, onActive){
  
  //reset page and position
  page = 1;
  total_pages = page;
  last_position = 0;
  
  //pull data and show them in rows in this sheet
  return getDataFromToggl(sheetName, useragent, workspace, clientId, start, end, project_ids, onActive);
  
  //ui.alert("Load complete");
  
}

/**
* Acquire data form toggl according parameters.
*/
function getDataFromToggl(sheetName, useragent, workspace, clientId, start, end, project_ids, onActive ){
  var ui = SpreadsheetApp.getUi();
  var ssa = SpreadsheetApp.getActiveSpreadsheet();
  var sheet, rate, format;
  //paginated load of data, pull first page, then see if there are others
  while(page <= total_pages){
    //for some reason, supplying project ids wont filter just projects for the client, but shows all of them
    //so either put projects, without client (bug maybe, but may display some other projects if client is 0)
    if(project_ids !==0){
      var data = callToggl("https://toggl.com/reports/api/v2/details", {
        "user_agent": useragent,
        "workspace_id": workspace, 
        "since": toYMD(start),
        "until": toYMD(end),
        "project_ids" : project_ids,
        "billable": "yes",
        "page": page
      });
    } else {
      //or set client id and show all of them
      var data = callToggl("https://toggl.com/reports/api/v2/details", {
        "user_agent": useragent,
        "workspace_id": workspace,
        "since": toYMD(start),
        "until": toYMD(end),
        "billable": "yes",
        "client_ids": clientId,
        "page": page
      });
    }
    if(page == 1){
      if(data == null || data['data'].length == 0){
        return false; 
      } else {
        if(onActive){
          //first is active
          sheet = ssa.getActiveSheet();
        } else {
          //others needs to be created
          sheet = ssa.insertSheet();
        }
        
        sheet.setName(sheetName);
        createSheetHeader(sheet); 
      }
    }
    
    var s, e, ss, es, total = data['data'].length;
    
    total_pages = Math.ceil(data.total_count / data.per_page);
    
    
    // all results to rows
    for (var i=0; i < total; i++) {
      ss = data['data'][i]['start'];
      ss = ss.substring(0, ss.length-6) + ".000Z";
      es = data['data'][i]['end']
      es = es.substring(0, es.length-6) + ".000Z";
      s = new Date(ss); //s.setHours(s.getHours()-2);
      e = new Date(es); //e.setHours(e.getHours()-2);
      rate = userProperties.getProperty(clientId + "-" + data['data'][i]['uid'] + "-price");
      format = userProperties.getProperty(clientId + "-" + data['data'][i]['uid'] + "-currency");
      sheet.getRange(last_position+i+2, 1).setValue(data['data'][i]['task']);
      sheet.getRange(last_position+i+2, 2).setValue(data['data'][i]['description']);
      sheet.getRange(last_position+i+2, 3).setValue(data['data'][i]['project']);
      sheet.getRange(last_position+i+2, 4).setValue(data['data'][i]['user']);
      sheet.getRange(last_position+i+2, 5).setValue(data['data'][i]['tags'].join(', '));
      sheet.getRange(last_position+i+2, 6).setValue(Utilities.formatDate(s,"GMT","yyyy-MM-dd")).setNumberFormat("yyyy-MM-dd");
      sheet.getRange(last_position+i+2, 7).setValue(Utilities.formatDate(e,"GMT","yyyy-MM-dd")).setNumberFormat("yyyy-MM-dd");
      sheet.getRange(last_position+i+2, 8).setValue(data['data'][i]['dur']/1000/60/60).setNumberFormat("0.000");
      sheet.getRange(last_position+i+2, 9).setValue(rate).setNumberFormat(format);
      sheet.getRange(last_position+i+2, 10).setValue("=H"+(last_position+i+2)+"*I"+(last_position+i+2)).setNumberFormat(format);
    }
    
    last_position = last_position+total;
    
    page++;
    
  }
  calculateSheetPrice(sheet, format);
  return true;
}

/**
* Get all projects of the client
*/
function getProjects(clientId){
  var data = callToggl("https://www.toggl.com/api/v8/clients/"+clientId+"/projects", {
                       });
  return data;
}

/**
* In order to get all users working for specific client it's needed to load every user in workspace
* and for each project get users, compare with the all users and select only those. Get project users
* doesn't contain name of the user, just ID's and there is no way to get user by ID yet.
* The query is cached.
*/
function getAllClientUsers(clientId){
  var cache, cacheName, cached, users = {}, version;
  version = scriptProperties.getProperty("VERSION");
  cache = CacheService.getScriptCache();
  cacheName = "toggl-clients-users" + getAuthDigest() + clientId + version; //if someone swith accounts, use that kind of cache name
  cached = cache.get(cacheName);
  
  if (cached != null && "" !== cached ) {
    var users = JSON.parse(cached);
    for (var k in users){
      if (users.hasOwnProperty(k)) {
        var userId = k;
        var userName = users[k].name;
        var price = userProperties.getProperty(clientId + "-" + userId + "-price");
        var currency = userProperties.getProperty(clientId + "-" + userId + "-currency");
        users[userId] = {name:userName, price:price, currency:currency};
      }
    }
    return JSON.stringify(users);
  }
  
  var projects = getProjects(clientId);
  
  var ui = SpreadsheetApp.getUi();
  var workspace = userProperties.getProperty("WORKSPACE_ID");
  var workspaceUsers = callToggl("https://www.toggl.com/api/v8/workspaces/"+workspace+"/users", {});
  if (typeof projects !== "undefined" && projects != null){
    for(var k = 0; k < projects.length; k++){
      var projectUsers = callToggl("https://www.toggl.com/api/v8/projects/"+projects[k].id+"/project_users", {});
      
      if (typeof projectUsers !== "undefined" && projectUsers != null){
        
        for(var i = 0; i < projectUsers.length; i++){
          for(var j = 0; j < workspaceUsers.length; j++){
            if(workspaceUsers[j].id === projectUsers[i].uid){
              var userId = workspaceUsers[j].id;
              var userName = workspaceUsers[j].fullname;
              var price = userProperties.getProperty(clientId + "-" + userId + "-price");
              var currency = userProperties.getProperty(clientId + "-" + userId + "-currency");
              users[userId] = {name:userName, price:price, currency:currency};
              break;
            }
          }  
        }   
      } else {
        for(var j = 0; j < workspaceUsers.length; j++){
          var userId = workspaceUsers[j].id;
          var userName = workspaceUsers[j].fullname;
          var price = userProperties.getProperty(clientId + "-" + userId + "-price");
          var currency = userProperties.getProperty(clientId + "-" + userId + "-currency");
          users[userId] = {name:userName, price:price, currency:currency};
        }
        break;
      }
    }
  }
  cache.put(cacheName, JSON.stringify(users), 60*60*24); // 24 hours cache
  return JSON.stringify(users);
}

/**
* Create header for the sheet
*/
function createSheetHeader(sheet){
  sheet.clear();
  //create header for the sheet
  //var ui = SpreadsheetApp.getUi();
  sheet.getRange(1, 1).setValue("Task").setFontWeight("bold");
  sheet.getRange(1, 2).setValue("Description").setFontWeight("bold");
  sheet.getRange(1, 3).setValue("Project").setFontWeight("bold");
  sheet.getRange(1, 4).setValue("User").setFontWeight("bold");
  sheet.getRange(1, 5).setValue("Tags").setFontWeight("bold");
  sheet.getRange(1, 6).setValue("Start").setFontWeight("bold");
  sheet.getRange(1, 7).setValue("End").setFontWeight("bold");
  sheet.getRange(1, 8).setValue("Duration").setFontWeight("bold");
  sheet.getRange(1, 9).setValue("Rate").setFontWeight("bold");
  sheet.getRange(1, 10).setValue("Amount").setFontWeight("bold");
  sheet.setFrozenRows(1); 
  
}

/**
* Calculate price of the table at the end.
*/
function calculateSheetPrice(sheet, format){
  //calculate price
  sheet.getRange(last_position+2, 8).setValue("=SUM(H1:H"+(last_position+1)+")").setFontWeight("bold").setNumberFormat("0.000");
  sheet.getRange(last_position+2, 10).setValue("=SUM(J1:J"+(last_position+1)+")").setFontWeight("bold").setNumberFormat(format);
  
  for (var j=1; j<11; j++) {
    sheet.autoResizeColumn(j);
  }  
}

/**
* When load data is clicked, user's prices and currences are saved in this function.
*/
function saveUsers(users){
  users = JSON.parse(users);
  var ui = SpreadsheetApp.getUi();
  for (var k in users){
    if (users.hasOwnProperty(k)) {
      userProperties.setProperty(k, users[k]);
    }
  }
}

/**
* Find users only within certain range. The latest users which are active.
*/
function getUsersWithinTheRange( clientId, timeRange, clientName ) {
  
  var start, end, day, now = new Date(), monday, useragent, workspace;
  now.setHours(0);
  now.setMinutes(0);
  now.setSeconds(0);
  
  if ( 'lastweek' == timeRange ) {
    day = now.getDay() || 7;
    if ( day !== 1 ) {
      now.setHours(-24 * (day - 1));
    }
    
    start = new Date(now.setDate(now.getDate()-7));
    end = new Date(now.setDate(now.getDate()+6));
  }
  
  if ( 'lastmonth' == timeRange ) {
    start = new Date(now.setDate(1));
    start.setMonth(start.getMonth()-1);
    end = new Date(now.setDate(0));
  }
  
  if ( 'thisweek' == timeRange ) {
    day = now.getDay() || 7;
    if ( day !== 1 ) {
      now.setHours(-24 * (day - 1));
    }
    start = new Date(now.setDate(now.getDate()));
    end = new Date(now.setDate(now.getDate()+6));
  }
  
  if ( 'thismonth' == timeRange ) {
    start = new Date(now.setDate(1));
    start.setMonth(start.getMonth());
    
    end = new Date(now);
    end.setMonth(end.getMonth()+1);
    end.setDate(0);
  }
  
  var cache, cacheName, cached, usersToShow = {}, version;
  version = scriptProperties.getProperty("VERSION");
  cache = CacheService.getScriptCache();
  cacheName = "toggl-clients-users" + getAuthDigest() + clientId + toYMD(start) + toYMD(end) + version; //if someone swith accounts, use that kind of cache name
  cached = cache.get(cacheName);
  
  if (cached != null && "" !== cached ) {
    var usersToShow = JSON.parse(cached);
    for (var k in usersToShow){
      if (usersToShow.hasOwnProperty(k)) {
        var userId = k;
        var userName = usersToShow[k].name;
        var price = userProperties.getProperty(clientId + "-" + userId + "-price");
        var currency = userProperties.getProperty(clientId + "-" + userId + "-currency");
        usersToShow[userId] = {name:userName, price:price, currency:currency};
      }
    }
    return JSON.stringify(usersToShow);
  }
  
  useragent = userProperties.getProperty("USER_AGENT");
  workspace = userProperties.getProperty("WORKSPACE_ID");
  
  page = 1;
  total_pages = page;
  
  var allClientUsers = getAllClientUsers(clientId);
  var usersInRange = {};
  
  while(page <= total_pages){
    var data = callToggl("https://toggl.com/reports/api/v2/details", {
      "user_agent": useragent,
      "workspace_id": workspace,
      "since": toYMD(start),
      "until": toYMD(end),
      "billable": "yes",
      "client_ids": clientId,
      "page": page
    });
    
    if(data == null || data['data'].length == 0){
      page++;
      continue; 
    }
    var total = data['data'].length;
    
    total_pages = Math.ceil(data.total_count / data.per_page);
    
    for (var i=0; i < total; i++) {
      usersInRange[data['data'][i]['user']] = data['data'][i]['user'];     
    }
    
    page++;
    
  }
  
  var usersParsed = JSON.parse(allClientUsers);
  
  for (var key in usersParsed) {
    if (usersParsed.hasOwnProperty(key)) {
      var obj = usersParsed[key];
      for (var prop in obj) {
        // important check that this is objects own property 
        // not from prototype prop inherited
        if(obj.hasOwnProperty(prop)){     
          if(prop == "name"){ 
            for (var user in usersInRange) {
              if(user == obj[prop]){
                usersToShow[key] = usersParsed[key];
              }
            }
          }
        }
      }
    }
  }
  
  
  cache.put(cacheName, JSON.stringify(usersToShow), 60*60*24); // 24 hours cache
  return JSON.stringify(usersToShow);
  
}
