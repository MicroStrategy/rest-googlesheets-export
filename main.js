// Google Script web apps run a function named "doGet" by default
function doGet(e) {

  config["searchVisName"] = e.parameter.searchvisname
  config["dossierUrl"] = e.parameter.dossierurl
  
  if (e.parameter.spreadsheetId != undefined)
    config["spreadsheetId"] = e.parameter.spreadsheetId
  
  var regex = /(.*)\/app\/(\w*)\/(\w*)/
  config["webserver"] = config["dossierUrl"].match(regex)[1]
  config["projectID"] = config["dossierUrl"].match(regex)[2]
  config["dossierID"] = config["dossierUrl"].match(regex)[3]
  
  // Connect to MicroStrategy and store AuthToken + Cookie in config structure
  config = mstrConnect(config);

  // Get Dossier structure and store the name in config structure
  var dossierDef = mstrGetDossierDefinition(config);
  config["dossierName"] = dossierDef.name;

  // Look for the searched visualization location. Need to add error management
  var visLocation = mstrGetVisLocationFromName(config, dossierDef);

  var dossierInstance = mstrGetDossierInstance(config);
  config["dossierInstanceId"] = dossierInstance.mid;

  // Extract target visualization Metadata and Data
  var visDetailedData = mstrGetVisData(config, visLocation);

  // Restructure visualization Metadata and Data in a more userfriendly formaty
  var dataset = extractDataSet(visDetailedData)  
  
  if(config["spreadsheetId"] != "")
    updateSpreadsheet(config, dataset)
  else
    config = createSpreadsheet(config, dataset)
  
  mstrLogout(config);
  
  return HtmlService.createHtmlOutput(
    "<form action='https://docs.google.com/spreadsheets/d/" + config["spreadsheetId"] + "' method='get' id='redirect'></form>" + 
    "<script>document.getElementById('redirect').submit();</script>")
}

function mstrConnect(config) {
  var jsonData = {
    'username': config["username"],
    'password': config["password"],
    'loginMode': 1,
    'maxSearch': 3,
    'workingSet': 0,
    'changePassword': false,
    'newPassword': 'string',
    'metadataLocale': 'en_us',
    'warehouseDataLocale': 'en_us',
    'displayLocale': 'en_us',
    'messagesLocale': 'en_us',
    'numberLocale': 'en_us',
    'timeZone': 'UTC',
    'applicationType': 35
  };

  var options = {
    'method' : 'post',
    'contentType': 'application/json',
    'accept': 'application/json',
    'payload' : JSON.stringify(jsonData)
  };
  
  var httpResponse = UrlFetchApp.fetch(config["webserver"]+"/api/auth/login", options);

  var headers = httpResponse.getHeaders();
  config["x-mstr-authtoken"] = headers["x-mstr-authtoken"]
  config["Set-Cookie"] = headers["Set-Cookie"].substring(0, 43)
  
  return config
}

function mstrLogout(config) {
  var options = {
    'method' : 'POST',
    'contentType': 'application/json',
    'accept': 'application/json',
    'headers' : {'x-mstr-authtoken' : config["x-mstr-authtoken"]}
  };

  UrlFetchApp.fetch(config["webserver"]+"/api/auth/logout", options);
}

function mstrGetDossierDefinition(config){
  var options = {
    'method' : 'GET',
    'contentType': 'application/json',
    'accept': 'application/json',
    'headers' : {'x-mstr-authtoken' : config["x-mstr-authtoken"], 'x-mstr-projectid' : config["projectID"],'cookie' : config["Set-Cookie"]}
  };

  var httpResponse = UrlFetchApp.fetch(config["webserver"]+"/api/dossiers/"+config["dossierID"]+"/definition", options);
  return JSON.parse(httpResponse)
}

function mstrGetVisLocationFromName(config, definition) {
  var name = config["searchVisName"]
  var visLocation = {
    'chapterKey' : undefined,
    'visKey' : undefined
  }
  var chapters = definition.chapters;
  for (var i = 0; i < chapters.length; i++){
    var chapter = chapters[i];
    var pages = chapter.pages;        
    for (var j = 0; j < pages.length; j++){
      var page = pages[j];
      var visualizations = page.visualizations;   
      for (var k = 0; k < visualizations.length; k++){
        var vis = visualizations[k];
        if(vis.name.toLowerCase() == name.toLowerCase()){
          Logger.log("FOUND MATCH");
          visLocation.chapterKey = chapter.key
          visLocation.visKey = vis.key
          return visLocation
        }
      }
    }
  }
  Logger.log("Could not locate " + name + "in your dossier")
  return visLocation
}

function mstrGetDossierInstance(config) {
  var options = {
    'method' : 'POST',
    'contentType': 'application/json',
    'accept': 'application/json',
    'headers' : {'x-mstr-authtoken' : config["x-mstr-authtoken"], 'x-mstr-projectid' : config["projectID"],'cookie' : config["Set-Cookie"]}
  };

  var httpResponse = UrlFetchApp.fetch(config["webserver"]+"/api/dossiers/"+config["dossierID"]+"/instances", options);

  return JSON.parse(httpResponse)
}

function mstrGetVisData(config, visLocation) {
  var options = {
    'method' : 'GET',
    'contentType': 'application/json',
    'accept': 'application/json',
    'headers' : {'x-mstr-authtoken' : config["x-mstr-authtoken"], 'x-mstr-projectid' : config["projectID"],'cookie' : config["Set-Cookie"]}
  };

  var httpResponse = UrlFetchApp.fetch(config["webserver"]+"/api/dossiers/"+config["dossierID"]+
                                       "/instances/"+config["dossierInstanceId"]+
                                       "/chapters/"+visLocation["chapterKey"]+
                                       "/visualizations/"+visLocation["visKey"], options);
  return JSON.parse(httpResponse)
}

function extractDataSet(visDetailedData) {
  var visAttributes = visDetailedData["result"]["definition"]["attributes"]
  var visMetrics = visDetailedData["result"]["definition"]["metrics"]
  var visData = visDetailedData["result"]["data"]["root"]["children"]
  var data = Array(visData.length)
  var dataLength = visData.length+2
  
  var attributes = Array(visAttributes.length)  
  for(var i = 0; i < visAttributes.length; i++) {
    attributes[i] = visAttributes[i]["name"]
  }
  var metrics = Array(visMetrics.length)  
  for(var i = 0; i < visMetrics.length; i++) {
    metrics[i] = visMetrics[i]["name"]
  }
  for(var i = 0; i < data.length; i++) {
    data[i] = Array(2)
    data[i][0] = visData[i]["element"]["name"]
    data[i][1] = visData[i]["metrics"][metrics[0]]["fv"]
  }

  Logger.log(attributes)
  Logger.log(metrics)
  Logger.log(data)
  var dataSet = {"attributes": attributes, "metrics": metrics, "data": data}
  return dataSet
}

function updateSpreadsheet(config, dataset) {
  var spreadsheetId = config["spreadsheetId"];
  var request = {
    majorDimension: "ROWS",
    values: [[dataset["attributes"][0],dataset["metrics"][0]]]
  };
  Sheets.Spreadsheets.Values.update(
    request,
    spreadsheetId,
    "Sheet1!A1:B1",
    {valueInputOption: "USER_ENTERED"}
  );
  
  var spreadsheetId = config["spreadsheetId"];
  var request = {
    majorDimension: "ROWS",
    values: dataset["data"]
  };
  Sheets.Spreadsheets.Values.update(
    request,
    spreadsheetId,
    "Sheet1!A2:B"+dataset["data"].length+1,
    {valueInputOption: "USER_ENTERED"}
  );
}

function createSpreadsheet(config, dataset) {
  currentDateTime = Utilities.formatDate(new Date(), "CET", "yyyy-MM-dd HH:mm:ss ")  
  var newSpreadSheet = SpreadsheetApp.create(config["dossierName"] + " - " + config["searchVisName"] + " - " + currentDateTime)
  var spreadsheetId = newSpreadSheet.getId()
  var request = {
    majorDimension: "ROWS",
    values: [[dataset["attributes"][0],dataset["metrics"][0]]]
  };
  Sheets.Spreadsheets.Values.update(
    request,
    spreadsheetId,
    "Sheet1!A1:B1",
    {valueInputOption: "USER_ENTERED"}
  );
  
  var request = {
    majorDimension: "ROWS",
    values: dataset["data"]
  };
  Sheets.Spreadsheets.Values.update(
    request,
    spreadsheetId,
    "Sheet1!A2:B"+dataset["data"].length+1,
    {valueInputOption: "USER_ENTERED"}
  );
  
  config["spreadsheetId"] = spreadsheetId;
  return config;
}
