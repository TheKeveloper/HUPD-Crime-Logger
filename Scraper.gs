function getSpreadsheet(){
  return SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1AMbEglG18BDz4-mQgTfAl4-jiT2Th_tKyIjwBEMDWF8/edit#gid=0");
}

//Modified from: https://ctrlq.org/code/20566-extract-text-pdf
function getPDFText(url) {  
  var blob = UrlFetchApp.fetch(url).getBlob();
  var resource = {
    title: blob.getName(),
    mimeType: blob.getContentType()
  };
  
  // Enable the Advanced Drive API Service
  var file = Drive.Files.insert(resource, blob, {ocr: true, ocrLanguage: "en"});
  
  // Extract Text from PDF file
  var doc = DocumentApp.openById(file.id);
  var text = doc.getBody().getText();
  Drive.Files.remove(file.id);
  return text; 
}

function Incident(){
  this.time = null;
  this.type = null;
  this.status = null;
  this.location = null;
  this.description = null;
  this.area = null;
  this.lat = null;
  this.long = null;
}

function getInfo(str){
  // Gets time and place address info
  var infoRegex = /\D+[0-9\-]+\D+(OPEN|CLOSED|ARREST) \d+:\d+ (AM|PM)/gm;
  var infos = str.match(infoRegex);
  var incidents = [];
  infos.forEach(function(info){
    var incident = new Incident();
    var lines = info.split(/\r\n|\r|\n/);
    var statusRegex = /(OPEN|CLOSED|ARREST)/;
    var timeRegex = /\d+:\d+ \D+/;
    // Accounting for weird formatting of the PDF
    if(lines.length == 2){
      incident.location = lines[0];
      incident.status = lines[1].match(statusRegex)[0].trim();
      incident.type = lines[1].substring(lines[1].indexOf(statusRegex)).trim();
      incident.time = lines[1].match(timeRegex)[0].trim();
    }
    else if (lines.length == 3){
      incident.type = lines[0].trim();
      incident.location = lines[1].trim();
      incident.status = lines[2].match(statusRegex)[0].trim();
      incident.time = lines[2].match(timeRegex)[0].trim();
    }
    else if(lines.length == 4){
      incident.type = lines[1].trim();
      incident.location = lines[2].trim();
      incident.status = lines[3].match(statusRegex)[0].trim();
      incident.time = lines[3].match(timeRegex)[0].trim();
    }
    else{
      incident.type = "ERROR";
      incident.location = "ERROR";
      incident.status = "ERROR";
      incident.time = "ERROR";
    }
    // Cleaning up in case of weird regex problems
    if(incident.location && incident.location[0] == "/"){
      incident.location = incident.location.substring(4).trim(); 
    }
    if(incident.type && incident.type[0] == "/"){
      incident.type = incident.type.substring(4).trim();
    }
    incident.type = incident.type.replace(/(OPEN|CLOSED|ARREST) \d+:\d+ (AM|PM)/gm, "");
    incidents.push(incident);
  });
  return incidents;
}

// Get the description strings for all incidents
function getDescriptions(str){
  return str.match(/(Officer|Officers) .+/g);
}

// Return areas as list
function getAreas(str){
  return str.match(/(AM |PM |\n)(ALLSTON|CAMBRIDGE|BOSTON)/g);
}

// Associate a list of incidents with the geolocations
function getGeo(incidents){
  try {
    incidents.forEach(function(incident){
      var response = Maps.newGeocoder().setBounds(42.310021, -71.122081, 42.382741, -70.993858).geocode(incident.location + " " + incident.area);
      if (response.results.length > 0){
        incident.lat = response.results[0].geometry.location.lat;
        incident.long = response.results[0].geometry.location.lng
      }
    });
  }
  catch (e){
    console.log(e);
  }
}

// Convert an incident object to an array for adding to spreadsheet
function IncidentToArr(incident, date){
  return [Utilities.formatDate(date, "EST", "MM-dd-yyyy"), date.getDay(), incident.time, incident.type.trim(), incident.status.trim(), 
            incident.location.trim(), incident.area.replace(/(AM | PM)/, "").trim(), incident.description.trim(), incident.lat, incident.long];
}

function scrape(date){
  // Get correctly formatted date for url
  var year = (date.getYear() - 2000).toString();
  var month = (date.getMonth() + 1).toString();
  if(month.length < 2){
    month = "0" + month;
  }
  var day = date.getDate().toString();
  if(day.length < 2){
    day = "0" + day;
  }
  var url = "https://www.hupd.harvard.edu/files/hupd/files/" + month + day + year + ".pdf";
  var timeRegex = /\d+:\d+ \D+ \d+:\d+ \D+ - \d+:\d+ \D{2}/;
  var pdfText = getPDFText(url);
  var incidents = getInfo(pdfText);
  var descriptions = getDescriptions(pdfText);
  var areas = getAreas(pdfText);
  //Do final cleanup of the text
  if(incidents.length == areas.length){
    for(var i = 0; i < incidents.length; i++){
      incidents[i].area = areas[i].match(/(AM |PM |\n)(ALLSTON|CAMBRIDGE|BOSTON)/)[0];
    }
  }
  if(incidents.length == descriptions.length){
    for(var i = 0; i < descriptions.length; i++){
      incidents[i].description = descriptions[i];
    }
  }
  getGeo(incidents);
  // Add to spreadsheet
  var spreadsheet = getSpreadsheet(); 
  var sheet = spreadsheet.getSheets()[0];
  var newRows = incidents.map(function(elt){ return IncidentToArr (elt, date);});
  if(newRows.length > 0){
    sheet.insertRowsAfter(1, newRows.length);
    sheet.getRange(2, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
  var logString = "Inserted " + newRows.length + " new rows on " + date.toDateString();
  Logger.log(logString);
  console.log(logString);
}

// Geocode the existing incidents in the spreadsheet
function geocodeExisting(){
  var spreadsheet = getSpreadsheet();
  var sheet = spreadsheet.getSheets()[0];
  var values = sheet.getRange("F:G").getValues();
  var coords = [];
  const start = 6; 
  const end = 43;
  for(var i = start; i <= end; i++){
    var response = Maps.newGeocoder().setBounds(42.310021, -71.122081, 42.382741, -70.993858).geocode(values[i][0] + " " + values[i][1]);
    if(response.results.length > 0){
      coords.push([response.results[0].geometry.location.lat, response.results[0].geometry.location.lng]);
    }
    else{
      coords.push([null, null]);
    }
  }
  
  sheet.getRange(start + 1, 9, coords.length, 2).setValues(coords);
}

// Scrapes from last report date to the current date
// Runs once every day
function main(){
  var spreadsheet = getSpreadsheet();
  var values = spreadsheet.getSheets()[0].getRange("A:A").getValues();
  var lastDate = new Date(values[1][0]);
  lastDate.setDate(lastDate.getDate() + 1); 
  var curDate = new Date();
  curDate.setDate(curDate.getDate() - 1);
  while(lastDate <= curDate){
    try{
      scrape(lastDate);
    }
    catch(e){
      console.log(e);
    }
    lastDate.setDate(lastDate.getDate() + 1);
  }
}

// Properly formats the spreadsheet by trimming whitespace
function cleanup(){
  var spreadsheet = getSpreadsheet();
  var sheet = spreadsheet.getSheets()[0];
  var range = sheet.getRange("D:H");
  var values = range.getValues();
  values = values.map(function(arr){
    arr[0] = arr[0].replace(/(OPEN|CLOSED|ARREST) \d+:\d+ (AM|PM)/gm, "").replace("PM ", "");
    return arr.map(function(str){
      return str.trim();
    });
  });
  range.setValues(values);
}

// Test function for miscellaneous work
function test(){
  for(var i = 25; i <= 31; i++){
    scrape(new Date(2018, 04, i));
  }
}