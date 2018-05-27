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
}

function getInfo(str){
  //Gets time and place address info
  var infoRegex = /\D+\d+\D+(OPEN|CLOSED|ARREST) \d+:\d+ (AM|PM)/gm;
  var infos = str.match(infoRegex);
  var incidents = [];
  infos.forEach(function(info){
    var incident = new Incident();
    var lines = info.split(/\r\n|\r|\n/);
    var statusRegex = /(OPEN|CLOSED|ARREST)/;
    var timeRegex = /\d+:\d+ \D+/;
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
    else{
      incident.type = "ERROR";
      incident.location = "ERROR";
      incident.status = "ERROR";
      incident.time = "ERROR";
    }
    if(incident.location && incident.location[0] == "/"){
      incident.location = incident.location.substring(4).trim(); 
    }
    if(incident.type && incident.type[0] == "/"){
      incident.type = incident.type.substring(4).trim();
    }
    incidents.push(incident);
  });
  return incidents;
}

function getDescriptions(str){
  return str.match(/(Officer|Officers) .+/g);
}

function getAreas(str){
  return str.match(/(AM |PM |\n)(ALLSTON|CAMBRIDGE|BOSTON)/g);
}

function IncidentToArr(incident, date){
  return [Utilities.formatDate(date, "EST", "MM-dd-yyyy"), date.getDay(), incident.time, incident.type, incident.status, 
            incident.location, incident.area, incident.description];
}

function scrape(date){
  var year = (date.getYear() - 2000).toString();
  var month = (date.getMonth()).toString();
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
  Logger.log(incidents.length);
  Logger.log(areas.length);
  Logger.log(descriptions.length);
  if(incidents.length == areas.length){
    for(var i = 0; i < incidents.length; i++){
      incidents[i].area = areas[i];
    }
  }
  if(incidents.length == descriptions.length){
    for(var i = 0; i < descriptions.length; i++){
      incidents[i].description = descriptions[i];
    }
  }
  
  var spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1AMbEglG18BDz4-mQgTfAl4-jiT2Th_tKyIjwBEMDWF8/edit#gid=0");
  var sheet = spreadsheet.getSheets()[0];
  var newRows = incidents.map(function(elt){ return IncidentToArr (elt, date);});
  if(newRows.length > 0){
    sheet.insertRowsAfter(1, newRows.length);
    sheet.getRange(2, 1, newRows.length, 8).setValues(newRows);
  }
  var logString = "Inserted " + newRows.length + " new rows on " + date.toDateString();
  Logger.log(logString);
  console.log(logString);
}

function main(){
  scrape(new Date(2018, 05, 08));
}