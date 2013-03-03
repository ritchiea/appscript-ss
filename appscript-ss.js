var ss = SpreadsheetApp.openById('SPREADSHEET_ID'); // .getSheetByName('v3status')

var SHEET_START = 6;
var SHEET_END = (getLastRow());
var PAGE_COL = 'A';
var STATUS_COL = 'B';
var NOTES_COL = 'C';
var PAGE_COL_NUM = 0;
var STATUS_COL_NUM = 1;
var NOTES_COL_NUM = 2;
var RANGE_START = 'A6';
var RANGE_END = 'C'+(getLastRow());
var STATUSES = ['Final', 'Draft', 'Need'];

var range = ss.getSheets()[0].getRange(RANGE_START+':'+RANGE_END);
ss.setNamedRange('Statuses', range);

STATUSES.forEach(function(status) {
  this['get'+status+'Percent'] = function() {
    var statuses = getAllStatuses();
    var instances = 0;
  
    statuses.forEach(function(element) {
      if (element === status) {
        instances++;
      }
      });
       
      percent = ((instances / statuses.length)*100).toFixed();

      return percent;    
  }
});

function getLastRow() {
  var rows = ss.getSheets()[0].getRange('A1:A100');
  var values = rows.getValues();
  
  for (var i = 99; i >= 0; i--) {
    
    if (values[i] != ""){
      var lastrow = i+1;
      
      return (lastrow);
    }
    
  }
}

function getAllPages() {
  var values = getT3Values();
  var pages = [];
  
  for (var i = 0; i <= (SHEET_END-SHEET_START); i++) {
    pages.push(values[i][0]);
  }
   
   // Browser.msgBox(pages);
   return pages;
  
}

function getAllStatuses() {
  var values = getT3Values();
  var statuses = [];
  
  for (var i = 0; i <= (SHEET_END-SHEET_START); i++) {
    statuses.push(values[i][1]);
  }
   
   // Browser.msgBox(statuses);
   return statuses;
  
}

function getStatus(page) {
  var values = getT3Values();
  var status;
  
  for (var i = 0; i < 3; i++) {
    if (values[i][0] === page) {
      status = values[i][1];
      return status;
    }
  }
  
  if (status === undefined) {
    return false;
  }
  
}

function getT3Values() {
  var values = range.getValues();
  return values;
}

function doGet(request) {
  var finalPercent = getFinalPercent();
  var draftPercent = getDraftPercent();
  var needPercent = getNeedPercent();
  var result = {
    Final: finalPercent,
    Draft: draftPercent,
    Need: needPercent
  };
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
