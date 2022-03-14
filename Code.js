// assuming yesterday's data sent to Gmail, queries Gmail for latest email send, pulls the data,
// and imports the data into the specified sheet
function importData() {
  var spreadsheetId = identifiers['destinationSpreadsheetId'];

  var threads = GmailApp.search(identifiers['gmailSearchQuery']);

  // assumes there is only one suitable thread returned above
  var message = threads[0].getMessages()[threads[0].getMessages().length - 1];

  // "split" invocation below creates a 2-item array
  var dataURL = message.getPlainBody().split("below:")[1].trim();
  var csvString = UrlFetchApp.fetch(dataURL).getBlob().getDataAsString();

  // pull everything but the column headers
  var data = Utilities.parseCsv(csvString, ",").slice(1);

  // -- Use if message has attachment -- 
  // getting the first (and only) attachment
  // var attachment = message.getAttachments()[0]; 
  // var attachmentAsCsv = Utilities.unzip(attachment)[0];

  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(identifiers['destinationSheetName']);

  // adds data to sheet, adding data below existing data
  // sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data); 

  // adds data to sheet, replacing data over existing data
  // offset row to 2 as needed if you want to keep the column headers in the sheet
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
}

// auto-drags down formulas on right side when data to the left is updated
function fillFormulas() {
  var spreadsheetId = identifiers['destinationSpreadsheetId'];
  var targetSheetName = identifiers['destinationSheetName'];
  var firstColumnOfFormulaColumns = "H";
  var lastColumnOfFormulaColumns = "H"; 
  var sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(targetSheetName);

  var positionOfLastTextInFirstFormulaColumn = lastRowForSingleColumn( sheet, columnIndex(firstColumnOfFormulaColumns) );

  // activate the last row of formulas, which will be gotten and set later
  var activeRangeReference = `${targetSheetName}!${firstColumnOfFormulaColumns}${positionOfLastTextInFirstFormulaColumn.toString()}:${lastColumnOfFormulaColumns}${positionOfLastTextInFirstFormulaColumn.toString()}`;

  sheet.getRange(activeRangeReference).activate();

  // get formulas from activated range; use R1C1 to maintain relative references
  var sourceFormulas = sheet.getRange(activeRangeReference).getFormulasR1C1();

  // use column A (1st column) to calc how many rows to paste the formulas into
  var destinationRangeRowCount = lastRowForSingleColumn(sheet, 1) - positionOfLastTextInFirstFormulaColumn;

  // define destination range -- the range in which formulas above will be assigned
  // offset +1 since destination range length == # rows to fill + 1, i.e. the single active row range
  var destinationRange = sheet.getActiveRange().offset(0, 0, destinationRangeRowCount + 1);

  // assign an empty array to build the range of formulas to set into the destination range;
  // dimensions of source formula array must match those of destination range
  var completeSourceFormulas = [];

  // pushing the same row of formulas as many times as there are rows in the destination range
  for(i = 0; i <= destinationRangeRowCount; i++) {
    completeSourceFormulas.push(sourceFormulas[0]);
  }

  destinationRange.setFormulasR1C1(completeSourceFormulas);
};

function lastRowForSingleColumn(sheet, column) {
  // Get the last row with data for the whole sheet.
  var numRows = sheet.getLastRow();
  
  // Get all data for the given column
  var data = sheet.getRange(1, column, numRows).getValues();
  
  // Iterate backwards and find first non empty cell
  for(var i = data.length - 1 ; i >= 0 ; i--){
    if (data[i][0] != null && data[i][0] != ""){
      return i + 1;
    }
  }
}

function columnIndex(columnChar) {
  if( columnChar.length === 1 ) {
    return columnChar.charCodeAt(0) - 64;
  } else {
    return ((columnChar.charCodeAt(0) - 64) * 26) + columnIndex(columnChar.slice(1));
  }
}

function sendData() {
  var file = createCsvInGDrive();
  sendEmailWithAttachment(file);
}

function createCsvInGDrive() {
  var sheetID = identifiers['destinationSpreadsheetId'];

  // get data in spreadsheet
  var ss = SpreadsheetApp.openById(sheetID);
  var sheet = ss.getSheetByName(identifiers['destinationSheetName']);
  var range = sheet.getRange("A1").getDataRegion();
  var data = range.getValues();

  var csvString = buildCsvString(data);

  // create file, and save file to specified folder
  var folderId = identifiers['driveFolderId'];
  var folder = DriveApp.getFolderById(folderId);

  var today = new Date();
  var m = today.getMonth() + 1; // month is zero-indexed
  var d = today.getDate();
  var y = today.getFullYear();

  var fileName = `Data ${m}-${d}-${y}.csv`
  var newFile = folder.createFile(fileName, csvString);
  return newFile;
}

function buildCsvString(data) {
  // work with data separately from column headers, as headers are structured differently from data
  var headers = data[0];

  var newData = data.slice(1).map ( row => {
    // wrap var's in brackets with multiple assignment
    var [month, date, year] = row[3].split('/');

    // remove leading zeros from month and day, if present
    var newMonth = month[0].startsWith('0') ? month.slice(-1) : month;
    var newDate = date[0].startsWith('0') ? date.slice(-1) : date;
    
    return row.slice(0, 3).concat([`${newMonth}/${newDate}/${year}`]).concat(row.slice(4));
  });

  var newDataWithHeaders = [ headers, ...newData ];
  return newDataWithHeaders.map( r => r.join(',') ).join('\n');
}

function sendEmailWithAttachment(file) {
  var recipients = identifiers['gmail']['recipients'];
  var subject = identifiers['gmail']['subject'];
  var body = identifiers['gmail']['body'];
  var attachments = [DriveApp.getFileById(file.getId())];

  GmailApp.sendEmail(recipients, subject, body, {
    attachments: attachments
  })
}


