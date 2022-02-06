// @OnlyCurrentDoc
function onOpen(e) {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Import CSV data ⬇️")
    .addItem("Import from URL", "importCSVFromUrl")
    .addItem("Import from Drive", "importCSVFromDrive")
    .addToUi();
}

//Displays an alert as a Toast message
function displayToastAlert(message) {
  SpreadsheetApp.getActive().toast(message, "⚠️ Alert"); 
}

//Imports a CSV file at a URL into the Google Sheet
function importCSVFromUrl() {
  // let url = promptUserForInput("Please enter the URL of the CSV file:");
  let url = 'https://raw.githubusercontent.com/anhphong22/NLP/main/pol.csv'
  let new_column = ['question', 'answer', 'status']

  let headers = Utilities.parseCsv(UrlFetchApp.fetch(url))[0];
  headers = headers.concat(new_column)

  let contents = UrlFetchApp.fetch(url).getContentText();
  contents = CSVToArray(contents, ',');
  
  contents[0] = headers;
  for(let row in contents){
    if(row > 0){
      contents[row] = contents[row].concat([ ...Array(new_column.length).keys() ].map( i => i = ''))
    }
  }

  let sheetName = writeDataToSheet(contents, new_column.length);
  displayToastAlert("The CSV file was successfully imported into " + sheetName + ".");
  onEdit()
}

function CSVToArray( strData, strDelimiter ){
    // Check to see if the delimiter is defined. If not,
    // then default to comma.
    strDelimiter = (strDelimiter || ";");

    // Create a regular expression to parse the CSV values.
    let objPattern = new RegExp(
        (
            // Delimiters.
            "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +

            // Quoted fields.
            "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +

            // Standard fields.
            "([^\"\\" + strDelimiter + "\\r\\n]*))"
        ),
        "gi"
        );


    // Create an array to hold our data. Give the array
    // a default empty first row.
    let arrData = [[]];

    // Create an array to hold our individual pattern
    // matching groups.
    let arrMatches = null;


    // Keep looping over the regular expression matches
    // until we can no longer find a match.
    while (arrMatches = objPattern.exec( strData )){

        // Get the delimiter that was found.
        let strMatchedDelimiter = arrMatches[ 1 ];

        // Check to see if the given delimiter has a length
        // (is not the start of string) and if it matches
        // field delimiter. If id does not, then we know
        // that this delimiter is a row delimiter.
        if (
            strMatchedDelimiter.length &&
            strMatchedDelimiter !== strDelimiter
            ){

            // Since we have reached a new row of data,
            // add an empty row to our data array.
            arrData.push( [] );

        }

        let strMatchedValue;

        // Now that we have our delimiter out of the way,
        // let's check to see which kind of value we
        // captured (quoted or unquoted).
        if (arrMatches[ 2 ]){

            // We found a quoted value. When we capture
            // this value, unescape any double quotes.
            strMatchedValue = arrMatches[ 2 ].replace(
                new RegExp( "\"\"", "g" ),
                "\""
                );

        } else {

            // We found a non-quoted value.
            strMatchedValue = arrMatches[ 3 ];

        }


        // Now that we have our value string, let's add
        // it to the data array.
        arrData[ arrData.length - 1 ].push( strMatchedValue );
    }

    // Return the parsed data.
    return( arrData );
}

//Imports a CSV file in Google Drive into the Google Sheet
function importCSVFromDrive() {
  let fileName = promptUserForInput("Please enter the name of the CSV file to import from Google Drive:");
  let files = findFilesInDrive(fileName);
  if(files.length === 0) {
    displayToastAlert("No files with name \"" + fileName + "\" were found in Google Drive.");
    return;
  } else if(files.length > 1) {
    displayToastAlert("Multiple files with name " + fileName +" were found. This program does not support picking the right file yet.");
    return;
  }
  let file = files[0];
  let contents = Utilities.parseCsv(file.getBlob().getDataAsString());
  let sheetName = writeDataToSheet(contents);
  Logger.log(contents)
  displayToastAlert("The CSV file was successfully imported into " + sheetName + ".");
}

//Prompts the user for input and returns their response
function promptUserForInput(promptText) {
  let ui = SpreadsheetApp.getUi();
  let prompt = ui.prompt(promptText);
  let response = prompt.getResponseText();
  return response;
}

//Returns files in Google Drive that have a certain name.
function findFilesInDrive(filename) {
  let files = DriveApp.getFilesByName(filename);
  let result = [];
  while(files.hasNext())
    result.push(files.next());
  return result;
}

// Return version from a sheet
function getVersion(sheet){
  return sheet.getName().replace( /^\D+/g, '')
}

// Rotating a matrix - W.T
function rotatingArray(W) {
    let result_W = []
    for (let j = 0; j < W[0].length; j++) {
      let x_W = []
      for (let i in W) {
        x_W.push(W[i][j])
      }
      if (x_W.length > 0) {
        result_W.push(x_W)
      }
    }
    if (result_W.length > 0) {
      return result_W
    }
}

// Return history change sheet (combine two last sheets)
function getHistory(ss, data, numNewColumn, nameSheet){
  let last_sheets = ss.getSheets()[0]
  let last_data = last_sheets.getRange(15,1,data.length,data[0].length).getValues()
  console.log(data.length, last_data.length, last_data[0].length)
  let status_combine = ""
  let deleted_lines = []
  let added_lines = []
  let updated_lines = []

  let last_data_T = rotatingArray(last_data)
  console.log(last_data_T.length, last_data_T[0].length)
}

// Generate UI for each inserted sheet 
function genUI(ss, data, nameSheet, indexSheet, numNewColumn){
  sheet = ss.insertSheet(nameSheet, indexSheet)

  sheet.getRange(15, 1, 1, data[0].length-numNewColumn).setBackground("#f1f1f1")
  sheet.getRange(15, 1, 1, data[0].length).setFontWeight("bold")
  sheet.getRange(15, 1, 1, data[0].length).setHorizontalAlignment("center")
  sheet.getRange(15, 1, 1, data[0].length).setBorder(true, true, true, true, true, true, '#b8f3f9', SpreadsheetApp.BorderStyle.SOLID)

  sheet.getRange(15, data[0].length-numNewColumn, 1, numNewColumn).setBackground("#46bcd6")
  sheet.getRange(15, data[0].length, 1, 1).setBackground("orange")

  sheet.getRange(15, 1, data.length, data[0].length).setValues(data);
  sheet.getRange(16, 1, data.length-1, data[0].length).setHorizontalAlignment("left")
  sheet.getRange(16, 1, 1, data[0].length).setBorder(null, true, true, true, true, true, '#b8f3f9', SpreadsheetApp.BorderStyle.DASHED)
  sheet.getRange(17, 1, data.length-1, data[0].length).setBorder(true, true, true, true, true, true, '#b8f3f9', SpreadsheetApp.BorderStyle.DASHED)

  sheet.setRowHeightsForced(15, data.length, 24);
  sheet.autoResizeColumn(1);

  sheet.hideColumns(sheet.getLastColumn()+1, sheet.getMaxColumns()-sheet.getLastColumn())
  sheet.hideRows(sheet.getLastRow()+1, sheet.getMaxRows()-sheet.getLastRow())

  let descriptions = getHistory(ss, data, numNewColumn, nameSheet)
  let headers_revision = ["Revision History", "Objectives"]
  let headers_rhistory = ["Version","Author","Issued","Descriptions"]
  let table_report = []
  let userEmail = Session.getActiveUser().getEmail()//PropertiesService.getUserProperties().getProperty("userEmail")
  table_report.push(headers_rhistory)
  table_report.push([nameSheet, userEmail, new Date().toLocaleDateString("vi-VN"), descriptions])
  sheet.getRange(4,1, table_report.length, table_report[0].length).setValues(table_report)
  
  // revision history
  sheet.getRange(3,1).setValue(headers_revision[0]).setFontSize(14).setFontWeight("bold").setBackground('#a2d2ff').setHorizontalAlignment('center').setBorder(true, true, true, true, true, true, '#b8f3f9', SpreadsheetApp.BorderStyle.SOLID)
  sheet.getRange(3,1,1,4).mergeAcross()
  sheet.getRange('A4:D13').setBorder(true, true, true, true, true, true, '#b8f3f9', SpreadsheetApp.BorderStyle.SOLID)
  sheet.getRange('A4:D4').setBackground('#eaf2ff')

  // objectives
  sheet.getRange(3,5).setValue(headers_revision[1]).setFontSize(14).setFontWeight("bold").setBackground('#a2d2ff').setHorizontalAlignment('center').setBorder(true, true, true, true, true, true, '#b8f3f9', SpreadsheetApp.BorderStyle.SOLID)
  sheet.getRange(3,5,1,data[0].length - 4).mergeAcross()
  sheet.getRange(4,5, 10, data[0].length - 4).merge().setBackground('#ffffff').setBorder(true, true, true, true, true, true, '#b8f3f9', SpreadsheetApp.BorderStyle.SOLID)

  sheet.getRange(14,1,1, data[0].length).mergeAcross()

  // domain
  sheet.getRange(2,1).setValue('Domain').setFontWeight("bold").setBackground('#a2d2ff').setHorizontalAlignment('center').setBorder(true, true, true, true, true, true, '#b8f3f9', SpreadsheetApp.BorderStyle.SOLID)
  sheet.getRange(2,2,1,3).mergeAcross().setBorder(true, true, true, true, true, true, '#fea1b2', SpreadsheetApp.BorderStyle.SOLID)
  sheet.getRange(2,5,1,data[0].length - 4).mergeAcross().setBackground('#ffffff')

  //title 
  sheet.getRange(1,1).setValue('DATA BUILDER').setFontSize(18).setHorizontalAlignment('center').setFontWeight("bold").setBackground('#ffffff')
  sheet.getRange(1,1,1, data[0].length).mergeAcross()
  return sheet.getName()
}

//Inserts a new sheet and writes a 2D array of data in it
function writeDataToSheet(data, numNewColumn) {
  data.pop()

  let ss = SpreadsheetApp.getActive();
  let last_sheets = ss.getSheets()
  let indexSheet = 0

  let last_sheets_name = last_sheets[0].getName()
  let version = getVersion(last_sheets[0])
  let step = parseFloat(version)

  if(last_sheets.length > 1){
    step = parseFloat(getVersion(last_sheets[0])) - parseFloat(getVersion(last_sheets[1]))
  }

  let name = last_sheets_name.replace(version,'')
  version = parseFloat(version) + step
  version = version.toFixed(2)
  version = version.toString()

  if(parseInt(version[version.length-1]) < 1){
    version = parseFloat(version).toFixed(1)
  }

  return genUI(ss, data, name+version, indexSheet, numNewColumn)
}

function onEdit(e) {
  var ss = SpreadsheetApp.getActiveSheet();
  var range = ss.getRange(1, 1, ss.getMaxRows(), ss.getMaxColumns());
  range.setFontFamily('Barlow');
}

// const spreadsheetID = '1Y0ybJwpHoWE3as_4TFoHB5d-W2iLTRf8_h8WV8oI-6c';
// function addMenuItem() {
//   SpreadsheetApp.getUi()
//   .createMenu('My Functions')
//   .addItem('Show Dialog', 'Data Versioning')
//   .addToUi();
// }

// function showModal() {
//   const userInterface = HtmlService.createHtmlOutputFromFile('index');
//   SpreadsheetApp.getUi().showModelessDialog(userInterface, 'copyright © VA Team');
// }

function trackVersion(){
  
}

