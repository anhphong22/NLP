
function onOpen(e) {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Ftech-DVC ⬇️")
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
  let url = 'https://raw.githubusercontent.com/anhphong22/NLP/main/pol_test.csv'
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
  auxiliary()
  sendtoTracker()
}

//Return a multi-dimensions array
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

/**
 * Displays an HTML-service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API.
 */
function showPicker() {
  var html = HtmlService.createHtmlOutputFromFile('Picker.html')
      .setWidth(600)
      .setHeight(425)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Pick csv file');
}

//Get Admin token
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}

/*
Function that takes item Id from Picker.html once user has made selection.
Creates clickable Url in spreadsheet.
Pastes in item Id to spreadsheet.
*/

function insertFileURL(id){
  // get Google Drive file by Id from Picker
  let file = DriveApp.getFileById(id);
  let new_column = ['question', 'answer', 'status']

  //insert file
  let headers = Utilities.parseCsv(file.getBlob().getDataAsString())[0];
  headers = headers.concat(new_column)
  
  let contents = CSVToArray(file.getBlob().getDataAsString(), ',')
  contents[0] = headers;
  for(let row in contents){
    if(row > 0){
      contents[row] = contents[row].concat([ ...Array(new_column.length).keys() ].map( i => i = ''))
    }
  }
  let sheetName = writeDataToSheet(contents, new_column.length);
  displayToastAlert("The CSV file was successfully imported into " + sheetName + ".");
  onEdit()
  auxiliary()
  sendtoTracker()
}

//Imports a CSV file in Google Drive into the Google Sheet
function importCSVFromDrive() {
  showPicker()
}

//Prompts the user for input and returns their response
function promptUserForInput(promptText) {
  let ui = SpreadsheetApp.getUi();
  let prompt = ui.prompt(promptText);
  let response = prompt.getResponseText();
  return response;
}

//Return version from a sheet
function getVersion(sheet){
  return sheet.getName().replace( /^\D+/g, '')
}

//Process text to clear
function process(text) {
  text = text.replace(/[0-9]/g, '');
  text = text.replace(/(\r\n\t|\n|\r)/gm, " ");
  text = text.replace(/[=]/g, " ");
  text = text.replace(/[:]/g, " ");
  text = text.replace(/[-]/g, " ");
  text = text.replace(/[>]/g, " ");
  text = text.replace(/[<]/g, " ");
  text = text.replace(/[@]/g, " ");
  text = text.replace(/\s+/g, ' ')
  text = text.replace(/[0-9]/g, ' ');
  text = text.replace("\\t ", "");
  text = text.replace("\n", "");
  text = text.replace("\n\t", "");
  text = text.replace("    ", "");
  text = text.toLocaleLowerCase();
  text = text.trim();
  text = text.trim();
  return text
}

// Return a dict contain the elements that be changed
function compareArrays (oldArray, newArray, data_primary_key){
    //init
    const result = {
        cells: [],
        columns: [],
        rows: []
    };
    //compare columns
    for(let i = 0; i < oldArray[0].length; i++) {
        if(i > newArray[0].length - 1) {
            result.columns.push({
                coor: [0, i],
                status: 'deleted',
                value: oldArray[0][i]
            });
            break;
        }
        if(i == oldArray[0].length - 1 && i < newArray[0].length - 1) {
            result.columns.push({
                coor: [0, i],
                status: 'added',
                value: newArray[0][i]
            });
            break;
        }
        if(oldArray[0][i].toString() !== newArray[0][i].toString()){
          result.columns.push({
                coor: [0, i],
                status: 'added',
                value: newArray[0][i]
            });
        }
    }
    //compare cells
    for (let i = 1; i < oldArray.length; i++) {
        if(newArray[i]){
          // console.log(newArray[i][0], data_primary_key, data_primary_key.indexOf(newArray[i][0]) == -1)
          // compare rows
          if(data_primary_key.indexOf(newArray[i][0]) == -1){
            result.rows.push({
                coor: [i, 0],
                status: 'added',
                value: newArray[i]
            });
          }else{
            for (let j = 1; j < newArray[i].length-1; j++) {
              // compare columns
              if(i-1 == 0 && oldArray[0].includes(newArray[0][j]) == -1){
                result.columns.push({
                    coor: [0, j],
                    status: 'added',
                    value: newArray[0][j-1]
                });
              }
              let index_key = data_primary_key.indexOf(newArray[i][0])
              if (index_key != -1) {  
                if (oldArray[index_key][j].toString().includes("GMT")) {
                  console.log(oldArray[index_key][j])
                  let new_date = new Date(oldArray[index_key][j]).toLocaleDateString("vi-VN")
                  oldArray[index_key][j] = new_date.toString()
                }
                if (oldArray[index_key][j].length == 0 && newArray[i][j].length > 0) {
                  result.cells.push({
                    coor: [i, j],
                    status: 'updated',
                    value: newArray[i][j]
                  });
                }
                else if (oldArray[index_key][j].length > 0 && newArray[i][j].length == 0) {
                  result.cells.push({
                    coor: [i, j],
                    status: 'deleted',
                    value: newArray[i][j]
                  });
                }
                else if (oldArray[index_key][j].toString() != newArray[i][j].toString()) {
                  if (oldArray[index_key][j].toString().includes('/') && newArray[i][j].toString().includes('/')) {
                    console.log(j, '--->', oldArray[0][j], newArray[0][j], process(oldArray[index_key][j].toString()) === process(newArray[i][j].toString()),oldArray[index_key][j].toString().length, newArray[i][j].toString().length)
                    if(oldArray[index_key][j].toString().length < 30 && newArray[i][j].toString().length < 30){
                      let diff_time = 0
                      let last_date = parseInt(oldArray[index_key][j].split('/')[0])
                      let now_date = parseInt(newArray[i][j].split('/')[0])
                      let last_month = parseInt(oldArray[index_key][j].split('/')[1])
                      let now_month = parseInt(newArray[i][j].split('/')[1])
                      let last_year = parseInt(oldArray[index_key][j].split('/')[2])
                      let now_year = parseInt(newArray[i][j].split('/')[2])
                      if (last_date < now_date) {
                        diff_time = now_date - last_date
                      } else {
                        if (last_month < now_month) {
                          if ((last_date == 31 && last_month > 2) || (last_date == 28 && last_month == 2)) {
                            diff_time = 1
                          }
                        } else {
                          diff_time = last_date - now_date
                        }
                      }
                      if (diff_time > 1) {
                        result.cells.push({
                          coor: [i, j],
                          status: 'updated',
                          value: newArray[i][j]
                        });
                      }else{
                        if(last_year != now_year){
                          result.cells.push({
                            coor: [i, j],
                            status: 'updated',
                            value: newArray[i][j]
                          });
                        }
                        if(last_month != now_month && diff_time != 1 && diff_time != 0){
                          console.log(diff_time, oldArray[index_key][0], last_month, now_month, last_year, now_year)
                          result.cells.push({
                            coor: [i, j],
                            status: 'updated',
                            value: newArray[i][j]
                          });
                        }
                      }
                    }else{
                      if (process(oldArray[index_key][j].toString()) != process(newArray[i][j].toString())){
                        result.cells.push({
                          coor: [i, j],
                          status: 'updated',
                          value: newArray[i][j]
                        });
                      }
                    }
                  } else {
                    result.cells.push({
                      coor: [i, j],
                      status: 'updated',
                      value: newArray[i][j]
                    });
                  }
                }
              }
            }
          }
        }
    }
    return result;
}

// Return description for each elements (cells, row, columns) from history
function genDes(name_element, data, ss){
  let des = ''
  let count_updated_cells = 0
  let count_added = 0
  let count_deleted_cells = 0
  let last_col = ss.getSheets()[0].getLastColumn()
  for (let el in data) {
    if (data[el].coor[1]+1 != last_col){
      if (data[el].status == 'updated') {
        count_updated_cells += 1
      }
      if (data[el].status == 'deleted') {
        count_deleted_cells += 1
      }
      if (data[el].status == 'added') {
        count_added += 1
      }
    }
  }
  if (count_updated_cells > 0) {
    des += 'updated ' + count_updated_cells + ' ' + name_element +'\n'
  }
  if (count_added > 0) {
    des += 'added ' + count_added + ' ' + name_element +'\n'
  }
  if(count_deleted_cells > 0){
    des += 'deleted '+ count_deleted_cells + ' ' + name_element +'\n'
  }
  return des
}

// Return history change sheet (combine two last sheets)
function getHistory(ss, data, numNewColumn, nameSheet){
  let last_sheets = ss.getSheets()[1]
  // console.log(last_sheets.getName(),  nameSheet, data.length)
  if(last_sheets.getName() != 'main' && last_sheets.getName() != 'Main' && last_sheets.getName() != nameSheet){
    let last_data = last_sheets.getRange(15,1,data.length,data[0].length).getValues()
    // console.log(data.length, last_data.length, last_data[0].length)
    let primary_keys = []
    let last_keys = last_sheets.getRange(15,1,data.length, 1).getValues()
    for(let k in last_keys){
      primary_keys.push(last_keys[k][0])
    }
    let history = compareArrays(last_data, data, primary_keys)
    // console.log(history.rows)
    let des = ''
    if(history.cells.length > 0){
      des += genDes('cells', history.cells, ss)
    }
    if(history.columns.length > 0){
      des += genDes('columns', history.columns,ss)
    }
    if(history.rows.length > 0){
      des += genDes('rows', history.rows, ss)
    }
    if(des.length == 0){
      return ['nothing changed', []]
    }else{
      return [des, history]
    }
  }
  else{
    return ['nothing changed', []]
  }
}

// Generate UI for each inserted sheet 
function genUI(ss, data, nameSheet, indexSheet, numNewColumn){
  sheet = ss.insertSheet(nameSheet, indexSheet)

  sheet.getRange(15, 1, 1, data[0].length-numNewColumn).setBackground("#f1f1f1")
  sheet.getRange(15, 1, 1, data[0].length).setFontWeight("bold")
  sheet.getRange(15, 1, 1, data[0].length).setHorizontalAlignment("center")
  sheet.getRange(15, 1, 1, data[0].length).setBorder(true, true, true, true, true, true, '#b8f3f9', SpreadsheetApp.BorderStyle.SOLID)

  sheet.getRange(15, data[0].length-numNewColumn+1, 1, numNewColumn).setBackground("#46bcd6")
  sheet.getRange(15, data[0].length, 1, 1).setBackground("orange")

  sheet.getRange(15, 1, data.length, data[0].length).setValues(data);
  sheet.getRange(16, 1, data.length-1, data[0].length).setHorizontalAlignment("left")
  sheet.getRange(16, 1, 1, data[0].length).setBorder(null, true, true, true, true, true, '#b8f3f9', SpreadsheetApp.BorderStyle.DASHED)
  sheet.getRange(17, 1, data.length-1, data[0].length).setBorder(true, true, true, true, true, true, '#b8f3f9', SpreadsheetApp.BorderStyle.DASHED)

  sheet.setRowHeightsForced(15, data.length, 24);
  sheet.autoResizeColumn(1);

  sheet.hideColumns(sheet.getLastColumn()+1, sheet.getMaxColumns()-sheet.getLastColumn())
  sheet.hideRows(sheet.getLastRow()+1, sheet.getMaxRows()-sheet.getLastRow())

  let history = getHistory(ss, data, numNewColumn, nameSheet)
  let headers_revision = ["Revision History", "Objectives"]
  let headers_rhistory = ["Version","Author","Issued","Descriptions"]
  let table_report = []
  let userEmail = Session.getActiveUser().getEmail()//PropertiesService.getUserProperties().getProperty("userEmail")
  table_report.push(headers_rhistory)
  let last_history_des = ss.getSheets()[1].getRange(5,1, ss.getSheets().length-1, headers_rhistory.length).getValues()
  if(last_history_des.length > 8){
    last_history_des.splice(last_history_des.length-9,8)
  }
  for(let r in last_history_des){
    if(last_history_des[r].length > 2 && last_history_des[r][0].length > 2){
      table_report.push(last_history_des[r])
    }
  }
  table_report.push([nameSheet, userEmail, new Date().toLocaleDateString("vi-VN"), history[0].trim('\n')])
  // console.log(table_report, table_report.length)
  sheet.getRange(4,1, table_report.length, table_report[0].length).setValues(table_report)
  
  // Limn the cells be changed
  for(let cell in history[1].cells){
    if(history[1].cells[cell].coor[1]+1 != sheet.getLastColumn()){
      if(history[1].cells[cell].status == 'updated'){
        sheet.getRange(history[1].cells[cell].coor[0]+15, history[1].cells[cell].coor[1]+1, 1, 1).setBackground('#77e093')
      }
      if(history[1].cells[cell].status == 'deleted'){
        sheet.getRange(history[1].cells[cell].coor[0]+15, history[1].cells[cell].coor[1]+1, 1, 1).setBackground('#f87567')
      }
      sheet.getRange(history[1].cells[cell].coor[0]+15, sheet.getLastColumn(), 1, 1).setValue(history[1].cells[cell].status)
    }
  }
  // Limn the rows be changed
  for(let row in history[1].rows){
    if(history[1].rows[row].status == 'added'){
      sheet.getRange(history[1].rows[row].coor[0]+15, history[1].rows[row].coor[1]+1, 1, history[1].rows[row].value.length).setBackground('#77e093')
    }
    if(history[1].rows[row].status == 'deleted'){
      sheet.getRange(history[1].rows[row].coor[0]+15, history[1].rows[row].coor[1]+1, 1, history[1].rows[row].value.length).setBackground('#f87567')
    }
    sheet.getRange(history[1].rows[row].coor[0]+15, sheet.getLastColumn(), 1, 1).setValue('row '+history[1].rows[row].status)

  }
  // Limn the columns be changed
  for(let column in history[1].columns){
    if(history[1].columns[column].status == 'added'){
      sheet.getRange(history[1].columns[column].coor[0]+15, history[1].columns[column].coor[1]+1, 1, 1).setBackground('#77e093')
    }
    if(history[1].columns[column].status == 'deleted'){
      sheet.getRange(history[1].columns[column].coor[0]+15, history[1].columns[column].coor[1]+1, 1, 1).setBackground('#f87567')
    }
    sheet.getRange(history[1].columns[column].coor[0]+15, sheet.getLastColumn(), 1, 1).setValue('column '+history[1].columns[column].status)
  }  

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
  sheet.getRange(1,2).setValue('DATA BUILDER').setFontSize(30).setHorizontalAlignment('center').setFontWeight("bold").setBackground('#ffffff')
  sheet.getRange(1,2,1, data[0].length).mergeAcross()
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

  if(last_sheets_name == 'Main' || last_sheets_name == 'main'){
    version = '0.1'
    name = 'ver '
    let name_newsheet = genUI(ss, data, name+version, indexSheet, numNewColumn)
    ss.deleteSheet(ss.getSheets()[1])
    return name_newsheet
  }else{
    return genUI(ss, data, name+version, indexSheet, numNewColumn)    
  }
}

function onEdit(e) {
  var ss = SpreadsheetApp.getActiveSheet();
  var range = ss.getRange(1, 1, ss.getMaxRows(), ss.getMaxColumns());
  range.setFontFamily('Barlow');
}

function auxiliary(){
  let ss = SpreadsheetApp.getActive();
  let domain_register = ss.getName()
  let ref_sheet = (ss.getUrl() + '#grid=' + ss.getSheetId())

  let sheet = ss.getSheets()[0]
  let objectives = sheet.getRange(4,5).getValue()

  sheet.getRange(2,2).setValue(domain_register)
  let blob = DriveApp.getFileById('1TUMw7PEpMggoqqsxUElUccVM8uVBjT60')
  let img = sheet.insertImage(blob, 1,1)
  let width = img.getWidth()
  let height = img.getHeight()
  img.setWidth(width).setHeight(height)
  sheet.setColumnWidth(1, width).setRowHeight(1,height)
  sheet.insertImage(blob, 1,1)
  return [domain_register, sheet.getName(),objectives, ref_sheet]

}


function sendtoTracker(){
  const sheet = SpreadsheetApp.openById('1rtm6k3xSUsZGtaqNLxB_vGBIQIkkkFhlryY2u4zT2EQ').getActiveSheet();

  let domain_registered = sheet.getRange('B8:B').getValues().filter(String);
  const lrow = sheet.getLastRow();
  let Avals= sheet.getRange("B1:B"+lrow).getValues();
  let Alast = '';

  if (domain_registered.length == 0){
    Alast = lrow - Avals.reverse().findIndex(c=>c[0] != '') + 1
    sheet.getRange(Alast, 2).setValue(auxiliary()[0])
  }
  else {
     for (let i= 0; i < domain_registered.length; i++){
        if (auxiliary()[0].toString() != domain_registered[i][0]) {
          Alast = lrow - Avals.reverse().findIndex(c=>c[0] != '') + 1;
          sheet.getRange(Alast, 2).setValue(auxiliary()[0]);
        }
        else{
          Alast = lrow - Avals.reverse().findIndex(c=>c[0] == auxiliary()[0].toString());
        }
  

    }
  }

  let issue_date = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy');
  sheet.getRange(Alast,3).setValue(auxiliary()[1])
  sheet.getRange(Alast,4).setValue(issue_date)
  if (auxiliary()[2].toString() == ''){
    sheet.getRange(Alast,5).setValue('The data was uploaded without any description')
  }
  else{
    sheet.getRange(Alast,5).setValue(auxiliary()[2])
  }
  sheet.getRange(Alast,6).setValue(auxiliary()[3])
  sheet.getRange(Alast,1).setValue(Alast-7)

}





  

