/*
Import CSV data to Google Spreadsheet

Based on:
http://collaborative-tools-project.blogspot.com.br/2012/05/getting-csv-data-into-google.html
http://www.bennadel.com/blog/1504-ask-ben-parsing-csv-strings-with-javascript-exec-regular-expression-command.htm

*/


function CSVToArray( strData, strDelimiter ){
  strDelimiter = (strDelimiter || ",");
  var objPattern = new RegExp(
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
  var arrData = [[]];
  var arrMatches = null;
  while (arrMatches = objPattern.exec( strData )){
        var strMatchedDelimiter = arrMatches[ 1 ];
        if (
            strMatchedDelimiter.length &&
            (strMatchedDelimiter != strDelimiter)
            ){
              arrData.push( [] );

        }

        if (arrMatches[ 2 ]){
            var strMatchedValue = arrMatches[ 2 ].replace(
                new RegExp( "\"\"", "g" ),
                "\""
                );

        } else {
            var strMatchedValue = arrMatches[ 3 ];
        }
        arrData[ arrData.length - 1 ].push( strMatchedValue );
    }
    return( arrData );
}


function encodeUTF8( s ){
  return unescape( encodeURIComponent( s ) );
}

function getCSV() { 
  var url = 'http://www.youserver.com/data.csv';
  var response = UrlFetchApp.fetch(url);
  //MailApp.sendEmail("youemail@host.com", "Error with CSV grabber:" + response.getResponseCode() , "Text of message goes here")
  //Logger.log( "RESPONSE " + response.getResponseCode()); 
  var data = encodeUTF8(response.getContentText().toString());   
  return data;
}

function importData(){
  var rawData = getCSV();
  var csvData = CSVToArray(rawData, ";");
  //Logger.log("CSV ITEMS " + csvData.length);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  //var sheet2 = ss.getSheets()[1];
  for (var i = 0; i < csvData.length; i++) {
    sheet.getRange(i+2, 1, 1, csvData[i].length).setValues(new Array(csvData[i]));
    //sheet2.getRange(i+2, 1, 1, csvData[i].length).setValues(new Array(csvData[i]));
  }
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Import Data",
    functionName : "importData"
  }];
  spreadsheet.addMenu("Custom Scripts", entries);
};
