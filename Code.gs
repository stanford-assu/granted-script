function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Import GrantEd Data', functionName: 'UploadFile'},
  ];
  spreadsheet.addMenu('GrantEd', menuItems);
}

function UploadFile(){
  var htmlOutput = HtmlService.createHtmlOutputFromFile('form.html').setWidth(300).setHeight(150);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'Upload GrantEd Report');
}

function getFileText(data, file, name, email) {
  //process upload
  var contentType = data.substring(5,data.indexOf(';'));
  bytes = Utilities.base64Decode(data.substr(data.indexOf('base64,')+7));
  blob = Utilities.newBlob(bytes, contentType, file);
  var html = blob.getDataAsString();
  //parse upload into DOM
  const $ = Cheerio.load(html);
  //convert DOM to array
  var data = [];
  $('tr').each(function (rowIndex, r) {
      var cols = [];
      $(this).find('th,td').each(function (colIndex, c) {
          cols.push($(this).text().trim());
      });
      data.push(cols);
  });
  //pull out data
  var keys = data[1];
  data = data.slice(2);
  
  appwide = [0,1,2,3,4,5,6,7,8,9] // these columns describe the application
  linewide = [10,11,12,13,14] // these columns describe the 
  
  //PROCESS INTO DATA STRUCTURE
  //apps{
  //      "UGS-000001" {
  //         k : v,
  //         lines : {
  //              k : v
  //         }
  //      }
  //    }
 
  var apps = {};
  
  for(i in data){
      var id = data[i][2];
      //this only runs for the first line of each app
      if(!(id in apps)){
          app = {}
          for(j in appwide){
              app[keys[appwide[j]]] = data[i][appwide[j]];
          }
          app["lines"] = [];
          apps[id] = app;
      }
      //this is added for every line
      line = {};
      for(j in linewide){
          line[keys[linewide[j]]] = data[i][linewide[j]];
      }
      apps[id]["lines"].push(line);
  }
  
  //print
  //SpreadsheetApp.getUi().alert("Data: " + JSON.stringify(apps));
  
  //ADD TO SHEET
 
  //Wipe Sheet!
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var range = sheet.getDataRange();
  range.setNumberFormat("0.########");
  range.clear();
  sheet.clear();
  sheet.setFrozenRows(2);
  
  //Generate values
  var values = [];
  var appLines = []; // the numbers of the first lines of apps
  
  values.push(["Application","Group","####","Application Title","Line Item","Request","Recommended","Description"]);
  values.push(["=COUNTA(A3:A)","","","","","=SUM(F3:F)/2","=SUM(G3:G)/2",""]);
  
  for (id in apps) {
    appLines.push(values.length + 1);
    var line = []
    var applines = apps[id]["lines"];
    line.push(id);
    line.push(apps[id]["Organization Name"]);
    line.push(apps[id]["ASSU Number"]);
    line.push(apps[id]["Title"]);
    line.push("Total");
    line.push("=SUM(F" + (values.length+2) + ":F" + (values.length + applines.length + 1) + ")");
    line.push("=SUM(G" + (values.length+2) + ":G" + (values.length + applines.length + 1) + ")");
    line.push("");
    values.push(line);
    for(i in applines){
      var line = ["","","",""]
      line.push(acctToLine(applines[i]["Account"]));
      line.push(applines[i]["Line Requested Amount"]);
      line.push(applines[i]["Line Recommended Amount"]);
      line.push(applines[i]["Line Description"]);
      values.push(line);
    }
    values.push(["","","","","","","",""]);
  }
  
  var range = sheet.getRange(1, 1, values.length, 8);
  //save to sheet
  range.setValues(values);
  
  sheet.getRange("F1:F").setNumberFormat("\"$\"#,##0.00");
  sheet.getRange("G1:G").setNumberFormat("\"$\"#,##0.00");
  sheet.getRange("C1:C").setHorizontalAlignment("center");
  
  var headerStyle = SpreadsheetApp.newTextStyle().setBold(true).setUnderline(true).setFontSize(10).build();
  sheet.getRange("A1:1").setTextStyle(headerStyle);
  sheet.getRange("A1:1").setHorizontalAlignment("center");
  
  
  var appStyle = SpreadsheetApp.newTextStyle().setBold(true).setFontSize(10).build();
  
  for(i in appLines){
    sheet.getRange(appLines[i], 1, 1, 8).setTextStyle(appStyle);
  }
  
}

function acctToLine(acct){
  subacct = acct.split("-")[3];
  
  switch(subacct){
    case "2900":
      return "Honoraria";
    case "2910":
      return "Event Services";
    case "2920":
      return "Event Food & Supplies";
    case "2930":
      return "Marketing & Printing";
    case "2940":
      return "Costumes & Uniforms";
    case "2945":
      return "Storage";
    case "2950":
      return "Office Supplies";
    case "2955":
      return "Equipment";
    case "2960":
      return "Meeting Food";
    case "2970":
      return "Royalties";
    case "2980":
      return "Registration & Tickets";
    case "2990":
      return "Travel";
    default:
      return subacct;
  }
}