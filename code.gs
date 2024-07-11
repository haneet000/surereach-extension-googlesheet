function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Surereach')
    .addItem('Start', 'doGet')
    .addToUi();
}

function doGet() {
  try {   
   
    var html = HtmlService.createHtmlOutputFromFile('finalchanges')
      .setWidth(600)
      .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'Surereach');
  } catch (error) {
    console.error('Error in doGet:', error);
    throw new Error('An error occurred while loading the web app. Please try again later.');
  }
}


function getHeaderValues() {
  console.log("Get Header")
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    
    if (!sheet) {
      throw new Error("No active sheet found");
    }
    
    var headerValues = [];
    var row = 1;
    var col = 1;
    var value = sheet.getRange(row, col).getValue();
    
    while (value !== "") {
      headerValues.push(value);
      col++;
      value = sheet.getRange(row, col).getValue();
    }
    
    Logger.log("Headers: " + headerValues.join(", "));
    return headerValues;
  } catch (error) {
    Logger.log("Error reading headers: " + error);
    return []; 
  }
}


function sleep(milliseconds) {
  var start = new Date().getTime();
  while ((new Date().getTime() - start) < milliseconds) {}
}


function showSheet(selectedHeader, accessToken) {
  try {
    var base_url = "https://api.surereach.io/api/v1/surereach/users/fetch-linkedin-data";
    var headers = {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + accessToken
    };
    
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getActiveSheet();
    
    if (sheet) {
      const values = sheet.getDataRange().getValues();
      const headerRow = values[0];
      var col = headerRow.indexOf(selectedHeader);

      if (col !== -1) {
        // Determine the index for the new column
        var lastColumnIndex = sheet.getLastColumn();
        var newColumnIndex = lastColumnIndex + 1;

        // Write the header for the new column
        var newColumnHeader = 'Phone Number'; // Change this to your desired header
        sheet.getRange(1, newColumnIndex).setValue(newColumnHeader);

        for (var rowIndex = 1; rowIndex < values.length; rowIndex++) {
          var linkedinUrl = values[rowIndex][col];

          if (linkedinUrl) {
            var payload = {"profile_url": linkedinUrl};
            
            try {
              Utilities.sleep(500);
              var response = UrlFetchApp.fetch(base_url, {
                'method': 'post',
                'headers': headers,
                'payload': JSON.stringify(payload)
              });
              
              if (response.getResponseCode() == 200) {
                var responseText = response.getContentText();
                var responseObj = JSON.parse(responseText);
                var phoneNumber = responseObj.data.phone_no;
                
                // Write the phone number to the new column
                sheet.getRange(rowIndex + 1, newColumnIndex).setValue(phoneNumber);
              } else if (response.getResponseCode() == 400 || response.getResponseCode() == 422) {
                sheet.getRange(rowIndex + 1, newColumnIndex).setValue("Invalid LinkedIn URL");
              } else {
                Logger.log("Unexpected response code: " + response.getResponseCode());
                sheet.getRange(rowIndex + 1, newColumnIndex).setValue("Error");
              }
            } catch (e) {
              sheet.getRange(rowIndex + 1, newColumnIndex).setValue("Error: " + e);
              Logger.log("Exception occurred: " + e);
            }
          } else {
            sheet.getRange(rowIndex + 1, newColumnIndex).setValue("LinkedIn URL not found");
          }
        }
      } else {
        Logger.log("Header not found: " + selectedHeader);
      }
    } else {
      Logger.log("Sheet not found");
    }
  } catch (error) {
    Logger.log("Error occurred: " + error);
  }
}



