function syncFeedbacksToSheet() {
  // 1. Script Properties
  var apiKey = PropertiesService.getScriptProperties().getProperty("AIRTABLE_API_KEY");
  var baseId = PropertiesService.getScriptProperties().getProperty("AIRTABLE_BASE_ID");
  var tableName = "Feedbacks";

  if (!apiKey || !baseId) {
    Logger.log("Error: Airtable API key or Base ID not found in Script Properties. Please set them up.");
    return;
  }

  // 2. Get the Google Sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  if (!sheet) {
    Logger.log("Error: Sheet 'Sheet1' not found. Please create it or adjust the sheet name.");
    return;
  }

  // 3. Airtable API URL
  var url = "https://api.airtable.com/v0/" + baseId + "/" + tableName;
  var headersObj = {
    "Authorization": "Bearer " + apiKey,
    "Content-Type": "application/json"
  };

  // 4. Specify the date-stamp field and Fetch ALL records from Airtable
  var dateStampField = "Date-stamp"; // Replace with your date-stamp field name
  var airtableData = getAirtableData(url, headersObj, dateStampField);

  if (!airtableData || airtableData.length === 0) {
    Logger.log("No records found in Airtable.");
    return;
  }

  // 5. Omit Specific Fields
  var omittedFields = ["Email address", "User", "Date"]; // Add the fields you want to exclude

  // 6. Prepare Data for Sheet
  var sheetData = [];
  var headers = [];

  // Extract headers from the first record (if available)
  if (airtableData[0] && airtableData[0].fields) {
    Object.keys(airtableData[0].fields).forEach(function(field) {
      if (!omittedFields.includes(field)) {
        headers.push(field);
      }
    });

    sheetData.push(headers); // Add headers as the first row
  }

  // Extract data for each record
  airtableData.forEach(function(record) {
    var row = [];
    headers.forEach(function(header) {
      row.push(record.fields[header] || "");
    });
    sheetData.push(row);
  });

  // 7. Write Data to Google Sheet (Batch Operations)
  // Clear columns A to H only
  var lastRow = sheet.getLastRow();
  var clearRange = sheet.getRange(1, 1, lastRow, 9); // Clear from A to I (9 columns)
  clearRange.clearContent();

  const BATCH_SIZE = 500; // Adjust batch size as needed, but be mindful of Google Sheets limits
  for (let i = 0; i < sheetData.length; i += BATCH_SIZE) {
    let batch = sheetData.slice(i, Math.min(i + BATCH_SIZE, sheetData.length));
    let range = sheet.getRange(i + 1, 1, batch.length, headers.length); // Adjust row start for each batch
    range.setValues(batch);
  }

  Logger.log("Data sync from Airtable to Google Sheet complete.");
}

// Helper function to fetch data from Airtable (same as before)
function getAirtableData(url, headersObj, sortField) {
  var records = [];
  var offset;
  do {
    var fetchUrl = url;
    if (offset) {
      fetchUrl += "?offset=" + offset;
    }
    if (sortField) {
      fetchUrl += (offset ? "&" : "?") + "sort[0][field]=" + sortField + "&sort[0][direction]=asc"; // Sort by the provided field
    }

    try {
      var response = UrlFetchApp.fetch(fetchUrl, {
        headers: headersObj,
        muteHttpExceptions: true,
      });

      var data = JSON.parse(response.getContentText());

      if (data.error) {
        Logger.log("Airtable Error: " + JSON.stringify(data.error));
        return;
      }

      records = records.concat(data.records);
      offset = data.offset;
    } catch (e) {
      Logger.log("Request failed: " + e.message);
      return;
    }
  } while (offset);

  return records;
}
