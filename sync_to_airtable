function syncToAirtable() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("db.a");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var apiKey = PropertiesService.getScriptProperties().getProperty("AIRTABLE_API_KEY");
  var baseId = PropertiesService.getScriptProperties().getProperty("AIRTABLE_BASE_ID");
  var tableName = "db";
  var url = "https://api.airtable.com/v0/" + baseId + "/" + tableName;
  var headersObj = {
    "Authorization": "Bearer " + apiKey,
    "Content-Type": "application/json"
  };

  // Fetch Airtable field types dynamically
  var airtableFieldTypes = getAirtableFieldTypes(baseId, tableName, headersObj);

  // Check if field types were fetched successfully
  if (Object.keys(airtableFieldTypes).length === 0) {
    Logger.log("Error: Could not fetch Airtable field types. Exiting.");
    return;
  }

  // Fetch existing records from Airtable
  var airtableData = getAirtableData(url, headersObj);

  // Log the number of records fetched
  Logger.log("Number of records fetched from Airtable: " + airtableData.length);

  // Check if airtableData is valid
  if (!airtableData) {
    Logger.log("Error: Failed to fetch data from Airtable. Exiting.");
    return;
  }

  var airtableRecordMap = {};
  airtableData.forEach(function(record) {
    // Use lowercase Email as the key
    airtableRecordMap[record.fields["Email"]] = record.id;
  });

  var recordsToCreate = [];
  var recordsToUpdate = [];
  var recordsToDelete = [];

  // Store counts in variables
  var fetchedCount = airtableData.length;
  var updatedCount = 0;
  var createdCount = 0;
  var deletedCount = 0;

  // Loop through each row in the sheet (excluding header)
  for (var i = 1; i < data.length; i++) {
    var sheetRow = data[i];
    var record = {};
    var userEmail = sheetRow[0]; // Email is in the first column (index 0)

    // Check data types column by column
    for (var j = 0; j < headers.length; j++) {
      var fieldName = headers[j];
      var fieldValue = sheetRow[j];
      var airtableFieldType = airtableFieldTypes[fieldName];

      // Check if the field exists in Airtable
      if (!airtableFieldType) {
        Logger.log(`Warning: Field "${fieldName}" not found in Airtable. Skipping.`);
        continue;
      }

      // Convert data type to match Airtable field type
      var convertedValue = convertDataType(fieldValue, airtableFieldType);

      // Additional check to ensure data type matches after conversion
      if (typeof convertedValue !== getExpectedJsType(airtableFieldType) && fieldValue !== null) {
        Logger.log(`Error: Data type mismatch for field "${fieldName}". Expected ${airtableFieldType}, got ${typeof convertedValue}. Skipping this field.`);
        continue;
      }

      // Special handling for currency fields
      if (fieldName === "Last_Payment_Amount" || fieldName === "Customer_LTV") {
        // Ensure the value is a number
        if (typeof fieldValue !== "number") {
          fieldValue = parseFloat(fieldValue); // Try to convert to number
          if (isNaN(fieldValue)) {
            Logger.log(`Warning: Could not convert ${fieldName} to number for record ${userEmail}. Skipping this field.`);
            continue;
          }
        }
        // Send as number without currency formatting
        record[fieldName] = fieldValue;
      } else if (fieldName === "User_Segment" && fieldValue === null) {
        // Handle null values in User_Segment
        record[fieldName] = ""; // Or any other default value you prefer
      } else {
        record[fieldName] = convertedValue;
      }
    }

    // Use lowercase Email for lookup
    var existingRecordId = airtableRecordMap[userEmail];

    if (existingRecordId) {
      // Check for duplicates before updating (using only Email for comparison)
      var isDuplicate = Object.keys(airtableRecordMap).some(function(recordId) {
        // Skip the current record being updated
        if (recordId === existingRecordId) return false;

        var existingRecord = airtableRecordMap[recordId];
        return existingRecord.email === userEmail;
      });

      if (isDuplicate) {
        // Duplicate found, add to recordsToDelete
        recordsToDelete.push(existingRecordId);
        deletedCount++;
      } else {
        // Not a duplicate, update it
        recordsToUpdate.push({
          id: existingRecordId,
          fields: record
        });
        updatedCount++;
      }

      // Remove from map as it's being updated (or deleted)
      delete airtableRecordMap[userEmail];

    } else {
      // Check for duplicates within the records that are about to be created (using only Email for comparison)
      var isDuplicate = recordsToCreate.some(function(existingRecord) {
        return existingRecord.fields.Email === userEmail;
      });

      if (isDuplicate) {
        // Duplicate found, add to recordsToDelete
        var duplicateRecordId = airtableData.find(function(airtableRecord) {
          return airtableRecord.fields.Email === userEmail;
        })?.id;

        if (duplicateRecordId) {
          recordsToDelete.push(duplicateRecordId);
          deletedCount++;
        }
      } else {
        // Not a duplicate, create it
        recordsToCreate.push({
          fields: record
        });
        createdCount++;
      }
    }
  }

  // Delete records that are not in the sheet
  for (var remainingRecordId in airtableRecordMap) {
    recordsToDelete.push(airtableRecordMap[remainingRecordId]);
    deletedCount++;
  }

  // Perform batch operations
  if (recordsToCreate.length > 0) {
    batchCreateRecords(url, recordsToCreate, headersObj);
  }
  if (recordsToUpdate.length > 0) {
    batchUpdateRecords(url, recordsToUpdate, headersObj);
  }
  if (recordsToDelete.length > 0) {
    batchDeleteRecords(url, recordsToDelete, headersObj);
  }

  // Fetch updated record count from Airtable
  var finalAirtableData = getAirtableData(url, headersObj);
  var totalCount = finalAirtableData.length;

  // Call the deduplication function
  deduplicateAirtable();

   // Fetch updated record count from Airtable *after* deduplication
  var finalAirtableData = getAirtableData(url, headersObj);
  var totalCount = finalAirtableData.length;

  // Log the counts after deduplication
  Logger.log("Sync complete!");
  Logger.log("Deduplication complete!");
  Logger.log("Records fetched from Airtable: " + fetchedCount);
  Logger.log("Records updated in Airtable: " + recordsToUpdate.length);
  Logger.log("Records added to Airtable: " + createdCount);
  Logger.log("Records deleted from Airtable: " + deletedCount);
  Logger.log("Total records currently in Airtable: " + totalCount);
}

// Deduplication function (using only Email for deduplication)
function deduplicateAirtable() {
  var baseId = PropertiesService.getScriptProperties().getProperty("AIRTABLE_BASE_ID");
  var tableName = "db";
  var url = "https://api.airtable.com/v0/" + baseId + "/" + tableName;
  var headersObj = {
    "Authorization": "Bearer " + PropertiesService.getScriptProperties().getProperty("AIRTABLE_API_KEY"),
    "Content-Type": "application/json"
  };

  // Fetch all records from Airtable
  var airtableData = getAirtableData(url, headersObj);
  if (!airtableData || airtableData.length === 0) {
    Logger.log("No records found in Airtable.");
    return;
  }

  var existingRecords = {};
  var duplicateRecords = [];

  // Define field to check for duplicates (Email only)
  var idField = "Email";

  airtableData.forEach(record => {
    var key = record.fields[idField]; // Use lowercase email as key

    if (existingRecords[key]) {
      // Duplicate found, add the duplicate record ID to the list for deletion
      duplicateRecords.push(record.id);
    } else {
      existingRecords[key] = record;
    }
  });

  // Delete duplicate records in batches
  if (duplicateRecords.length > 0) {
    batchDeleteRecords(url, duplicateRecords, headersObj);
    Logger.log("Deleted " + duplicateRecords.length + " duplicate records.");
  } else {
    Logger.log("No duplicates found.");
  }
}

// Helper function to map Airtable field types to JavaScript types
function getExpectedJsType(airtableFieldType) {
  switch (airtableFieldType) {
    case "number":
    case "currency":
    case "percent":
      return "number";
    case "checkbox":
      return "boolean";
    default:
      return "string";
  }
}

// Fetch Airtable field types dynamically
function getAirtableFieldTypes(baseId, tableName, headersObj) {
  var metadataUrl = `https://api.airtable.com/v0/meta/bases/${baseId}/tables`;
  try {
    var response = UrlFetchApp.fetch(metadataUrl, { headers: headersObj, muteHttpExceptions: true });
    var httpCode = response.getResponseCode();
    if (httpCode !== 200) {
      Logger.log("Airtable Metadata API returned HTTP code: " + httpCode);
      Logger.log(response.getContentText());
      return {};
    }

    var data = JSON.parse(response.getContentText());

    if (data.error) {
      Logger.log("Airtable Metadata Error: " + JSON.stringify(data.error));
      return {};
    }

    var fieldTypes = {};
    var table = data.tables.find(t => t.name === tableName);
    if (table) {
      table.fields.forEach(field => {
        fieldTypes[field.name] = field.type;
      });
    }

    return fieldTypes;
  } catch (error) {
    Logger.log("Failed to fetch Airtable field types: " + error.message);
    return {};
  }
}

// Convert data type to match Airtable field type (date formatting removed)
function convertDataType(value, fieldType) {
  if (value === "" || value === null || value === undefined) return null;

  Logger.log(`Converting ${fieldType} value: ${value}`);

  // Special handling for "Name" field (single line text)
  if (fieldType === "Name") {
    return String(value); // Force conversion to string
  }

  switch (fieldType) {
    case "number":
    case "currency":
      var parsedValue = parseFloat(String(value).replace(/[^0-9.-]+/g, ""));
      if (isNaN(parsedValue)) {
        Logger.log("Warning: Could not parse number from value: " + value);
        return null;
      }
      return parsedValue;
    case "percent":
      return parseFloat(value) || 0;
    case "phoneNumber":
      return formatPhoneNumber(value);
    case "checkbox":
      return String(value) === "true";
    case "singleSelect":
      // For singleSelect, ensure the value is a string and trim whitespace
      var convertedValue = String(value).trim();
      Logger.log(`Converted ${fieldType} value: ${convertedValue}`);
      return convertedValue;
    case "multipleSelects":
      // For multipleSelects, ensure the value is an array of strings
      if (Array.isArray(value)) {
        return value.map(String).map(item => item.trim());
      } else {
        return String(value).split(",").map(item => item.trim());
      }
    case "email":
      return String(value).trim();
    default:
      return String(value);
  }
}

function formatPhoneNumber(value) {
  return String(value).replace(/\D/g, "").replace(/^\+/, ""); // Remove "+" prefix
}

// Fetch existing records from Airtable
function getAirtableData(url, headersObj) {
  var records = [];
  var offset;
  do {
    var fetchUrl = offset ? url + "?offset=" + offset : url;
    try {
      var response = UrlFetchApp.fetch(fetchUrl, { headers: headersObj, muteHttpExceptions: true });

      // Log the response code and content
      Logger.log("Response code: " + response.getResponseCode());
      Logger.log("Response content: " + response.getContentText());

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

function batchCreateRecords(url, records, headersObj) {
  const BATCH_SIZE = 10;
  for (let i = 0; i < records.length; i += BATCH_SIZE) {
    var batch = records.slice(i, i + BATCH_SIZE);
    var postOptions = { method: "POST", headers: headersObj, payload: JSON.stringify({ records: batch }) };
    try {
      var response = UrlFetchApp.fetch(url, postOptions);
      var httpCode = response.getResponseCode();
      if (httpCode !== 200) {
        Logger.log("Batch Create error. HTTP code: " + httpCode);
        Logger.log(response.getContentText());
      }
    } catch (e) {
      Logger.log("Batch Create error: " + e.message);
    }
    Utilities.sleep(200);
  }
}

function batchUpdateRecords(url, records, headersObj) {
  const BATCH_SIZE = 10;
  for (let i = 0; i < records.length; i += BATCH_SIZE) {
    var batch = records.slice(i, i + BATCH_SIZE);
    var patchOptions = { method: "PATCH", headers: headersObj, payload: JSON.stringify({ records: batch }) };
    try {
      var response = UrlFetchApp.fetch(url, patchOptions);
      var httpCode = response.getResponseCode();
      if (httpCode !== 200) {
        Logger.log("Batch Update error. HTTP code: " + httpCode);
        Logger.log(response.getContentText());
      }
    } catch (e) {
      Logger.log("Batch Update error: " + e.message);
    }
    Utilities.sleep(200);
  }
}

function batchDeleteRecords(url, recordIds, headersObj) {
  const BATCH_SIZE = 10; // Airtable allows deleting up to 10 records per request
  for (let i = 0; i < recordIds.length; i += BATCH_SIZE) {
    let batch = recordIds.slice(i, i + BATCH_SIZE);
    let deleteUrl = url + "?records=" + batch.join("&records="); // Corrected URL construction

    try {
      let response = UrlFetchApp.fetch(deleteUrl, {
        method: "DELETE",
        headers: headersObj,
        muteHttpExceptions: true
      });

      let httpCode = response.getResponseCode();
      if (httpCode !== 200) {
        Logger.log("Batch Delete error. HTTP code: " + httpCode);
        Logger.log(response.getContentText());
      }
    } catch (e) {
      Logger.log("Batch Delete error: " + e.message);
    }

    Utilities.sleep(200); // Add a delay to avoid hitting Airtable rate limits
  }
}
