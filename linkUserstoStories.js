// Linking function (users to stories) in Airtable using Google Apps Script

function linkStoriesToUsersOptimized() {
  // Retrieve API key and Base ID from script properties
  const AIRTABLE_API_KEY = PropertiesService.getScriptProperties().getProperty("AIRTABLE_API_KEY");
  const BASE_ID = PropertiesService.getScriptProperties().getProperty("AIRTABLE_BASE_ID");
  
  // Define Airtable table names
  const STORIES_TABLE = "Table 1"; 
  const USERS_TABLE = "Table 2"; 

  // Set up authorization headers for Airtable API requests
  const headers = {
    "Authorization": "Bearer " + AIRTABLE_API_KEY,
    "Content-Type": "application/json"
  };

  // Function to fetch records from a given table with optional filtering
  function fetchRecords(table, filterByFormula = null) {
    let records = [];
    let offset = null;

    do {
      let url = `https://api.airtable.com/v0/${BASE_ID}/${table}`;
      let params = {};
      if (offset) params.offset = offset;
      if (filterByFormula) params.filterByFormula = filterByFormula;

      if (Object.keys(params).length > 0) {
        url += "?" + Object.keys(params).map(key => encodeURIComponent(key) + "=" + encodeURIComponent(params[key])).join("&");
      }

      let response = UrlFetchApp.fetch(url, { headers });
      let json = JSON.parse(response.getContentText());

      records = records.concat(json.records);
      offset = json.offset;

    } while (offset);

    Logger.log(`Fetched ${records.length} records from ${table}`);
    return records;
  }

  // Fetch all user records from the Users table
  const userRecords = fetchRecords(USERS_TABLE);

  // Create a mapping of user emails to their record IDs
  let userMap = {};
  userRecords.forEach(user => {
    let email = user.fields["Email"];
    if (email) {
      userMap[email] = user.id;
    }
  });

  // Fetch only stories that have a Customer Email and are not yet linked to a User
  const filterFormula = `AND({Customer Email} != "", NOT(User))`;
  const storyRecords = fetchRecords(STORIES_TABLE, filterFormula);

  let updates = [];
  // Iterate over story records and prepare updates where Customer Email matches a user email
  storyRecords.forEach(story => {
    let customerEmail = story.fields["Customer Email"];
    if (customerEmail && userMap[customerEmail]) {
      updates.push({
        id: story.id,
        fields: { "User": [userMap[customerEmail]] }
      });
      Logger.log(`Linking Customer Email ${customerEmail} to User ID ${userMap[customerEmail]}`);
    }
  });

  // Function to update Airtable records in batches (10 at a time)
  function batchUpdateRecords(updates) {
    let url = `https://api.airtable.com/v0/${BASE_ID}/${STORIES_TABLE}`;

    while (updates.length > 0) {
      let batch = updates.splice(0, 10);
      let payload = JSON.stringify({ records: batch });

      let response = UrlFetchApp.fetch(url, {
        method: "PATCH",
        headers,
        payload
      });

      Logger.log("Batch Update Response: " + response.getContentText());

      Utilities.sleep(200); // Prevent hitting rate limits
    }
  }

  // If there are updates to be made, perform batch updates
  if (updates.length > 0) {
    batchUpdateRecords(updates);
    Logger.log("All stories successfully linked to users!");
  } else {
    Logger.log("No updates needed.");
  }
}
