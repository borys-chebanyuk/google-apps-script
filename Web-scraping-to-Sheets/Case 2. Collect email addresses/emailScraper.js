function getEmailsFromPage() {
  // Get the 'Emails' sheet from the active spreadsheet
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Emails');

  // Get the last non-empty row in the sheet (used as the upper bound for the loop)
  let lr = ss.getLastRow();
  // Read the starting row index from cell F1 (row 1, column 6)
  let i = ss.getRange(1, 6).getValue()


  // Iterate from starting row 'i' up to the last row 'lr'
  while (i <= lr) {
    // Read the URL to fetch from column E (column index 5) of the current row
    let url = ss.getRange(i, 5).getValue()

    try {
      // Fetch the URL. muteHttpExceptions=true prevents the fetch from throwing on HTTP errors.
      let response2 = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
      // Get HTTP response code (e.g., 200, 404) and write it into column F (column index 6)
      let responseCode = response2.getResponseCode()
      ss.getRange(i, 6).setValue(responseCode);
      // Get the page content as text for regex parsing
      content = response2.getContentText();      
    } catch (e) {
      // If fetching the URL throws an exception (network error, invalid URL), log it and clear the status cell
      Logger.log("Error fetching URL: " + e);
      ss.getRange(i, 6).setValue(' ');
    } finally {

      // Regular expression to match email-like patterns
      let regex = /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/g;
      // Apply regex to the page content and collect matches
      let matches = content.match(regex);

      // If matches were found, process them
      if (matches && matches.length) {
        for (let u = 0; u < matches.length; u++) {
          // If the matched item is an empty string, write a blank into column A and skip
          if (matches[u] == "") {
            
            ss.getRange(i, 1).setValue(' ');
          } else {
        
            // Remove duplicate email addresses found on the same page
            let uniqueMatches = matches.filter(function(item, pos) {
              return matches.indexOf(item) == pos;
            });
            
            // Join unique emails into a single comma-separated string
            uniqueMatches = uniqueMatches.join(', ');
            // Write the comma-separated list into column A for the current row
            ss.getRange(i, 1).setValue(uniqueMatches)
            //Logger.log(uniqueMatches);
            
          }

          // NOTE: original code increments 'u' inside the loop in addition to the for-loop increment.
          // Keep the original logic intact (though this double increment effectively skips every other match).
          u++;
        }
      }
      
    }
    
    // Move to the next row
    i++;
  };

}
