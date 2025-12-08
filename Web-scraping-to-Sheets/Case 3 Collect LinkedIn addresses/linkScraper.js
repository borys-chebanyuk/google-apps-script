function getLinkedInLinks() {
  // Get the "LinkedIn" sheet from the active spreadsheet
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LinkedIn');

  // Get the last filled row in the sheet
  let lr = ss.getLastRow();

  // Starting row for processing (taken from cell D2)
  let i = ss.getRange(2, 4).getValue()

  // Loop through each row until the last one
  while (i <= lr) {

    // Get the URL from column A of the current row
    let url = ss.getRange(i, 1).getValue()

    try {
      // Fetch the page content and allow exceptions to be returned
      let response2 = UrlFetchApp.fetch(url, {muteHttpExceptions: true});

      // Save HTTP response status code into column C
      let responseCode = response2.getResponseCode()
      ss.getRange(i, 3).setValue(responseCode);

      // Extract the HTML/text content of the page
      content = response2.getContentText();      
    } catch (e) {

      // Log fetch errors and put blank value into column C
      Logger.log("Error fetching URL: " + e);
      ss.getRange(i, 3).setValue(' ');

    } finally {

      // Regex to detect LinkedIn profile/company links on the page
      let regex = /https?:\/\/(www\.)?linkedin\.com\/(in|company)\/[a-zA-Z0-9_-]+\/?/g;

      // Find all matches in the page content
      let matches = content.match(regex);

      // If matches exist, process them
      if (matches && matches.length) {

        for (let u = 0; u < matches.length; u++) {

          // If the match is empty, write blank into column B
          if (matches[u] == "") {
            ss.getRange(i, 2).setValue(' ');
          } else {

            // Remove duplicate LinkedIn links from the match list
            let uniqueMatches = matches.filter(function(item, pos) {
              return matches.indexOf(item) == pos;
            });
            
            // Convert the unique list into a comma-separated string
            uniqueMatches = uniqueMatches.join(', ');

            // Write the unique LinkedIn links into column B
            ss.getRange(i, 2).setValue(uniqueMatches)
            //Logger.log(uniqueMatches);
          }

          u++;
        }
      }
      
    }
    
    // Move to the next row
    i++;
  };

}
