function getjobs() {

  // Get the active sheet where job URLs are entered by the user
  let ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // Get last row number on the sheet
  let lr = ss.getLastRow();
  // Read the starting index/counter from cell A1 (used as loop start)
  let i = ss.getRange(1, 1).getValue();;

  // Loop from i to last row (process every row in the sheet)
  while (i <= lr) {

    // If column E (5) contains a value (user pasted an Upwork URL),
    // copy that URL to temporary column Y (25)
    if (ss.getRange(i, 5) != "") {
      let link = ss.getRange(i, 5).getValue();
      ss.getRange(i, 25).setValue(link);
    }

    // If temporary column Y (25) contains a value, extract the job_key
    // by splitting the URL after 'jobs/' and overwrite column Y with the job_key
    if (ss.getRange(i, 25) != "") {
      let splitUrl = ss.getRange(i, 25).getValue();
      splitUrl = splitUrl.split('jobs/')[1];
      ss.getRange(i, 25).setValue(splitUrl);
    }

      // Read job_key from column Y (25)
      let job_key = ss.getRange(i, 25).getValue();
      // Build Upwork Jobs API URL (JSON)
      let baseUrl = 'https://www.upwork.com/api/profiles/v1/jobs/'
      let format = '.json'
      let url = baseUrl + job_key + format

      // Get access token from the 'admin' sheet cell A1
      let ss1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('admin');
      let access_token = ss1.getRange(1, 1).getValue();

      // Additional UrlFetch options: GET, include Authorization header with Bearer token.
      // muteHttpExceptions:true allows reading non-200 responses for debugging/handling.
      const options = {
        method: 'GET',
        muteHttpExceptions: true,
        'headers': {
          accept: 'application/json',
          'Authorization': 'Bearer ' + access_token
        }
      }

      // Perform the HTTP request and parse JSON response
      let response = UrlFetchApp.fetch(url, options).getContentText();
      let result = JSON.parse(response);
 
      // Write obtained data into sheet cells for the current row (i)

      // Job title -> column F (6)
      let title = result.profile.op_title;
      ss.getRange(i, 6).setValue(title);

      // Job description -> column G (7). Use CLIP wrap strategy.
      let description = result.profile.op_description;
      ss.getRange(i, 7).setValue(description).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

      // Client country -> column R (18)
      let geo = result.profile.buyer.op_country;
      ss.getRange(i, 18).setValue(geo);
      
      // Client total spend -> column S (19). Align to the right.
      let tot_charge = result.profile.buyer.op_tot_charge;
      ss.getRange(i, 19).setValue(tot_charge).setHorizontalAlignment('right');;

      // Total hires -> column T (20)
      let tot_asgs = result.profile.buyer.op_tot_asgs;
      ss.getRange(i, 20).setValue(tot_asgs);
      
      // Job creation datetime (op_ctime) -> split on 'T' and write to column Z (26)
      let job_created_date = result.profile.op_ctime;
      job_created_date = job_created_date.split('T');
      ss.getRange(i, 26).setValue(job_created_date);
    
      // Extract time part from op_ctime (after 'T') and remove milliseconds -> column AA (27)
      let job_created_time = result.profile.op_ctime;
      job_created_time = job_created_time.split('T')[1];
      job_created_time = job_created_time.split('.');
      ss.getRange(i, 27).setValue(job_created_time);

      // Small time fragment slice from op_ctime -> column AD (30)
      let job_created_time1 = result.profile.op_ctime;
      job_created_time1 = job_created_time1.slice(10, 13);
      ss.getRange(i, 30).setValue(job_created_time1);

      // If date was written to column Z (26), set a formula in column AB (28)
      // that calculates the weekday name from that date.
      if (ss.getRange(i, 26) != "") {
        let weekday = '=TEXT(Z' + i + '; "dddd")';
        ss.getRange(i, 28).setValue(weekday); // The script takes the publication date from the cell and returns the day of the week
      }

      // If total spend (S / 19) exists, calculate average check = total_spent / hires and write to U (21)
      if (ss.getRange(i, 19) != "") {
        let total_spent = ss.getRange(i, 19).getValue();
        let hires = ss.getRange(i, 20).getValue();
        let average_check = total_spent / hires;
        ss.getRange(i, 21).setValue(average_check);
      }
      
      // Highlight weekday cell AB (28) with green for business days and red for weekends.
      const weekday = ss.getRange(i, 28).getValue();

      if (weekday === "Monday" || weekday === "Tuesday"|| weekday === "Wednesday" || weekday === "Thursday" || weekday === "Friday") {
        // Business days -> light green background
        ss.getRange(i, 28).setBackground('#b7e1cd'); // Do something if the cell contains data
    
      } else {
        // Weekends -> light red background
        (weekday === "Saturday"|| weekday === "Sunday")
        ss.getRange(i, 28).setBackground('#f4cccc'); // Do something if the cell contains data
      }
            
    // Move to the next row
    i++;
  };

}

