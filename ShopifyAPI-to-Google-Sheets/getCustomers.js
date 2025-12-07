function getCustomers() {

  // URL for Shopify API request (replace with your store domain)
  let url = "https://api-to-google-sheets.myshopify.com/admin/api/2023-01/customers.json"

  // Request headers including your Shopify Access Token
  let headers = {
    'X-Shopify-Access-Token': 'YOUR_TOKEN'
  }

  // Request configuration: GET method and headers
  const options = {
    method: 'GET',
    muteHttpExceptions: true,
    headers: headers
  }

  // Fetching customer JSON data from Shopify
  let response = UrlFetchApp.fetch(url, options).getContentText();
  let result = JSON.parse(response);

  // Reference to the target sheet
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');

  // Clear the sheet before writing new data
  sheet.clear()

  // Header row for the spreadsheet
  let headerRow = ['Email', 'First name', 'Last name', 'Total spent', 'Phone', 'First name', 'Last name', 'Company', 'Address', 'City', 'Province', 'Country', 'ZIP', 'Phone', 'Province code', 'Country code', 'Country name']
  sheet.appendRow(headerRow)
  sheet.getRange('A1:Z1').activate().setFontWeight('bold');
  
  // Loop through all customers and write each to the sheet
  for(let i = 0; i<result.customers.length; i++){
    let row = [
      result.customers[i].email,
      result.customers[i].first_name, 
      result.customers[i].last_name, 
      result.customers[i].total_spent, 
      result.customers[i].phone, 
      result.customers[i].addresses[0].first_name, 
      result.customers[i].addresses[0].last_name, 
      result.customers[i].addresses[0].company, 
      result.customers[i].addresses[0].address1, 
      result.customers[i].addresses[0].city, 
      result.customers[i].addresses[0].province, 
      result.customers[i].addresses[0].country, 
      result.customers[i].addresses[0].zip, 
      result.customers[i].addresses[0].phone, 
      result.customers[i].addresses[0].province_code, 
      result.customers[i].addresses[0].country_code, 
      result.customers[i].addresses[0].country_name
    ]

    // Append the customer row
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers').appendRow(row)

    // Auto-resize columns for readability
    sheet.getRange('A:Z').activate();
    sheet.autoResizeColumns(1, 26);
  };
}
