// === Configuration: insert your OpenAI API key here ===
const API_KEY = 'YOUR_API_KEY';


// Main function — can be run manually or via a trigger
function replyGenerator() {
// OpenAI Completion API endpoint
const url = 'https://api.openai.com/v1/completions'


// Access active spreadsheet and Sheet1
const ss = SpreadsheetApp.getActive();
const sheet = ss.getSheetByName('Sheet1');


// Determine last row with data
let lr = ss.getLastRow();


// Loop through rows starting from row 2 (row 1 contains headers)
for (let i = 2; i <= lr; i++) {


// Get the incoming email text from column D
const data = sheet.getRange(i, 4).getValue();


// Build a prompt with instructions + email text
const prompt = 'Analyze the letter' + data + 'and write a response in one sentence'


// API parameters
const max_tokens = 4000;
const temperature = 0;
const model = "text-davinci-003"


// Send POST request to OpenAI via helper function r()
const request = r(url, 'post', API_KEY, {
"model": model,
"prompt": prompt,
"max_tokens": max_tokens,
"temperature": temperature
});


// Extract generated text from response
const response = request.choices?.[0]?.text || request;


// Write the reply into column E and enable wrap
sheet.getRange(i, 5).setValue(response).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}
}


// Helper function to perform HTTP requests (UrlFetchApp)
function r(url, method, token, data) {
// Request configuration
let params = {
method: method,
muteHttpExceptions: true,
contentType: 'application/json;',
payload: JSON.stringify(data),
'headers': {
Authorization: 'Bearer ' + token
}
};


// Execute request and parse JSON
var r = UrlFetchApp.fetch(url, params);
r = JSON.parse(r)
return r;
};