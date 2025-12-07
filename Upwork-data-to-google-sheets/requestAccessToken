function requestAccessToken() {
  // 'admin' sheet holds tokens: access token in A1 and refresh token in A2
  let ss1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('admin');
  // Replace these with your actual Upwork OAuth credentials
  let clientId = 'CLIENT_ID';
  let clientSecret = 'CLIENT_SECRET';
  // Read current refresh token from admin!A2
  let refreshToken = ss1.getRange(2, 1).getValue();
  // Redirect URI registered in your OAuth app (for Apps Script example)
  let redirectUri = 'https://script.google.com/macros/d/YOUR_ID/usercallback';


  // Payload for token refresh request (grant_type=refresh_token)
  let payload = {
    'grant_type': 'refresh_token',
    'client_id': clientId,
    'client_secret': clientSecret,
    'refresh_token': refreshToken,
    'redirect_uri' : redirectUri
  };

  // HTTP POST options: application/x-www-form-urlencoded
  let options = {
    'method': 'post',
    'contentType': 'application/x-www-form-urlencoded',
    'payload': payload
  };

  // Request to Upwork OAuth2 token endpoint to obtain new tokens
  let response = UrlFetchApp.fetch('https://www.upwork.com//api/v3/oauth2/token', options);
  let result = JSON.parse(response.getContentText());

  // Extract new access_token and refresh_token from response
  let access_Token = result.access_token;
  let refresh_Token = result.refresh_token;

  // Log tokens for debugging (remove logs in production if necessary)
  Logger.log(access_Token);
  Logger.log(refresh_Token);

  // Write new tokens back to admin sheet: A1 = access_token, A2 = refresh_token
  ss1.getRange(1, 1).setValue(access_Token);
  ss1.getRange(2, 1).setValue(refresh_Token);

}
