<h1>Upwork Jobs Importer — Google Apps Script</h1>

This project contains Google Apps Script utilities that automate importing job details from Upwork into a Google Spreadsheet.
The script reads job URLs from the sheet, extracts the job key, queries the Upwork API, and writes structured job data into row fields.
A second script refreshes the Upwork OAuth access token on a daily schedule using a refresh token.

Features
1. getjobs() — Main Upwork Job Importer

Reads Upwork job URLs from column E.

Extracts the job key from URLs (the part after /jobs/).

Sends an authorized GET request to the Upwork API:

https://www.upwork.com/api/profiles/v1/jobs/{job_key}.json


Writes job data into specific spreadsheet columns:

job title

description

client country

total client spend

number of hires

creation date / time

calculated average check

Applies conditional formatting:

business days highlighted green, weekends red.

2. requestAccessToken() — OAuth Token Refresher

Uses refresh_token to obtain a new access_token and refresh_token.

Updates the admin sheet automatically.

Intended to run daily via a Google Apps Script trigger.

Uses:

https://www.upwork.com/api/v3/oauth2/token

Requirements

Upwork account with a registered OAuth application

Client ID, Client Secret, and valid refresh_token

Google Spreadsheet with:

A sheet for job URLs (the script uses the active sheet)

A sheet named admin:

Cell	Meaning
admin!A1	access_token
admin!A2	refresh_token

Update the following fields inside the script:

let clientId = 'CLIENT_ID';
let clientSecret = 'CLIENT_SECRET';
let redirectUri = 'https://script.google.com/macros/d/YOUR_ID/usercallback';

Spreadsheet Column Structure

Below is the exact layout used by the script:

Column	Index	Purpose
A	1	Script counter (i)
E	5	User input — Upwork job URL
Y	25	Temporary: extracted job key
F	6	Job Title
G	7	Job Description
R	18	Client Country
S	19	Total Client Spend
T	20	Total Hires
U	21	Average Check (Spend / Hires)
Z	26	Job creation date
AA	27	Job creation time
AB	28	Day of week (=TEXT(Zi, "dddd"))
AD	30	Time fragment from Upwork timestamp
How to Install
1. Prepare the Google Sheet

Create at least two sheets:

Main sheet — where links are entered

admin — for storing tokens

Fill token cells:

admin!A1 — access_token  
admin!A2 — refresh_token

2. Add the Script

Open Google Sheets → Extensions → Apps Script

Paste both functions (getjobs and requestAccessToken)

Insert your Upwork credentials

Save

3. Grant Permissions

On the first run Google will request:

Spreadsheet read/write access

External API access (UrlFetchApp)

Approve.

4. Set Up Triggers

Go to:
Extensions → Apps Script → Triggers

Create:

Daily trigger → run requestAccessToken()

Optional: manual or scheduled trigger → run getjobs()
