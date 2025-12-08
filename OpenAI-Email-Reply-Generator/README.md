<h1>OpenAI Email Reply Generator</h1>

<h2>Description</h2>
<p>
  This Google Apps Script automates the generation of short, AI-powered replies to incoming emails stored inside Google Sheets.  
  It uses the OpenAI Completion API (<strong>model: text-davinci-003</strong>) to analyze an email and produce a one-sentence response.
</p>
<p>
  This project is suitable for a GitHub portfolio as an example of integrating Google Sheets, Apps Script, and OpenAI for workflow automation.
</p>

<h2>Features</h2>
<ul>
  <li>Reads email text from Google Sheets</li>
  <li>Sends a prompt to the OpenAI Completion API</li>
  <li>Writes the generated reply back into the sheet</li>
  <li>Easy setup — just add your API key and prepare a sheet</li>
</ul>

<h2>How It Works</h2>
<ol>
  <li>The script opens a Google Sheet named <strong>Sheet1</strong>.</li>
  <li>It loops through rows starting from row 2 (row 1 is expected to contain headers).</li>
  <li>For each row, it retrieves the email text stored in column D.</li>
  <li>A prompt is generated and sent to the OpenAI API.</li>
  <li>The generated reply is written into column E with text wrapping enabled.</li>
</ol>

<h2>Requirements</h2>
<ul>
  <li>A Google account with access to Google Apps Script</li>
  <li>An OpenAI API key (model availability may vary — this example uses <strong>text-davinci-003</strong>)</li>
</ul>

<h2>Spreadsheet Setup</h2>
<p>Your sheet <strong>Sheet1</strong> should look like this:</p>

<table border="1" cellspacing="0" cellpadding="6">
  <tr>
    <th>A</th>
    <th>B</th>
    <th>C</th>
    <th>D (Incoming Email)</th>
    <th>E (AI Reply)</th>
  </tr>
  <tr>
    <td>optional</td>
    <td></td>
    <td></td>
    <td>email text</td>
    <td>generated reply</td>
  </tr>
</table>

<p>Row 1 contains headers; data starts from row 2.</p>

<h2>Installation &amp; Execution</h2>
<ol>
  <li>Open Google Sheets → <em>Extensions → Apps Script</em>.</li>
  <li>Paste the script into the editor.</li>
  <li>Insert your OpenAI API key into:<br>
    <pre><code>const API_KEY = 'YOUR_API_KEY';</code></pre>
  </li>
  <li>Save the project and run the <strong>go()</strong> function.</li>
</ol>
<p><strong>Note:</strong> On the first run Google will request permissions (Sheets access and external API calls via <code>UrlFetchApp</code>).</p>

<h2>Security Notes</h2>
<ul>
  <li>Do not commit your API key to public repositories — use <code>PropertiesService</code> if needed.</li>
  <li>Avoid sending sensitive data to external APIs unless necessary.</li>
  <li>You may modify the model, temperature, or <code>max_tokens</code> depending on your needs and API limits.</li>
</ul>
