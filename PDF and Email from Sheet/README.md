<h1>📄 PDF Generator & Email Sender (Google Apps Script)</h1>

<p>
This script automatically generates PDF files based on data from Google Sheets and sends them to users via email.  
The project demonstrates working with HTML templates, spreadsheet processing, PDF generation, and email delivery using Google Apps Script.
</p>

<hr>

<h2>🚀 Features</h2>

<ul>
  <li><strong>Reads data from Google Sheets</strong>
    <ul>
      <li>Works with a sheet named <code>data</code>.</li>
      <li>The first row is treated as headers and skipped.</li>
    </ul>
  </li>

  <li><strong>Generates personalized HTML content</strong>
    <ul>
      <li>Each row produces an HTML template containing first name, last name, and company name.</li>
    </ul>
  </li>

  <li><strong>Converts HTML into PDF</strong>
    <ul>
      <li>The script creates an HTML Blob and converts it into a PDF file.</li>
    </ul>
  </li>

  <li><strong>Sends an email to the recipient</strong>
    <ul>
      <li>The generated PDF is attached to the email.</li>
      <li>The HTML content is included in the email body.</li>
    </ul>
  </li>

  <li><strong>Updates send status</strong>
    <ul>
      <li>Places the value <code>sent</code> into column E after successful email delivery.</li>
    </ul>
  </li>
</ul>

<hr>

<h2>📊 Google Sheets Data Structure</h2>

<p>The table must contain the following columns:</p>

<table>
  <tr><th>Column</th><th>Description</th></tr>
  <tr><td>A</td><td>First Name</td></tr>
  <tr><td>B</td><td>Last Name</td></tr>
  <tr><td>C</td><td>Company</td></tr>
  <tr><td>D</td><td>Recipient Email</td></tr>
  <tr><td>E</td><td>Send Status</td></tr>
</table>
<img width="682" height="183" alt="image" src="https://github.com/user-attachments/assets/aebcaf4e-3a6c-4608-9b0f-5f2b19cfc7c6" />
<hr>

<h2>🧩 How the Script Works</h2>

<h3>1. Open sheet by ID:</h3>
<pre><code>const ss = SpreadsheetApp.openById(id).getSheetByName('data');
</code></pre>

<h3>2. Get all data, skipping headers:</h3>
<pre><code>const data = ss.getDataRange().getValues().slice(1);
</code></pre>

<h3>3. Generate HTML content:</h3>
<pre><code>let html = `<h1>${row[0]} ${row[1]}</h1>`;
html += `<div>Hello, is your company ${row[2]}?</div>`;
</code></pre>

<h3>4. Create a PDF from HTML:</h3>
<pre><code>const blob = Utilities.newBlob(html, MimeType.HTML)
  .setName(`${row[0]} ${row[1]}.pdf`);
</code></pre>

<h3>5. Send an email:</h3>
<pre><code>MailApp.sendEmail({
  to: email,
  subject,
  htmlBody: html,
  attachments: [blob.getAs(MimeType.PDF)],
});
</code></pre>

<h3>6. Update send status:</h3>
<pre><code>ss.getRange(ind + 2, 5).setValue('sent');
</code></pre>

<hr>

<h2>🔧 Setup</h2>

<ol>
  <li>Create a Google Spreadsheet and add a sheet named <code>data</code>.</li>
  <li>Fill the sheet using the structure described above.</li>
  <li>Open <strong>Extensions → Apps Script</strong> and paste the script.</li>
  <li>Insert your sheet ID:
    <pre><code>const id = 'YOUR_SHEET_ID';</code></pre>
  </li>
  <li>Grant permissions when prompted.</li>
</ol>

<hr>

<h2>▶️ Run the Script</h2>

<p>Simply execute the function:</p>

<pre><code>pdfMaker();
</code></pre>

<p>The script will process all rows, generate personalized PDF files, and send emails automatically.</p>

<hr>

<h2>📬 Use Cases</h2>

<ul>
  <li>Mass mailing of documents to employees or customers</li>
  <li>Generating personalized PDFs (certificates, letters, notifications)</li>
  <li>HR workflow automation</li>
  <li>Automatic delivery of commercial offers</li>
</ul>




