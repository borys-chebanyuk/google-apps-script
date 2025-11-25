function pdfMaker() {
  const id = 'YOUR_SHEET_ID';
  const ss = SpreadsheetApp.openById(id).getSheetByName('data');
  const data = ss.getDataRange().getValues().slice(1); //Get the entire data range and turn it into an array
  //.slice(1) - removes the first line (headings)

  //Next, we iterate through each row of the data array
  data.forEach((row, ind)=>{ 
    let html = `<h1>${row[0]} ${row[1]}</h1>`; //Create an HTML header with the first and last name
    html += `<div>Hello, is your company ${row[2]}?</div>`; //Add a line with the company name
    html += `<div>Thank You</div>`; //Adding the final text
    
    const blob = Utilities.newBlob(html,MimeType.HTML); 
    //Utilities.newBlob(content, contentType, name) — Creates a Blob object (Creating a Blob (a file in Google Apps Script memory) from an HTML string)
    //MIME— HTML type
    blob.setName(`${row[0]} ${row[1]}.pdf`); //We assign a name to the future PDF file

    const email = row[3]; //recipient's e-mail
    const subject = `New PDF ${row[0]} ${row[1]}`; //Forming the subject of the letter
    MailApp.sendEmail({
      to:email,
      subject:subject,
      htmlBody:html,
      attachments:[blob.getAs(MimeType.PDF)], // converts content to PDF
    }); //Sending an email

    const upRange = ss.getRange(ind+2,5).setValue('sent'); //We write the sent status in the 5th column.
    Logger.log(blob);
    })
}

