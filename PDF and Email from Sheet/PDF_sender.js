function pdfMaker() {
  const id = '1D_YwWmtmIUZRHHV0etybQ8yVjfT4NrZY8AkN7_AiTVk';
  const ss = SpreadsheetApp.openById(id).getSheetByName('data');
  const data = ss.getDataRange().getValues().slice(1); //Получаем весь диапазон данных и превращаем его в массив
  //.slice(1) убирает первую строку (заголовки)

  //Перебираем каждую строку массива данных
  data.forEach((row, ind)=>{ 
    let html = `<h1>${row[0]} ${row[1]}</h1>`; //cоздаём HTML-заголовок с именем и фамилией
    html += `<div>Hello, is your company ${row[2]}?</div>`; //Добавляем строку с названием компании
    html += `<div>Thank You</div>`; //Добавляем финальный текст
    
    const blob = Utilities.newBlob(html,MimeType.HTML); //Utilities.newBlob(content, contentType, name) — создаёт объект Blob (Создаём Blob (файл в памяти Google Apps Script) из HTML-строки) //MIME-тип — HTML
    blob.setName(`${row[0]} ${row[1]}.pdf`); //Задаём имя будущему PDF файлу

    const email = row[3]; //e-mail получателя
    const subject = `New PDF ${row[0]} ${row[1]}`; //Формируем тему письма
    MailApp.sendEmail({
      to:email,
      subject:subject,
      htmlBody:html,
      attachments:[blob.getAs(MimeType.PDF)], // конвертирует содержимое в PDF
    }); //Отправляем письмо

    const upRange = ss.getRange(ind+2,5).setValue('sent'); //Записываем в 5-й столбец статус sent
    Logger.log(blob);
    })
}
