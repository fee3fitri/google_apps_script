function getData() {
  // Getting active sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('new list');

  // Extracting the data into an array 
  let lastColumn = sheet.getLastColumn();
  let lastRow = sheet.getLastRow();
  let values = sheet.getSheetValues(2, 1, lastRow, lastColumn);

  return values;
}

function sendEmail() {
  let values = getData();

  for (let i = 0; i < values.length; i++) {
    // Extracting data per column name
    const name = values[i][0];
    const firstName = name.split(' ')[0];
    const email = values[i][1];
    const company = values[i][2];
    const position = values[i][3];
    const job = values[i][4];
    const link = values[i][5];
    const resume = values[i][6]

    // HTML email
    const subject = `
      Hi ${firstName}, I\'m Fifi Shelton. ${job}. Check Me Out
    `

    const body = `
      <p>Hi ${firstName},<p>
      <p>Today I applied for the <a href="${link} target="_blank">${job}</a> position and I noticed that you are a ${position} at ${company}. While I do not know if you are the right person to connect with, I definitely have the front-end developer experience your team is seeking. I believe I would be a great fit because I have the necessary experience and I'm capable of quickly learning new technical skills.</p>
      <p>I imagine that you are really busy, but I would enjoy the opportunity to hop on a quick Zoom call to learn more about this role and the company.</p>
      <p>Would you possibly be free for a 15-minute Zoom call next week?</p>
      <p>In advance, I have attached my <a href="${resume}">resume</a> for your review. I appreciate your consideration and look forward to hearing from you.</p>
    `

    if (email === "") continue;

    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: body,
      name: 'Fifi Shelton'
    })
  }
}

// Move the data to sent list sheet
// Move the data to sent list sheet
function moveData() {
  const newSheet = getNewSheet();
  const sentSheet = getSentSheet();
  const sentLastCol = sentSheet.getLastColumn();
  const values = getData();

  console.log(values);

  for (let i = 0; i < values.length; i++) {
    let sentRange = sentSheet.getRange(`A${i + 2}:H${i + 2}`);

    const today = new Date();
    const localeDateTime = today.toLocaleString();
    const comma = localeDateTime.indexOf(',');
    const localeDate = localeDateTime.slice(0, comma);

    if (values[i]) values[i].push(localeDate);

    sentRange.setValues([values[i]]);
    console.log(values[i])
  }
}