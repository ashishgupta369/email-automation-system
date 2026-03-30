function sendEmailsBest() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let data = sheet.getDataRange().getValues();

  data.forEach((row, index) => {
    if (index === 0) return;

    let name = row[0];
    let email = row[1];
    let status = row[2];

    if (!email || status.toString().trim().toLowerCase() !== "pending") return;

    let subject = "Hello " + name;
    let message = "Hi " + name + ", this is an automated email.";

    GmailApp.sendEmail(email, subject, message);

    sheet.getRange(index + 1, 3).setValue("Done");
  });
}
