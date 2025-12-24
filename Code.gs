function verifyCertificates() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("Form Responses 1"); 

  const data = sheet.getDataRange().getValues();

  // Expected course keyword 
  const EXPECTED_COURSE = "Database Management Systems (RDBMS) & Microsoft Fabric SQL";

  // Column indices (0-based)
  const ROLL_COL = 2;     // Column C
  const NAME_COL = 1;     // Column B (email address owner name usually)
  const EMAIL_COL = 3;    // Column D (personal gmail id)
  const URL_COL = 5;      // Column F
  const STATUS_COL = 6;   // Column G

  // Add status header if missing
  if (!data[0][STATUS_COL]) {
    sheet.getRange(1, STATUS_COL + 1).setValue("Verification Status");
  }

  for (let i = 1; i < data.length; i++) {

    // Skip already verified rows
    if (data[i][STATUS_COL]) continue;

    const studentEmail = data[i][EMAIL_COL];
    const verificationUrl = data[i][URL_COL];

    // Student name is usually NOT reliable from Form,
    // so we infer name from certificate itself
    let statusMessage = "";

    if (!verificationUrl) {
      sheet.getRange(i + 1, STATUS_COL + 1)
           .setValue("❌ NO VERIFICATION LINK PROVIDED");
      continue;
    }

    try {
      const response = UrlFetchApp.fetch(verificationUrl, {
        muteHttpExceptions: true,
        followRedirects: true
      });

      if (response.getResponseCode() !== 200) {
        statusMessage = "❌ INVALID OR UNREACHABLE LINK";
      } else {

        const pageText = response.getContentText().toLowerCase();

        const courseMatch = pageText.includes(EXPECTED_COURSE);

        if (courseMatch) {
          statusMessage = "✅ VERIFIED";
        } else {
          statusMessage = "❌ WRONG COURSE / INVALID CERTIFICATE";
        }
      }

    } catch (error) {
      statusMessage = "❌ ERROR FETCHING CERTIFICATE LINK";
    }

    // Write status to sheet
    sheet.getRange(i + 1, STATUS_COL + 1).setValue(statusMessage);

    // Send status email
    if (studentEmail) {
      MailApp.sendEmail({
        to: studentEmail,
        subject: "OS Certificate Verification Status",
        body:
          "Hello,\n\n" +
          "Your Operating Systems certificate verification result is:\n\n" +
          statusMessage + "\n\n" +
          "Regards,\n" +
          "Agneay B Nair"
      });
    }
  }
}
