/**
 * Sends an email to client if IARC checks the evaluation check box.
 */

// https://support.google.com/docs/forum/AAAABuH1jm0hR40qh02UWE/?hl=en&gpf=%23!topic%2Fdocs%2FhR40qh02UWE


function evaluation(event) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Active");
    var startRow = 2; // Start at second row because the first row contains the data labels
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var dataRange = sheet.getRange(startRow, 1, lastRow, lastCol);
    var date = Utilities.formatDate(new Date(), "GMT+10", "yyyy/MM/dd");
    // Fetch values for each row in the Range.
    var data = dataRange.getValues();
    var htmlBody = HtmlService.createTemplateFromFile('eval-email');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s = event.source.getActiveSheet();
    var r = event.source.getActiveRange();


    for (var i = 0; i < data.length - 1; ++i) {
        var row = data[i];
        var emailAddress = row[5]; // sixth column
        var name = row[4]; // fifth column
        htmlBody.name = name;
        htmlBody.email = emailAddress;
        const htmlForEmail = htmlBody.evaluate().getContent();

        var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
        Logger.log("Remaining email quota: " + emailQuotaRemaining);



        var cell = row[19]; // checkbox
        var sentCol = 21;


        if (cell == true) {

            MailApp.sendEmail(emailAddress, "Hi " + name + ",\n Can you please evaluate the Visa Assist program?", "Please open your email with a client that supports HTML", {
                    htmlBody: htmlForEmail,
                    bcc: "addyouremail",
                    cc: "addanotheremail"
                }

            );
            sheet.getRange(startRow + i, sentCol).setValue(date);
            var row = r.getRow();
            var numColumns = s.getMaxColumns();
            var targetSheet = ss.getSheetByName("Completed");
            var target = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
            s.getRange(row, 1, 1, numColumns).moveTo(target);
            s.deleteRow(row);

        }

    }
}