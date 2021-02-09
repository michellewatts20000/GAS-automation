/**
 * Sends an email to IARC if cell has stayed blank for more than 4 days - skips weekends.
 */

function blank() {
    var day = new Date();
    if (day.getDay() > 5 || day.getDay() == 0) {
        return;
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Active");
    var startRow = 2; // Start at second row because the first row contains the data labels
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var dataRange = sheet.getRange(startRow, 1, lastRow, lastCol);
    // Fetch values for each row in the Range.
    var data = dataRange.getValues();

    var todayDate = Utilities.formatDate(new Date(), "GMT+10", "yyyy/MM/dd");
    var htmlBody = HtmlService.createTemplateFromFile('blank-template');

    var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
    Logger.log("Remaining email quota: " + emailQuotaRemaining);


    for (var i = 0; i < data.length - 1; ++i) {
        var row = data[i];
        var cell = row[17]; // Check column to see if blank or not
        var date = new Date(row[18]); //convert js date into gs date
        var formattedDate = Utilities.formatDate(date, "GMT+10", "yyyy/MM/dd"); //format date using gs method

        var cellTime = row[18];

        var emailAddress = row[21];

        Logger.log(todayDate);
        Logger.log(row[18]);
        Logger.log(formattedDate);
        Logger.log(emailAddress);

        var name = row[1];
        htmlBody.name = row[4];
        htmlBody.union = row[3];
        htmlBody.referrer = row[1];
        htmlBody.referred = formattedDate;
        const htmlForEmail = htmlBody.evaluate().getContent();

        if (cellTime == "") {
            return;
        } else if (cell == "" && todayDate > formattedDate) {

            MailApp.sendEmail("addyouremail", "Hi IARC, Could you please let " + row[1] + " what is happening with their client", "Please open your email with a client that supports HTML", {
                    htmlBody: htmlForEmail,
                    bcc: "addyouremail"
                }

            );


        }

    }
}