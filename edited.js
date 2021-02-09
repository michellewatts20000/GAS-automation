/**
 * Sends an email to Referrer if a change has been made to Row R.
 */

function onEdit(e) { //"e" receives the event object
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    var range = e.range; //The range of cells edited
    var columnOfCellEdited = range.getColumn(); //Get column number
    var active_range = sheet.getActiveRange();
    var emailAddress = sheet.getRange(active_range.getRowIndex(), 3).getValue();
    var name = sheet.getRange(active_range.getRowIndex(), 2).getValue();
    var client = sheet.getRange(active_range.getRowIndex(), 5).getValue();
    var clientEmail = sheet.getRange(active_range.getRowIndex(), 6).getValue();
    var contents = sheet.getRange(active_range.getRowIndex(), 18).getValue();

    var todayDate = Utilities.formatDate(new Date(), "GMT+10", "yyyy/MM/dd");
    var htmlBody = HtmlService.createTemplateFromFile('edited');
    htmlBody.name = name;
    htmlBody.email = emailAddress;
    htmlBody.client = client;
    htmlBody.date = todayDate;
    htmlBody.contents = contents;
    const htmlForEmail = htmlBody.evaluate().getContent();

    if (columnOfCellEdited === 18) // Column 18 is Column R
    {

        MailApp.sendEmail(emailAddress, "Hi " + name + ",\n your client: " + client + "\n was updated", "Please open your email with a client that supports HTML", {
                htmlBody: htmlForEmail,
                cc: clientEmail,
                bcc: "visaassistrobot@unionsnsw.org.au"
            }

        );

    }
}