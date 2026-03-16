// =============================================
// IT Helpdesk Ticketing System
// =============================================

var SHEET_NAME    = "Tickets";
var HELPDESK_NAME = "IT Help Desk";
var YOUR_EMAIL    = "your.email@gmail.com"; 

var COL_EMAIL      = 2;  // B — Email address
var COL_NAME       = 3;  // C — Full Name
var COL_CATEGORY   = 5;  // E — Issue Category
var COL_PRIORITY   = 6;  // F — Priority
var COL_SUMMARY    = 7;  // G — Issue Summary
var COL_TICKET_ID  = 11; // K — Ticket ID
var COL_STATUS     = 12; // L — Status
var COL_RESOLUTION = 14; // N — Resolution Notes
var COL_DATE_RES   = 15; // O — Date Resolved

// ---- Fires when a form is submitted ----
function onFormSubmit(e) {
  try {
    var sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(SHEET_NAME);
    var row = e.range.getRow();

    Utilities.sleep(2000); // wait for Ticket ID formula

    var email    = sheet.getRange(row, COL_EMAIL).getValue();
    var name     = sheet.getRange(row, COL_NAME).getValue();
    var category = sheet.getRange(row, COL_CATEGORY).getValue();
    var priority = sheet.getRange(row, COL_PRIORITY).getValue();
    var summary  = sheet.getRange(row, COL_SUMMARY).getValue();
    var ticketId = sheet.getRange(row, COL_TICKET_ID).getValue();

    sheet.getRange(row, COL_STATUS).setValue("Open");

    if (!email || !email.includes("@")) {
      Logger.log("Invalid email at row " + row + ": " + email);
      return;
    }

    GmailApp.sendEmail(email,
      "[" + ticketId + "] IT Support Request Received",
      "Hi " + name + ",\n\n" +
      "Your IT support request has been received.\n\n" +
      "Ticket ID : " + ticketId + "\n" +
      "Category  : " + category + "\n" +
      "Priority  : " + priority + "\n" +
      "Summary   : " + summary + "\n\n" +
      "Our team will follow up shortly.\n" +
      "Keep this ticket ID for your reference.\n\n" +
      "— " + HELPDESK_NAME
    );

    Logger.log("Confirmation sent to: " + email);

  } catch(err) {
    Logger.log("onFormSubmit error: " + err.message);
  }
}

// ---- Fires when you manually edit any cell ----
function onEdit(e) {
  try {
    var sheet = e.source.getActiveSheet();
    if (sheet.getName() !== SHEET_NAME) return;

    var col = e.range.getColumn();
    var row = e.range.getRow();
    var val = e.range.getValue();

    if (col !== COL_STATUS) return;
    if (val !== "Resolved") return;
    if (row <= 1) return;

    var email    = sheet.getRange(row, COL_EMAIL).getValue();
    var name     = sheet.getRange(row, COL_NAME).getValue();
    var ticketId = sheet.getRange(row, COL_TICKET_ID).getValue();
    var summary  = sheet.getRange(row, COL_SUMMARY).getValue();
    var notes    = sheet.getRange(row, COL_RESOLUTION).getValue();

    sheet.getRange(row, COL_DATE_RES).setValue(new Date());

    if (!email || !email.includes("@")) {
      Logger.log("Invalid email for resolution at row " + row);
      return;
    }

    GmailApp.sendEmail(email,
      "[" + ticketId + "] Your IT Request Has Been Resolved",
      "Hi " + name + ",\n\n" +
      "Your support ticket has been resolved.\n\n" +
      "Ticket ID  : " + ticketId + "\n" +
      "Issue      : " + summary + "\n" +
      "Resolution : " + notes + "\n\n" +
      "If the issue returns, please submit a new ticket.\n\n" +
      "— " + HELPDESK_NAME
    );

    Logger.log("Resolution email sent to: " + email);

  } catch(err) {
    Logger.log("onEdit error: " + err.message);
  }
}