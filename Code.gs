function documentTrackerEdit(e) {

  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Sirf Document Tracker sheet
  if (sheet.getName() !== "Document Tracker") return;

  // Header ko ignore karo
  if (row === 1) return;

  // B = Invoice Amount
  // C = Currency
  // D = Entry Date & Time
  // E = Exchange Rate
  // F = INR Value
  // G = E-Way Bill Status

  // Sirf Invoice Amount ya Currency edit hone par kaam kare
  if (col !== 2 && col !== 3) return;

  const amount = sheet.getRange(row, 2).getValue();
  const currency = String(sheet.getRange(row, 3).getValue()).trim().toUpperCase();

  // Amount aur currency dono hone chahiye
  if (!amount || !currency) return;

  // Agar Entry Date & Time pehle se hai to usko change mat karo
  const entryCell = sheet.getRange(row, 4);

  if (!entryCell.getValue()) {
    entryCell.setValue(new Date());
    entryCell.setNumberFormat("dd/MM/yyyy HH:mm");
  }

  // Sirf INR, EUR aur USD allow kar rahe hain
  if (!["INR", "EUR", "USD"].includes(currency)) {
    sheet.getRange(row, 5).setValue("Invalid Currency");
    return;
  }

  let rate = 1;

  // INR ke liye rate 1
  if (currency === "INR") {

    rate = 1;

  } else {

    // Entry date nikalo
    const entryDate = sheet.getRange(row, 4).getValue();

    // Date ko YYYY-MM-DD me convert karo
    const dateString = Utilities.formatDate(
      new Date(entryDate),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    // Frankfurter historical rate API
    const url =
      "https://api.frankfurter.dev/v1/" +
      dateString +
      "?base=" +
      currency +
      "&symbols=INR";

    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();

    if (code !== 200) {
      sheet.getRange(row, 5).setValue("Rate Error");
      return;
    }

    const data = JSON.parse(response.getContentText());

    if (!data.rates || !data.rates.INR) {
      sheet.getRange(row, 5).setValue("Rate Not Found");
      return;
    }

    rate = data.rates.INR;
  }

  // Exchange Rate save karo
  sheet.getRange(row, 5).setValue(rate);

  // INR Value calculate karo
  const inrValue = Number(amount) * Number(rate);

  sheet.getRange(row, 6).setValue(inrValue);

  // E-Way Bill Status
  if (inrValue > 50000) {
    sheet.getRange(row, 7).setValue("REQUIRED");
  } else {
    sheet.getRange(row, 7).setValue("NOT REQUIRED");
  }
}


function sendDocumentReminders() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Document Tracker");

  if (!sheet) return;

  const data = sheet.getDataRange().getValues();

  // Column numbers
  const CI = 1;
  const AMOUNT = 2;
  const CURRENCY = 3;
  const ENTRY_DATE = 4;
  const EWAY_STATUS = 7;
  const INVOICE = 9;
  const PACKING_LIST = 10;
  const CUSTOMER_BILL = 11;
  const EWAY_BILL = 12;
  const AIRWAY_BILL = 13;
  const SENIOR_NAME = 14;
  const SENIOR_EMAIL = 15;
  const STATUS = 16;
  const LAST_REMINDER = 17;
  const NEXT_REMINDER = 18;
  const REMINDER_COUNT = 19;

  const today = new Date();

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const ciNumber = row[CI - 1];
    const amount = row[AMOUNT - 1];
    const currency = row[CURRENCY - 1];

    // Empty row ignore
    if (!ciNumber || !amount || !currency) continue;

    const seniorName = row[SENIOR_NAME - 1];
const seniorEmail = String(row[SENIOR_EMAIL - 1]).trim();

// Blank ya invalid email ko skip karo
if (
  !seniorEmail ||
  seniorEmail.toUpperCase() === "SENIOR EMAIL" ||
  !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(seniorEmail)
) {
  continue;
}
    // Documents
    const invoice = String(row[INVOICE - 1]).trim().toLowerCase();
    const packingList = String(row[PACKING_LIST - 1]).trim().toLowerCase();
    const customerBill = String(row[CUSTOMER_BILL - 1]).trim().toLowerCase();
    const ewayBill = String(row[EWAY_BILL - 1]).trim().toLowerCase();
    const airwayBill = String(row[AIRWAY_BILL - 1]).trim().toLowerCase();

    const ewayRequired =
      String(row[EWAY_STATUS - 1]).trim().toUpperCase() === "REQUIRED";

    let pendingDocuments = [];

    // Invoice
    if (invoice !== "received") {
      pendingDocuments.push("Invoice");
    }

    // Packing List
    if (packingList !== "received") {
      pendingDocuments.push("Packing List");
    }

    // Customer Bill
    if (customerBill !== "received") {
      pendingDocuments.push("Customer Bill");
    }

    // E-Way Bill only if REQUIRED
    if (ewayRequired && ewayBill !== "received") {
      pendingDocuments.push("E-Way Bill");
    }

    // Airway Bill
    if (airwayBill !== "received") {
      pendingDocuments.push("Airway Bill");
    }

    // Agar koi document pending nahi hai
    if (pendingDocuments.length === 0) {

      sheet.getRange(i + 1, STATUS).setValue("COMPLETED");
      sheet.getRange(i + 1, NEXT_REMINDER).clearContent();

      continue;
    }

    // Documents pending hain
    sheet.getRange(i + 1, STATUS).setValue("PENDING");

    // Next reminder cell
    const nextReminderCell = sheet.getRange(i + 1, NEXT_REMINDER);
    const nextReminder = row[NEXT_REMINDER - 1];

    // Agar first time hai to pehla reminder 2 din baad schedule karo
    if (!nextReminder) {

      const firstReminder = new Date();
      firstReminder.setDate(firstReminder.getDate() + 2);

      nextReminderCell.setValue(firstReminder);
      nextReminderCell.setNumberFormat("dd/MM/yyyy HH:mm");

      continue;
    }

    // Abhi reminder ka time nahi hua
    if (new Date(nextReminder) > today) {
      continue;
    }

    // Email subject
    const subject =
      "Reminder - Pending Documents - CI " + ciNumber;

    // Email body
    let body = "";

    body += "Dear " + (seniorName || "Sir/Madam") + ",\n\n";

    body +=
      "This is an automatic reminder regarding the pending documents for the following shipment:\n\n";

    body += "CI Number: " + ciNumber + "\n";
    body += "Invoice Amount: " + amount + " " + currency + "\n\n";

    body += "Pending Documents:\n";

    pendingDocuments.forEach(function(doc) {
      body += "• " + doc + "\n";
    });

    body +=
      "\nKindly provide/complete the pending documents.\n\n";

    body +=
      "This is an automatic reminder from Document Tracker.";

    // Send email
    MailApp.sendEmail({
      to: seniorEmail,
      subject: subject,
      body: body
    });

    // Last reminder date
    sheet.getRange(i + 1, LAST_REMINDER)
      .setValue(new Date());

    sheet.getRange(i + 1, LAST_REMINDER)
      .setNumberFormat("dd/MM/yyyy HH:mm");

    // Next reminder = 2 days later
    const nextDate = new Date();
    nextDate.setDate(nextDate.getDate() + 2);

    nextReminderCell.setValue(nextDate);
    nextReminderCell.setNumberFormat("dd/MM/yyyy HH:mm");

    // Reminder count
    const oldCount =
      Number(row[REMINDER_COUNT - 1]) || 0;

    sheet.getRange(i + 1, REMINDER_COUNT)
      .setValue(oldCount + 1);
  }
}
