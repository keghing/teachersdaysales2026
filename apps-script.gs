// ============================================================
//  HARI GURU SHOP — Google Apps Script Backend
//  Paste this entire file into your Google Apps Script editor
// ============================================================

// 🔧 STEP 1: Paste your Google Sheet ID here
const SHEET_ID = "19ddNNyKGbXEU3NZEa19G1tkq_akSvpJQJdd7maPk5gI";

// Sheet tab names — don't change unless you rename your tabs
const ORDERS_TAB = "Orders";
const MAILS_TAB  = "Secret Mails";
const ORDERS_HEADERS = [
  "Order #", "Date & Time", "Buyer Name", "Class",
  "Phone", "Items Ordered", "Total (RM)",
  "Payment Method", "Payment Screenshot", "Remark", "Status"
];

// ============================================================
//  RUN THIS FIRST — authorizes both Sheets and Drive
//  In the editor: select authorizeApp from the dropdown → click ▶ Run
//  You MUST approve every permission popup that appears.
// ============================================================
function authorizeApp() {
  DriveApp.getRootFolder();
  SpreadsheetApp.openById(SHEET_ID);
  Logger.log("Authorization complete!");
}

// ============================================================
//  RUN THIS to verify Drive upload works before going live
//  Select testDriveUpload → click ▶ Run → check Execution Log
// ============================================================
function testDriveUpload() {
  try {
    const testBlob = Utilities.newBlob("test", "text/plain", "drive-test.txt");
    const file = DriveApp.createFile(testBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Logger.log("✅ Drive upload works! File URL: " + file.getUrl());
    file.setTrashed(true); // clean up test file
  } catch (err) {
    Logger.log("❌ Drive upload FAILED: " + err.message);
  }
}

function ensureOrdersHeader(sheet) {
  sheet = sheet || SpreadsheetApp.openById(SHEET_ID).getSheetByName(ORDERS_TAB);
  if (!sheet) {
    throw new Error(`Sheet tab "${ORDERS_TAB}" was not found.`);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(ORDERS_HEADERS);
  } else {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers.indexOf("Remark") === -1) {
      const statusCol = headers.indexOf("Status") + 1;
      if (statusCol > 0) {
        sheet.insertColumnBefore(statusCol);
        sheet.getRange(1, statusCol).setValue("Remark");
      } else {
        sheet.insertColumnAfter(sheet.getLastColumn());
        sheet.getRange(1, sheet.getLastColumn()).setValue("Remark");
      }
    }
  }

  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold")
                                             .setBackground("#F06040")
                                             .setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);
}

// ============================================================
//  doPost — receives new orders from the website
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.openById(SHEET_ID);

    if (data.type === "order") {
      const ordersSheet = ss.getSheetByName(ORDERS_TAB);

      ensureOrdersHeader(ordersSheet);

      // Save screenshot to Google Drive if provided
      let screenshotUrl = "—";
      if (data.screenshot && data.screenshot.startsWith("data:image")) {
        try {
          const base64 = data.screenshot.replace(/^data:image\/\w+;base64,/, "");
          const ext    = data.screenshot.match(/data:image\/(\w+)/)[1] || "jpg";
          const blob   = Utilities.newBlob(Utilities.base64Decode(base64), `image/${ext}`, `payment-${data.id}.${ext}`);
          const file   = DriveApp.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          screenshotUrl = file.getUrl();
        } catch (imgErr) {
          screenshotUrl = "Upload failed: " + imgErr.message;
        }
      }

      const itemsStr = data.items.map(i => {
        const teacher = i.teacher ? ` (For: ${i.teacher})` : "";
        return `${i.name} x${i.qty}${teacher}`;
      }).join(" | ");

      ordersSheet.appendRow([
        data.id,
        data.ts,
        data.buyer,
        data.cls,
        data.phone || "—",
        itemsStr,
        data.total,
        data.pay === "tng" ? "Touch 'n Go" : "Cash",
        screenshotUrl,
        data.remark || "—",
        "Pending"
      ]);

      ordersSheet.autoResizeColumns(1, 11);

      // Write Secret Mails
      if (data.mails && data.mails.length > 0) {
        const mailSheet = ss.getSheetByName(MAILS_TAB);

        if (mailSheet.getLastRow() === 0) {
          mailSheet.appendRow([
            "Order #", "From (Buyer)", "Class",
            "To (Teacher)", "Message", "Anonymous?", "Delivered?"
          ]);
          mailSheet.getRange(1, 1, 1, 7).setFontWeight("bold")
                                         .setBackground("#E91E63")
                                         .setFontColor("#FFFFFF");
          mailSheet.setFrozenRows(1);
        }

        data.mails.forEach(m => {
          if (m.teacher) {
            mailSheet.appendRow([
              data.id, data.buyer, data.cls,
              m.teacher, m.message,
              m.anon ? "Yes (Anonymous)" : "No (Shown)",
              "No"
            ]);
          }
        });

        mailSheet.autoResizeColumns(1, 7);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, id: data.id }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
//  doGet — lets admin panel fetch all orders as JSON
// ============================================================
function doGet(e) {
  try {
    const ss          = SpreadsheetApp.openById(SHEET_ID);
    const ordersSheet = ss.getSheetByName(ORDERS_TAB);
    const mailSheet   = ss.getSheetByName(MAILS_TAB);
    ensureOrdersHeader(ordersSheet);

    const ordersData = ordersSheet.getDataRange().getValues();
    const mailsData  = mailSheet ? mailSheet.getDataRange().getValues() : [];

    const orders = ordersData.slice(1).map(r => ({
      id:         r[0], ts:     r[1], buyer:  r[2], cls:    r[3],
      phone:      r[4], items:  r[5], total:  r[6], pay:    r[7],
      screenshot: r[8], remark: r[9] || "", status: r[10]
    }));

    const mails = mailsData.slice(1).map(r => ({
      orderId: r[0], buyer: r[1], cls: r[2],
      teacher: r[3], message: r[4], anon: r[5], delivered: r[6]
    }));

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, orders, mails }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
