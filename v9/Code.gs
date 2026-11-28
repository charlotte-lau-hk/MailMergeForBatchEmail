// @ts-nocheck
/*******************************************************
 * Mail Merge for Batch Email (v9.3.0)
 * Maintained by Charlotte Lau
 * Last Update: 2026-11-28 (error handling for the whole run, lock, full-run log, quota handling, plain-text body)
 * GitHub: https://github.com/charlotte-lau-hk/MailMergeForBatchEmail
 * Features: Rich-text, inline images, QR codes, attachments, multiple templates,
 *   auto-rerun, AI template suggestions, dynamic email preview
 *******************************************************/
// Mail limit: 
//     100/day for consumer Google account
//    1500/day for an Google Apps for Education account
// source: https://developers.google.com/apps-script/guides/services/quotas

function onOpen(e) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // [2026-11-28] Menu first, so Initialize Sheets is available even if sheets are missing
  SpreadsheetApp.getUi().createMenu("Mail Merge for Batch Email")
      .addItem('📑 Initialize Sheets', 'initializeSheets_')
      .addItem('📁 Pick Folder for Attachment', 'launchFolderPicker_')
      .addSeparator()
      .addItem('🔍 Preview Merged Emails', 'previewMergedEmails_')
      .addItem('📧 Send Emails Now!', 'preSendEmails_')
      .addSeparator()
      .addItem('📬 Check Mail Quota', 'checkMailQuota_')
      .addToUi();

  updateDashboard_(ss, {"status": "Initializing..."});
  let dashboard = ss.getSheetByName("Dashboard");
  if (dashboard) dashboard.activate();

  const documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty("SpreadsheetId", ss.getId()); // store doc id for time-driven trigger
  let triggerRecord = documentProperties.getProperty("triggerId");
  if (triggerRecord) {
    updateDashboard_(ss, {"status": "Trigger exists.", "trigger": triggerRecord});
  } else {
    updateDashboard_(ss, {"status": "Clean start.", "trigger": "---"});
  }

  protectSheets_(ss, ["Dashboard", "Usage"]);
  updateDashboard_(ss, {"quota": "(Run Check Mail Quota to update)"})
}

function onInstall(e) {
  onOpen(e);
}

// Validate settings before sending or previewing emails
function validateSettings_() {
  // first run, do checking first
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  updateDashboard_(ss, {"status": "Validating Settings"});
  SpreadsheetApp.flush(); // show now

  const documentProperties = PropertiesService.getDocumentProperties();
  var err = 0;
  
  // Check existence of necessary worksheets
  if (!ss.getSheetByName("Dashboard")) err++;
  if (!ss.getSheetByName("Datasheet")) err++;
  if (!ss.getSheetByName("Settings")) err++;
  if (!ss.getSheetByName("Template1")) err++;
  if (err) {
    var response = ui.alert("Fatal Error",
                          "Necessary worksheets are missing! Cannot proceed!"
                          +'\nRun "Initialize Sheets" in the menu to rebuild.',
                          ui.ButtonSet.OK);
    return false;
  }

  var settingsSheet = ss.getSheetByName("Settings");
  var ccList = settingsSheet.getRange("B2").getValue();
  var bccList = settingsSheet.getRange("B3").getValue();
  // Validate Cc and Bcc lists
  if (ccList.length>0) {
    let ccArray = ccList.split(",").map(s => s.trim());
    let err = 0;
    ccArray.forEach((addr) => {
      if (!isValidEmail_(addr)) err++;
    })
    if (err) {
      Logger.log("Invalid email in Cc list! Check it please.");
      var response = ui.alert("Email Address Error in Cc List",
                    "There is an error in the Cc list."
                    +"\nPlease check email addresses in the list.",
                    ui.ButtonSet.OK);
      return false;              
    }
  }
  if (bccList.length>0) {
    let bccArray = bccList.split(",").map(s => s.trim());
    let err = 0;
    bccArray.forEach((addr) => {
      if (!isValidEmail_(addr)) err++;
    })
    if (err) {
      Logger.log("Invalid email in Bcc list! Check it please.");
      var response = ui.alert("Email Address Error in Bcc List",
                    "There is an error in the Bcc list."
                    +"\nPlease check email addresses in the list.",
                    ui.ButtonSet.OK);
      return false;
    }
  }

  // Check folder existence before sending
  var templateSheet = ss.getSheetByName("Settings");
  var folderId = templateSheet.getRange("B4").getValue();
  if (folderId.length==0) folderId = null;

  // Check folder existence
  if (folderId != null) {
    try {
      DriveApp.getFolderById(folderId);
    } catch(e) {
      var response = ui.alert("Folder Error",
                          "The folder does not exist or you don't have permission to access it!"
                          +"\nClear or correct the Folder ID!",
                          ui.ButtonSet.OK);
      return false;
    }
  }

  documentProperties.setProperties({
    "ccList": ccList,
    "bccList": bccList,
    "folderId": folderId || "",
    "sendAs":  settingsSheet.getRange("B5").getValue(),
    "replyTo": settingsSheet.getRange("B6").getValue(),
    "noReply": settingsSheet.getRange("B7").getValue()
  });
  // clear status on dashboard
  updateDashboard_(ss, {"status": "Validation done."});

  return true;
}

// [2026-11-27] Single rule for "To Send": TRUE (checkbox) or "Y"/"Yes" (any case)
function isMarkedToSend_(value) {
  if (typeof value == "boolean") return value;
  if (typeof value == "string") return ["y", "yes"].indexOf(value.trim().toLowerCase()) != -1;
  return false;
}

// Email address validation
function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// Protect Sheets
function protectSheets_(ss, sheets) {
  sheets.forEach(sname => {
    let sheet = ss.getSheetByName(sname);
    if (sheet) {
      sheet.protect().setDescription("Read-only. Take care to edit.").setWarningOnly(true);
    }
  });
}

// Update Dashboard
// - ss: SpreadSheet object
// - quota: Mail Quota Remaining
// - lastRun: Last Run Time
// - status: Current Status
// - trigger: Trigger Status
function updateDashboard_(ss, opts = null) {
  let {quota, lastRun, status, trigger} = opts;
  let sheet = ss.getSheetByName("Dashboard");
  if (sheet) {
    if (quota) sheet.getRange("B5").setValue(quota);
    if (lastRun) sheet.getRange("B6").setValue(lastRun);
    if (status) sheet.getRange("B7").setValue(status);
    if (trigger) sheet.getRange("B8").setValue("{"+trigger+"}");
    sheet.autoResizeColumn(2);
    sheet.setColumnWidth(2, sheet.getColumnWidth(2) * 1.1);
  }
  SpreadsheetApp.flush(); // show now
}

// List rows to send
function listRowsToSend_() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName("Datasheet");
  let numRows = dataSheet.getMaxRows() - 1;
  let rangeData = dataSheet.getRange(2, 1, numRows, 2).getValues();
  // check rows to send and not yet send
  let toSendList = [];
  for (let i=0; i<numRows; i++) {
    let toSend = rangeData[i][0];
    let done = rangeData[i][1];
    // only those marked to send but not yet sent
    if (!isMarkedToSend_(toSend)) {
      continue;
    }
    if ((typeof done == "string") && (done == "Done")) {
      continue;
    }
    toSendList.push(i+1);
  };

  return toSendList;
}

// Prepare to send emails
function preSendEmails_() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!validateSettings_()) {
    Logger.log("Settings invalid. Stop.");
    updateDashboard_(ss, {"status": "Settings invalid. Stop."});
    return;
  }

  // Adding confirmation box
  // Ref: https://developers.google.com/apps-script/reference/base/ui#alerttitle-prompt-buttons
  var response = ui.alert('Confirm Sending Emails',
                          "You are about to send batch emails. Proceed?\nNote: Any time-based triggers will be cleared.",
                          ui.ButtonSet.OK_CANCEL);
  if (response == ui.Button.CANCEL) {
    return;
  }

  // check remaining quota before proceed
  let emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  let toSend = listRowsToSend_();
  if (toSend.length > emailQuotaRemaining) {
    var response = ui.alert('Mail Quota Not Enough',
                            "You have " + toSend.length + " mails to send but "
                            + "your remaining quota is only " + emailQuotaRemaining
                            + ". Proceed any way?",
                            ui.ButtonSet.OK_CANCEL);
    if (response == ui.Button.CANCEL) {
      return;
    }
  }

  sendEmails_({ fresh: true });
}

// handler for File Picker functions
function setAttachmentFolderId_(folderId) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings").getRange("B4").setValue(folderId);
}

// to call Folder Picker
function launchFolderPicker_() {
  FP_showFolderPicker(setAttachmentFolderId_);
}

// Check and show mail quota
function checkMailQuota_() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  updateDashboard_(ss, {"quota": emailQuotaRemaining});
  var ui = SpreadsheetApp.getUi();
  ui.alert("Remaining Daily Quota",
           "Your remaining mail quota is: " + emailQuotaRemaining,
           ui.ButtonSet.OK);
}

// [2026-11-28] v9.3.0: One sending run may span several executions (rerun trigger).
// A run started from the menu passes {fresh: true}; the rerun trigger passes its event object.
function sendEmails_(e) {
  const fresh = (e && e.fresh === true);
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss == null) {
    // get id from properties in case of time-driven trigger
    let id = PropertiesService.getDocumentProperties().getProperty("SpreadsheetId");
    ss = SpreadsheetApp.openById(id);
  }

  // Refuse to start while another run is sending
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    showAlert_("Sending in Progress", "Another sending run is in progress. Please wait for it to finish.");
    return;
  }

  RUN_ENDED_ = false;
  if (fresh) clearRunLog_();
  try {
    sendBatch_(ss, fresh);
  } catch (err) {
    // Any unexpected error: record it, stop, and still send the log
    log_("ERROR: " + err.message);
    console.error(err);
    try { clearTriggers_(); } catch (e2) { log_("Clearing triggers failed: " + e2.message); }
    try { updateDashboard_(ss, {"status": "Stopped with error: " + err.message + " (see log email)"}); } catch (e4) { console.error(e4); }
    try { endRun_(); } catch (e3) { console.error(e3); }
  } finally {
    if (!RUN_ENDED_) appendRunLog_(); // keep this part of the log for the next execution
    lock.releaseLock();
  }
}

function sendBatch_(ss, fresh) {
  let status = fresh ? "Start running" : "Continue running (from trigger)";
  let now = new Date();
  let timeRun = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm:ss");
  PropertiesService.getDocumentProperties().setProperty("runTime", now.getTime().toString());
  updateDashboard_(ss, {"status": status, "lastRun": timeRun, "trigger": "---"});
  log_(status);
  clearTriggers_();

  var dataSheet = ss.getSheetByName("Datasheet");
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows() - 1, dataSheet.getMaxColumns());

  // Get settings configurations
  const documentProperties = PropertiesService.getDocumentProperties();
  let settings = documentProperties.getProperties(); // at this stage, property must exist
  let {ccList, bccList, folderId, sendAs, replyTo, noReply} = settings;
  if (!folderId) folderId = null;

  // converter for showdown [added in v5.00]
  // Ref: https://github.com/anshulguleria/markdown-to-doc
  // Options: https://github.com/showdownjs/showdown
  var converter = new showdown.Converter({tables:true,underline:true});

  // Check folder existence
  var folder = null;
  if (folderId != null) {
    folder = DriveApp.getFolderById(folderId);
  } else {
    log_("Attachment folder is NULL. Emails with attachments will not be sent!");
  }

  // Each email uses one quota unit per recipient (To + Cc + Bcc); keep 1 for the log email
  let emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  const countAddresses = (list) => (list ? list.split(",").filter(s => s.trim() != "").length : 0);
  const unitsPerEmail = 1 + countAddresses(ccList) + countAddresses(bccList);

  // Create one JavaScript object per row of data.
  var objects = getRowsData_(dataSheet, dataRange);

  for (var i=0; i<objects.length; ++i) {
    // Get a row object
    var rowData = objects[i];
    // [2026-11-27] Actual sheet row of this record (empty rows are skipped by getRowsData_)
    var rowNum = dataRange.getRow() + rowData.__rowIndex;

    if (!isMarkedToSend_(rowData.toSend)) {
      continue;
    }

    var done = rowData.done;
    if ((typeof done == "string") && (done == "Done")) {
      continue;
    }

    if (!isValidEmail_(rowData.emailAddressToSend)) {
      log_("Row " + rowNum + ": invalid email: " + rowData.emailAddressToSend);
      dataSheet.getRange(rowNum, 2).setValue("Error: Invalid email");
      continue;
    }

    var template = rowData.template;
    var templateSheet = ss.getSheetByName(template);
    if (!templateSheet) {
      log_("Row " + rowNum + ": skipped (template not found: " + template + ")");
      dataSheet.getRange(rowNum, 2).setValue("Error and Skipped (template not found)");
      continue;
    }

    var subjTemplate = templateSheet.getRange("B1").getValue();
    var bodyTemplate = templateSheet.getRange("B2").getValue();

    // Generate a personalized email.
    // Given a template string, replace markers (for instance ${"First Name"}) with
    // the corresponding value in a row object (for instance rowData.firstName).
    var emailSubject = fillInTemplateFromObject_(subjTemplate, rowData);
    var emailBody = fillInTemplateFromObject_(bodyTemplate, rowData);

    // Adding HTML (rich-text) [added in v5.00]
    var htmlBody = converter.makeHtml(emailBody);
    // [2026-11-28] Readable plain-text version for mail apps that do not show HTML
    var plainBody = htmlToPlainText_(htmlBody);

    // Get a list of objects of attachments (also checking existence)
    var attachments = new Array;
    var fileList = [];
    var attachmentError = 0;
    if ((rowData.attachmentList !== undefined) && (rowData.attachmentList != "")) {
      fileList = rowData.attachmentList.split(',').map(s => s.trim());
    }
    // [2026-11-27] Folder for this row only (subfolder no longer carries over to later rows)
    var rowFolder = folder;
    if ((rowFolder != null) && (rowData.subfolder !== undefined)) {
      var subfolders = rowFolder.getFoldersByName(rowData.subfolder);
      if (subfolders.hasNext()) {
        rowFolder = subfolders.next();
      }
    }
    // If no folderID given BUT attachment list is not empty, skip this email!
    if ((rowFolder == null) && (fileList.length >0)) {
      attachmentError++;
    } else {
      for (var j=0; j<fileList.length; j++) {
        var files = rowFolder.getFilesByName(fileList[j]);
        if (files.hasNext()) {
          // note: assume only one file with the same name in the folder
          attachments.push(files.next());
        } else {
          // file not found, assume error and skip this record
          attachmentError++;
        }
      }
    }

    if (attachmentError) {
      log_("Row " + rowNum + ": skipped (attachment error): " + rowData.emailAddressToSend);
      dataSheet.getRange(rowNum, 2).setValue("Error and Skipped (attachment error)");
      continue;
    }

    // [2022-04-21] Prepare inline images
    // [2026-11-27] An image that cannot be loaded skips this row instead of stopping the run
    var inlineImages = new Object();
    var imageError = "";
    for (const [key, value] of Object.entries(rowData)) {
      if (key.substring(0,7)=="imglink" && value != "") {
        try {
          inlineImages[key] = UrlFetchApp.fetch(value).getBlob().setName(key);
        } catch (e) {
          imageError = key + " cannot be fetched";
          break;
        }
      } else if (key.substring(0,7)=="imgfile" && value != "") {
        if (rowFolder == null) {
          imageError = key + " needs a Folder ID";
          break;
        }
        var files = rowFolder.getFilesByName(value);
        if (files.hasNext()) {
          inlineImages[key] = files.next().getBlob().setName(key);
        } else {
          imageError = key + " file not found";
          break;
        }
      } else if (key.substring(0,6)=="qrdata" && value != "") {
        // [2026-10-09] use qrcode javascript library
        try {
          inlineImages[key] = makeQrBlob_(value, key);
        } catch (e) {
          log_("Row " + rowNum + ": QR code generation failed for " + value + ": " + e.message);
        }
      }
    }

    if (imageError) {
      log_("Row " + rowNum + ": skipped (image error: " + imageError + "): " + rowData.emailAddressToSend);
      dataSheet.getRange(rowNum, 2).setValue("Error and Skipped (image error: " + imageError + ")");
      continue;
    }

    // [2026-11-28] Stop cleanly before the daily quota runs out
    if (emailQuotaRemaining - unitsPerEmail < 1) {
      log_("Daily email quota reached. Stopped before row " + rowNum + ". Run again tomorrow to continue.");
      clearTriggers_();
      updateDashboard_(ss, {"status": "Stopped: daily email quota reached. Run again tomorrow to continue.",
                            "quota": emailQuotaRemaining});
      endRun_();
      return;
    }

    // Send out now
    var opts = new Object();
    if (attachments.length > 0) opts['attachments'] = attachments;
    if (ccList != "") opts['cc'] = ccList;
    if (bccList != "") opts['bcc'] = bccList;
    if (sendAs != "") opts['name'] = sendAs;
    if (replyTo != "") opts['replyTo'] = replyTo;
    if (noReply == true || String(noReply).toLowerCase() == "true") opts['noReply'] = true;
    if (Object.keys(inlineImages).length>0) opts['inlineImages'] = inlineImages;
    opts['htmlBody'] = htmlBody;

    // Ref: https://developers.google.com/apps-script/reference/mail/mail-app#sendEmail(String,String,String,Object)
    MailApp.sendEmail(rowData.emailAddressToSend, emailSubject, plainBody, opts);
    emailQuotaRemaining -= unitsPerEmail;
    log_("Row " + rowNum + ": sent to " + rowData.emailAddressToSend + " (" + attachments.length + " attachment(s))");

    // Mark "Done" in Datasheet
    dataSheet.getRange(rowNum, 2).setValue("Done");
    SpreadsheetApp.flush();

    // Checking time-out
    if (isTimedOut_()) {
      // prepare trigger for re-run after 30 seconds
      log_("Execution time limit reached. Continuing in 30 seconds.");
      updateDashboard_(ss, {"quota": emailQuotaRemaining});
      setRerunTrigger_();
      return;
    };
  }

  // Up to this point, all records processed
  log_("All rows processed. Email quota remaining: " + (emailQuotaRemaining - 1));
  updateDashboard_(ss, {"status": "Completed", "quota": emailQuotaRemaining - 1});
  clearTriggers_();
  endRun_();
}

// ---------- Run log (kept across executions of one run) [2026-11-28] ----------
var RUN_LOG_ = [];        // lines from this execution
var RUN_ENDED_ = false;   // true once the log email has been sent and the stored log cleared
const RUN_LOG_CHUNK_ = 3000;    // characters per property (max 9 KB per value, up to 3 bytes per character)
const RUN_LOG_MAX_ = 90000;     // keep the stored log well under the 500 KB property limit

function log_(message) {
  Logger.log(message);
  RUN_LOG_.push(Utilities.formatDate(new Date(), "GMT+8", "HH:mm:ss") + "  " + message);
}

function loadRunLog_() {
  const props = PropertiesService.getDocumentProperties();
  const count = parseInt(props.getProperty("runLogChunks") || "0", 10);
  let text = "";
  for (let i = 0; i < count; i++) text += props.getProperty("runLog_" + i) || "";
  return text;
}

function clearRunLog_() {
  const props = PropertiesService.getDocumentProperties();
  const count = parseInt(props.getProperty("runLogChunks") || "0", 10);
  for (let i = 0; i < count; i++) props.deleteProperty("runLog_" + i);
  props.deleteProperty("runLogChunks");
}

function appendRunLog_() {
  if (RUN_LOG_.length == 0) return;
  let text = loadRunLog_() + RUN_LOG_.join("\n") + "\n";
  RUN_LOG_ = [];
  if (text.length > RUN_LOG_MAX_) {
    text = "(earlier lines omitted)\n" + text.substring(text.length - RUN_LOG_MAX_);
  }
  clearRunLog_();
  const props = PropertiesService.getDocumentProperties();
  const chunks = {};
  let n = 0;
  for (let i = 0; i < text.length; i += RUN_LOG_CHUNK_) {
    chunks["runLog_" + n] = text.substring(i, i + RUN_LOG_CHUNK_);
    n++;
  }
  chunks["runLogChunks"] = String(n);
  props.setProperties(chunks);
}

// Send the whole run's log to the account running the script, then clear it
function endRun_() {
  appendRunLog_();
  const body = loadRunLog_();
  const recipient = Session.getEffectiveUser().getEmail();
  const timeStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
  const subject = "** Mail Merge Log ** (" + timeStr + ")";
  const html = body.split("\n").join("<br>");
  MailApp.sendEmail(recipient, subject, body, { htmlBody: html, noReply: true });
  clearRunLog_();
  RUN_ENDED_ = true;
}

// Alert only when a user interface exists (not in trigger runs)
function showAlert_(title, message) {
  Logger.log(title + ": " + message);
  try {
    let ui = SpreadsheetApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (e) {
    // no UI available
  }
}

// [2026-11-28] Convert the HTML body to readable plain text
function htmlToPlainText_(html) {
  return String(html)
    .replace(/<img[^>]*>/gi, "[image]")
    .replace(/<a [^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi,
             (m, url, text) => (text == url ? url : text + " (" + url + ")"))
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<li[^>]*>/gi, "- ")
    .replace(/<\/li>/gi, "\n")
    .replace(/<\/(p|div|h[1-6]|tr|ul|ol|table|blockquote)>/gi, "\n\n")
    .replace(/<\/t[dh]>/gi, "\t")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/&#(\d+);/g, (m, n) => String.fromCharCode(parseInt(n, 10)))
    .replace(/&amp;/g, "&")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

// [2025-03-25] New feature for v9
// Initialize sheets 
function initializeSheets_() {
  let currentUser = Session.getActiveUser().getEmail();
  let ui = SpreadsheetApp.getUi();
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  // To setup dashboard with default content
  const defaults = {
    "Dashboard": {
      data: [
          [ "Total Emails", '=countif(Datasheet!A:A,"Yes")+countif(Datasheet!A:A,"Y")+countif(Datasheet!A:A,TRUE)' ],
          [ "Emails Done", '=countif(Datasheet!B:B,"Done")' ],
          [ "Emails Remaining", "=B1-B2" ],
          [ "Progress", "=TO_PERCENT(B2/B1)" ],
          [ "Remaining Quota", "(Run Check Mail Quota in menu to update)" ],
          [ "Last Run Time", "---" ],
          [ "Status", "Idle (Job not started yet)" ],
          [ "Trigger ID", "---" ]
      ],
      format: {
        "border": true,
        "firstColumnBold": true
      }
    },
    "Template1": {
      data: [
        [ "Subject",
          'Result for ${"UserName"} (${"UserID"})'],
        [ "Body (Markdown)",
          'Dear ${"UserName"} (${"UserID"}),\n\n' +
          'Your **elective subject** results:\n' +
          '- X1: ${"X1"}\n' +
          '- X2: ${"X2"}\n' +
          '- X3: ${"X3"}\n\n' +
          '**__imgfile1__**  \n' +
          '${"imgfile1"}\n\n' +
          '**__imglink1__**  \n' +
          '${"imglink1"}\n\n' +
          '**__qrdata1__**  \n' +
          '${"qrdata1"}\n\n' +
          '<br>\n' +
          'Administrator'
        ]
      ],
      format: {
        "border": true,
        "firstColumnBold": true
      }
    },
    "Datasheet": {
      data: [
        [ "To Send", "Done?", "Email Address to Send", "Template", "Attachment List", "Subfolder", "UserID", "UserName", "X1", "X2", "X3", "imgfile1", "qrdata1", "imglink1"],
        [ "Yes", "", "user@example.com", "Template1", "", "", "1001", "Alice Chan", "Physics", "Chemistry", "Biology", "sunset1.jpg", "https://github.com/charlotte-lau-hk", "https://raw.githubusercontent.com/charlotte-lau-hk/MailMergeForBatchEmail/main/test-data/sunset2.jpg"],
        [ "", "", "user@example.com", "Template1", "ABC.pdf,PQR.pdf", "", "1002", "Bob Wong", "Geography", "History", "Economics", "", "", ""],
        [ "", "", "user@example.com", "Template1", "XYZ.pdf", "subfolder1", "1003", "Cathy Lee", "Literature", "Visual Arts", "Music", "", "", ""]
      ],
      format: {
        "firstRowBold": true
      }
    },
    "Settings": {
      data: [
        [ "Settings", "Values", "Remarks" ],
        [ "Cc", "", "comma-separated list of email addresses" ],
        [ "Bcc", "", "comma-separated list of email addresses" ],
        [ "Folder ID", "1TLhS8KsSIBbmR9rTYQ-dCeB-hFTQm9da", "ID of template folder, required for attachment" ],
        [ "Send as", "", "The name to show as sender" ],
        [ "Reply to", "", "Email address to receive user reply" ],
        [ "No reply", false, "No reply: TRUE or FALSE" ]
      ],
      format: {
        "border": true,
        "firstRowBold": true,
        "firstColumnBold": true
      }
    },
    "Usage": {
      data: [
        [ "Mail Merge for Batch Email (v9.3)" ],
        [ "Maintained by: Charlotte Lau" ],
        [ "GitHub: https://github.com/charlotte-lau-hk/MailMergeForBatchEmail" ],
        [ "Guide: https://www.syclau.hk/mail-merge-for-batch-email" ],
        [ "To use: Make a copy of this file" ],
        [ "** Reference **" ],
        [ "Markdown Cheatsheet: https://github.com/adam-p/markdown-here/wiki/Markdown-Cheatsheet" ],
        [ "Online Markdown Editor: https://jbt.github.io/markdown-editor/" ],
        [ "QR Code Generator: qrcode-generator by Kazuhiko Arase (https://github.com/kazuhikoarase/qrcode-generator)" ]
      ],
      format: {
        "firstRowBold": true
      }
    }
  }

  // To setup datasheet with sample content (3 rows)
  const setupSheet = (sheet, sheetName, data, format) => {
    Logger.log("Setting up: " + sheetName);
    let range = sheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);
    if (format["border"]) {
      range.setBorder(true, true, true, true, true, true);
    }
    if (format["firstRowBold"]) {
      sheet.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
    }
    if (format["firstColumnBold"]) {
      sheet.getRange(1, 1, data.length, 1).setFontWeight('bold');
    }
    if (sheetName == "Dashboard") {
      Logger.log("setHorizontalAlignment");
      sheet.getRange(1, 2, data.length, 1).setHorizontalAlignment("center");
    } else if (sheetName == "Datasheet") {
      sheet.getRange("A:F").setBackground("#fffdd0");
    }
    range.setVerticalAlignment('top');
    sheet.autoResizeRows(1, data.length);
    for (let i = 1; i <= data[0].length; i++) {
      sheet.autoResizeColumn(i);
      sheet.setColumnWidth(i, sheet.getColumnWidth(i) * 1.1);
    }
  }

  Object.keys(defaults).forEach((sheetName) => {
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      // If exist, ask if to replace content
      let response = ui.alert("⚠️ Replace " + sheetName + "?",
                              "The " + sheetName + " sheet exists. Replace it with default content?\n" +
                              "Note: All current data will be lost.",
                              ui.ButtonSet.YES_NO);
      if (response == ui.Button.NO) {
        return;
      }
    } else {
      // Create if not exist
      Logger.log("Create sheet: " + sheetName)
      sheet = ss.insertSheet(sheetName, 0);
    }
    sheet.clear();
    setupSheet(sheet, sheetName, defaults[sheetName].data, defaults[sheetName].format);
  });

  // final setup
  let emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  updateDashboard_(ss, {"quota": emailQuotaRemaining});
  protectSheets_(ss, ["Dashboard", "Usage"]);
  SpreadsheetApp.flush();
}

// [2025-04-04] New timeout handleing mechanism
// [2025-04-09] Update to use BigInt
function isTimedOut_() {
  let now = BigInt(Date.now());
  let lastRun = BigInt(PropertiesService.getDocumentProperties().getProperty("runTime") || 0);
  if (now - lastRun >= 330 * 1000) {
    Logger.log("Timeout, prepare for re-run");
    return true;
  } else {
    return false;
  }
}

// Set re-run trigger
function setRerunTrigger_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Clear existing triggers to avoid overlap
  clearTriggers_();
  // Create one-time trigger for re-run
  let triggerId = ScriptApp.newTrigger("sendEmails_")
                      .timeBased()
                      .after(30000)
                      .create()
                      .getUniqueId()
                      .toString();
  updateDashboard_(ss, {"status": "Trigger set for re-run (wait 30s)", "trigger": triggerId});
  PropertiesService.getDocumentProperties().setProperty("triggerId", triggerId);
  Logger.log("Trigger set: "+ triggerId);
}

// Clear triggers
function clearTriggers_() {
  let triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    Logger.log("Deleting trigger: " + triggers[i].getUniqueId().toString());
    ScriptApp.deleteTrigger(triggers[i]);
    Utilities.sleep(1000);
  }
  PropertiesService.getDocumentProperties().deleteProperty("triggerId");
}

// make QR blob
function makeQrBlob_(text, key, cellPx = 13) {
  const qr = qrcode(0, 'M');
  qr.addData(text);
  qr.make();
  const url = qr.createDataURL(cellPx, cellPx * 2); // margin is in pixels
  return Utilities.newBlob(Utilities.base64Decode(url.split(',')[1]), 'image/gif', key + '.gif');
}
