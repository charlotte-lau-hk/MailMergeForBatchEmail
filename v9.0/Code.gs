/*******************************************************
 * Mail Merge for Batch Email (v9.0)
 * Maintained by Charlotte Lau
 * Last Update: 2025-04-04
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
  updateDashboard_(ss, {"status": "Initializing..."});
  ss.getSheetByName("Dashboard").activate();

  SpreadsheetApp.getUi().createMenu("Mail Merge for Batch Email")
      .addItem('📑 Initialize Sheets', 'initializeSheets_')
      .addItem('📁 Pick Folder for Attachment', 'launchFolderPicker_')
      .addSeparator()
      .addItem('🔍 Preview Merged Emails', 'previewMergedEmails_')
      .addItem('📧 Send Emails Now!', 'preSendEmails_')
      .addSeparator()
      .addItem('📬 Check Mail Quota', 'checkMailQuota_')
      .addToUi();

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
    "folderId": folderId,
    "sendAs":  settingsSheet.getRange("B5").getValue(),
    "replyTo": settingsSheet.getRange("B6").getValue(),
    "noReply": settingsSheet.getRange("B7").getValue(),
    "qrApiUrl": settingsSheet.getRange("B8").getValue() || "https://qrcode.tec-it.com/API/QRCode?quietzone=2&dpi=150&&data=" // default QR API URL
  });
  // clear status on dashboard
  updateDashboard_(ss, {"status": "Validation done."});

  return true;
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
    let yesTxt = new Array("y", "yes");
    if ((typeof toSend == "undefined") || 
        ((typeof toSend == "string") && (yesTxt.indexOf(toSend.toLowerCase()) == -1)) ||
        ((typeof toSend == "boolean" && (!toSend)))
       ){
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

  sendEmails_();
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

function sendEmails_() {
  // Start sending email now
  let status = "Start running (from UI)";
  // start sending emails
  let now = new Date();
  let timeRun = Utilities.formatDate(now, "GMT+8", "yyyy-MM-dd HH:mm:ss");
  PropertiesService.getDocumentProperties().setProperty("runTime", now.getTime());
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss == null) {
    // get id from properties in case of time-driven trigger
    let id = PropertiesService.getDocumentProperties().getProperty("SpreadsheetId");
    ss = SpreadsheetApp.openById(id);
    status = "Start running (from Trigger)";
  }
  // Show status in Dashboard
  updateDashboard_(ss, {"status": status, "lastRun": timeRun, "trigger": "---"});
  clearTriggers_()


  var dataSheet = ss.getSheetByName("Datasheet");
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows() - 1, dataSheet.getMaxColumns());

  // Get settings configurations
  const documentProperties = PropertiesService.getDocumentProperties();
  let settings = documentProperties.getProperties(); // at this stage, property must exist
  let {ccList, bccList, folderId, sendAs, replyTo, noReply, qrApiUrl} = settings;
  if (folderId.length==0) folderId = null;

  // converter for showdown [added in v5.00]
  // Ref: https://github.com/anshulguleria/markdown-to-doc
  // Options: https://github.com/showdownjs/showdown
  var converter = new showdown.Converter({tables:true,underline:true});

  // Check folder existence
  var folder = null;
  if (folderId != null) {
    folder = DriveApp.getFolderById(folderId);
  } else {
    Logger.log("Attachment folder is NULL. Emails with attachments will not be sent!")
  }

  // Create one JavaScript object per row of data.
  var objects = getRowsData_(dataSheet, dataRange);

  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  var lastUpdate = 0;
  let emailQuotaRemaining;
  for (var i=0; i<objects.length; ++i) {
    // Get a row object
    var rowData = objects[i];

    var toSend = rowData.toSend;
    if ((typeof toSend == "undefined") ||
        ((typeof toSend == "string") && (toSend.toLowerCase() != "yes")) ||
        ((typeof toSend == "boolean") && (!toSend))) {
      Logger.log("Skipped (not to send): "+rowData.emailAddressToSend);
      continue;
    }
    
    var done = rowData.done;
    if ((typeof done == "string") && (done == "Done")) {
      Logger.log("Skipped (done): "+rowData.emailAddressToSend);
      continue;
    }

    if (!isValidEmail_(rowData.emailAddressToSend)) {
      Logger.log("Invalid email: " + rowData.emailAddressToSend);
      dataSheet.getRange(i+2, 2).setValue("Error: Invalid email");
      continue;
    }

    var template = rowData.template;
    var templateSheet = ss.getSheetByName(template);
    if (!templateSheet) {
      Logger.log("Skipped (template not found): "+template);
      dataSheet.getRange(i+2,2).setValue("Error and Skipped (template not found)");
      continue;
    }

    // Ready to send
    updateDashboard_(ss, {"status": "Sending email to " + rowData.emailAddressToSend});
    var subjTemplate = templateSheet.getRange("B1").getValue();
    var bodyTemplate = templateSheet.getRange("B2").getValue();

    // Generate a personalized email.
    // Given a template string, replace markers (for instance ${"First Name"}) with
    // the corresponding value in a row object (for instance rowData.firstName).
    var emailSubject = fillInTemplateFromObject_(subjTemplate, rowData);
    var emailBody = fillInTemplateFromObject_(bodyTemplate, rowData);
    
    // Adding HTML (rich-text) [added in v5.00]
    // Ref: https://github.com/anshulguleria/markdown-to-doc
    var htmlBody = converter.makeHtml(emailBody);

    // Get a list of objects of attachments (also checking existence)
    var attachments = new Array;
    var fileList = [];
    var attachmentError = 0;
    if ((rowData.attachmentList !== undefined) && (rowData.attachmentList != "")) {
      fileList = rowData.attachmentList.split(',').map(s => s.trim());
    }
    // If no folderID given BUT attachment list is not empty, skip this email!
    if ((folder == null) && (fileList.length >0)) {
      attachmentError++;
    } else {
      // check sub-folder
      if (rowData.subfolder !== undefined) {
        var subfolders = folder.getFoldersByName(rowData.subfolder);
        if (subfolders.hasNext()) {
          folder = subfolders.next();
        }
      }
      for (var j=0; j<fileList.length; j++) {
        var files = folder.getFilesByName(fileList[j]);
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
      Logger.log("Skipped (attachment error): "+rowData.emailAddressToSend);
      dataSheet.getRange(i+2,2).setValue("Error and Skipped (attachment error)");
      continue;
    }

    // [2022-04-21] Prepare inline images
    var inlineImages = new Object();
    for (const [key, value] of Object.entries(rowData)) {
      if (key.substring(0,7)=="imglink" && value != "") {
        var imgBlob = UrlFetchApp.fetch(value).getBlob().setName(key);
        inlineImages[key] = imgBlob;
      } else if (key.substring(0,7)=="imgfile" && value != "") {
        var files = folder.getFilesByName(value);
        if (files.hasNext()) {
          var imgBlob = files.next().getBlob().setName(key);
          inlineImages[key] = imgBlob;
        }
      } else if (key.substring(0,6)=="qrdata" && value != "") {
        var qrImgLnk = qrApiUrl + encodeURI(value);
        try {
          var imgBlob = UrlFetchApp.fetch(qrImgLnk).getBlob().setName(key);
          inlineImages[key] = imgBlob;
        } catch (e) {
          Logger.log("QR code fetch failed for " + value + ": " + e.message);
        }
      }
    }
   
    // Send out now
    var opts = new Object();
    if (attachments.length > 0) opts['attachments'] = attachments;
    if (ccList != "") opts['cc'] = ccList;
    if (bccList != "") opts['bcc'] = bccList;
    if (sendAs != "") opts['name'] = sendAs;
    if (replyTo != "") opts['replyTo'] = replyTo;
    if (noReply) opts['noReply'] = true;
    if (Object.keys(inlineImages).length>0) opts['inlineImages'] = inlineImages;
    opts['htmlBody'] = htmlBody;

    // Ref: https://developers.google.com/apps-script/reference/mail/mail-app#sendEmail(String,String,String,Object)
    MailApp.sendEmail(rowData.emailAddressToSend, emailSubject, emailBody, opts);
    let logMesage = "Mail sent to "+rowData.emailAddressToSend+": \""+emailSubject+"\" with "+attachments.length+" attachment(s).";
    updateDashboard_(ss, {"status": logMesage});

    // Mark "Done" in Datasheet
    dataSheet.getRange(i+2,2).setValue("Done");

    // Update email quota remaining
    MailApp.getRemainingDailyQuota();
    updateDashboard_(ss, {"quota": emailQuotaRemaining});

    // Log this action
    Logger.log(logMesage);

    // Checking time-oout
    if (isTimedOut_()) {
      // prepare trigger for re-run after 30 seconds
      Logger.log("Execution time Exceeded - Setting trigger")
      setRerunTrigger_();
      return;
    };
  }
  
  // Up to this point, all records processed
  updateDashboard_(ss, {"status": "Completed"});

  // Send the complete log to the person running the script
  // Ref: https://developers.google.com/apps-script/reference/base/logger#getLog()
  var recipient = Session.getActiveUser().getEmail();
  let timeStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
  var subject = "** Mail Merge Log ** (" + timeStr + ")";
  var body = Logger.getLog();
  let html = body.split("\n").join("<br>"); // [2024-11-11] Better mail log format
  MailApp.sendEmail(recipient, subject, body, { htmlBody: html, noReply: true });

  // Update email quota remaining
  // https://developers.google.com/apps-script/reference/mail/mail-app#getRemainingDailyQuota()
  emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Email quota remaining: "+emailQuotaRemaining);

  updateDashboard_(ss, {"quota": emailQuotaRemaining});

  // Clear triggers, if any, as it is the end nnow
  clearTriggers_();
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
          [ "Total Emails", '=countif(Datasheet!A:A,"Yes")+countif(Datasheet!A:A,TRUE)' ],
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
        [ "Yes", "", currentUser, "Template1", "", "", "1001", "Alice Chan", "Physics", "Chemistry", "Biology", "https://raw.githubusercontent.com/charlotte-lau-hk/MailMergeForBatchEmail/master/test-data/sunset2.jpg", "sunset1.jpg"],
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
        [ "No reply", false, "No reply: TRUE or FALSE" ],
        [ "QR API URL", "https://qrcode.tec-it.com/API/QRCode?quietzone=2&dpi=150&&data=", "Default: https://qrcode.tec-it.com/API/QRCode?quietzone=2&dpi=150&&data=" ]
      ],
      format: {
        "border": true,
        "firstRowBold": true,
        "firstColumnBold": true
      }
    },
    "Usage": {
      data: [
        [ "Mail Merge for Batch Email (v9.0)" ],
        [ "Maintained by: Charlotte Lau" ],
        [ "GitHub: https://github.com/charlotte-lau-hk/MailMergeForBatchEmail" ],
        [ "Guide: https://www.syclau.hk/mail-merge-for-batch-email" ],
        [ "To use: Make a copy of this file" ],
        [ "** Reference **" ],
        [ "Markdown Cheatsheet: https://github.com/adam-p/markdown-here/wiki/Markdown-Cheatsheet" ],
        [ "Online Markdown Editor: https://jbt.github.io/markdown-editor/" ],
        [ "Online QR Code Generator: https://qrcode.tec-it.com/en" ]
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
function isTimedOut_() {
  let now = Date.now();
  let lastRun = parseInt(PropertiesService.getDocumentProperties().getProperty("runTime"));
  if (isNaN(lastRun) || now - lastRun >= 330 * 1000) {
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
}
