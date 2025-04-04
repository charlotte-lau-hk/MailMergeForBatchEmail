// [2025-03-26] New feature for v9
// Preview merged emails
function previewMergedEmails_() {
  let ui = SpreadsheetApp.getUi();
  if (!validateSettings_()) {
    ui.alert("⚠️ Settings Invalid",
             "Recheck the Settings please",
             ui.ButtonSet.OK);
    return;
  }

  let html = HtmlService.createTemplateFromFile("PreviewMerged")
      .evaluate()
      .setWidth(960)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setHeight(640);
  ui.showModalDialog(html, "🔍 Preview Merged Emails");
}

function setupPreview() {
  const documentProperties = PropertiesService.getDocumentProperties();
  let settings = documentProperties.getProperties(); // at this stage, property must exist
  let { ccList, bccList, folderId } = settings;

  let ss = SpreadsheetApp.getActiveSpreadsheet();

  let folderName = "N/A";
  let folder = null;
  if (folderId != "") {
    try {
      folder = DriveApp.getFolderById(folderId);
      folderName = folder.getName();
    } catch(e) {
      folderId = "⚠️"+folderId+" (Error: "+e.message+")";
    }
  }
  let dataSheet = ss.getSheetByName("Datasheet");
  let numRows = dataSheet.getMaxRows() - 1;
  let rangeData = dataSheet.getRange(2, 1, numRows, 2).getValues();
  let setupInfo = {
    "ccList": ccList,
    "bccList": bccList,
    "folderInfo": { "id": folderId, "name": folderName },
    "list": []
  }
  // check rows to send and not yet send
  setupInfo["list"] = listRowsToSend_();
  /*
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
    setupInfo["list"].push(i+1);
  };
  */

  return setupInfo;
}

function previewRow(rowNum) {
  const documentProperties = PropertiesService.getDocumentProperties();
  let settings = documentProperties.getProperties(); // at this stage, property must exist
  let { folderId, qrApiUrl } = settings;

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let dataSheet = ss.getSheetByName("Datasheet");
  let dataRange = dataSheet.getRange(rowNum+1, 1, 1, dataSheet.getMaxColumns());

  // Create one JavaScript object per row of data.
  let objects = getRowsData_(dataSheet, dataRange, 1); // columnHeadersRowIndex is 1
  let rowData = objects[0];
  let row = {
    "row": rowNum,
    "to": rowData.emailAddressToSend,
    "template": rowData.template,
    "subfolderInfo": {"id": null, "name": "N/A" },
    "attachments": []
  };

  // get list of attachments
  let fileList = [];
  let folder = null;
  if (folderId != "") {
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch(e) {
      Logger.log("Cannot get folder: " + folderId);
    }
  }

  if ((rowData.attachmentList !== undefined) && (rowData.attachmentList != "")) {
    fileList = rowData.attachmentList.split(",").map(s => s.trim());
    if ((folder != null) && (rowData.subfolder !== undefined)) {
      let subfolders = folder.getFoldersByName(rowData.subfolder);
      if (subfolders.hasNext()) {
        folder = subfolders.next();
        row["subfolderInfo"]["name"] = rowData.subfolder;
        row["subfolderInfo"]["id"] = folder.getId();
      }
    }
    if (fileList[0].length > 0) {
      for (let j=0; j<fileList.length; j++) {
        if (folder != null) {
          let files = folder.getFilesByName(fileList[j]);
          if (files.hasNext()) {
            let file = files.next();
            row["attachments"].push({"id": file.getId(), "name": fileList[j]});
          } else {
            row["attachments"].push({"id": null, "name": fileList[j]+" (⚠️ File not found)"});
          }
        } else {
          row["attachments"].push({"id": null, "name": fileList[j]+" (⚠️ Folder not found)"});
        }
      }
    }
  }

  // get email subject and body in markdown format
  let templateSheet = ss.getSheetByName(row["template"]);
  if (templateSheet) {
    // template exists, fill in content here
    let subjTemplate = templateSheet.getRange("B1").getValue();
    let bodyTemplate = templateSheet.getRange("B2").getValue();
    row["subject"] = fillInTemplateFromObject_(subjTemplate, rowData);
    row["body"] = fillInTemplateFromObject_(bodyTemplate, rowData, { "folder": folder, "qrApiUrl": qrApiUrl });
  } else {
    // template does not exist
    row["template"] += "⚠️ Not found: "+row["template"];
    row["subject"] = "N/A";
    row["body"] = "N/A";
  }

  return row;
}

