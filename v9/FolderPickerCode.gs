function FP_showFolderPicker(handler) {
  var html = HtmlService.createTemplateFromFile('FolderPicker')
      .evaluate()
      .setWidth(480)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setHeight(480);
  SpreadsheetApp.getUi()
      .showModalDialog(html, "📁 Folder Picker");
  // saving the handler function name in user properties
  PropertiesService.getUserProperties().setProperty("FP_handler", handler.name);
}

// get list of subfolders and list of files
function FP_getFolderContents(folderId) {
  try {
    var location = DriveApp.getFolderById(folderId);
    var folders = location.getFolders();
    var files = location.getFiles();
    var folderList = [];
    var fileList = [];
    while (folders.hasNext()) {
      var folder = folders.next();
      folderList.push([folder.getId(), folder.getName()]);
    }
    while (files.hasNext()) {
      var file = files.next();
      fileList.push(file.getName() + " (" + (file.getSize() / 1024).toFixed(1) + " KB)");
    }
    folderList.sort((a, b) => a[1].localeCompare(b[1]));
    fileList.sort();
    return { folders: folderList, files: fileList };
  } catch (e) {
    Logger.log("Fetch error: " + e.message);
    return { folders: [], files: [] };
  }
}

// [2026-11-28] Functions the folder picker may call (replaces eval)
const FP_HANDLERS_ = {
  "setAttachmentFolderId_": (folderId) => setAttachmentFolderId_(folderId)
};

// return here
function FP_returnFolderId(folderId) {
  var handlerName = PropertiesService.getUserProperties().getProperty("FP_handler");
  var handlerFunc = FP_HANDLERS_[handlerName];
  if (!handlerFunc) {
    Logger.log("Invalid handler: " + handlerName);
    return;
  }
  handlerFunc(folderId);
  PropertiesService.getUserProperties().deleteProperty("FP_handler"); // clear after use
}
