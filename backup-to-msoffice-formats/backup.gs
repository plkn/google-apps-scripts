// entry point, run it in the App Script editor
function run() {
  var token = ScriptApp.getOAuthToken();
  var maxNestingLevel = 10;
  var rootFolder = DriveApp.getRootFolder();
  var backupRoot = rootFolder.getFoldersByName("__backup__");

  if (!backupRoot.hasNext()) {
    backupRoot = DriveApp.createFolder("__backup__");
  } else {
    backupRoot = backupRoot.next();
  }

  backupFolder(rootFolder, backupRoot, 0);

  function backupFolder(folderToBackup, parentFolderToBackupTo, nestingLevel) {
    if (nestingLevel > maxNestingLevel) {
      return;
    }

    nestingLevel++;
    var folderName = folderToBackup.getName();
    if (folderName == "__convert__" || folderName == "__backup__") {
      Logger.log("Skip folder " + folderName);
      return;
    }

    Logger.log("DIR:" + "\t".repeat(nestingLevel) + folderName);

    var folderToBackupTo = parentFolderToBackupTo.getFoldersByName(folderName);

    if (folderToBackupTo.hasNext()) {
      folderToBackupTo = folderToBackupTo.next();
    } else {
      folderToBackupTo = parentFolderToBackupTo.createFolder(folderName);
    }

    backupDocs(folderToBackup, folderToBackupTo, nestingLevel);
    backupTables(folderToBackup, folderToBackupTo, nestingLevel);

    var nestedFolders = folderToBackup.getFolders();

    while (nestedFolders.hasNext()) {
      var nestedFolder = nestedFolders.next();
      backupFolder(nestedFolder, folderToBackupTo, nestingLevel);
    }
  }

  function getBackedFiles(folderToBackupTo, mimeType) {
    var backedFiles = [];
    var backedFilesIter = folderToBackupTo.getFilesByType(mimeType);
    while (backedFilesIter.hasNext()) {
      backedFiles.push(backedFilesIter.next().getName());
    }
    return backedFiles;
  }

  function backupFiles(
    folderToBackup,
    folderToBackupTo,
    officeMime,
    gdriveMime,
    exportFormat,
    nestingLevel
  ) {
    var backedFiles = getBackedFiles(folderToBackupTo, officeMime);

    var docs = folderToBackup.getFilesByType(gdriveMime);
    while (docs.hasNext()) {
      var indent = "\t".repeat(nestingLevel + 1);
      var doc = docs.next();

      var backupFileName = doc.getName() + "." + exportFormat;

      if (backedFiles.findIndex((_) => _ == backupFileName) >= 0) {
        Logger.log(
          `${indent}SKIP: ${folderToBackup.getName()}\\${doc.getName()}`
        );
        continue;
      }

      Logger.log(
        `${indent}BACK: ${folderToBackup.getName()}\\${doc.getName()}`
      );

      var url =
        exportFormat == "docx"
          ? `https://docs.google.com/feeds/download/documents/export/Export?id=${doc.getId()}&exportFormat=${exportFormat}`
          : `https://docs.google.com/spreadsheets/d/${doc.getId()}/export?format=${exportFormat}`;

      var blob = UrlFetchApp.fetch(url, {
        headers: {
          Authorization: "Bearer " + token,
        },
      }).getBlob();

      DriveApp.createFile(blob)
        .setName(backupFileName)
        .moveTo(folderToBackupTo);
    }
  }

  function backupTables(folderToBackup, folderToBackupTo, nestingLevel) {
    backupFiles(
      folderToBackup,
      folderToBackupTo,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.google-apps.spreadsheet",
      "xlsx",
      nestingLevel
    );
  }

  function backupDocs(folderToBackup, folderToBackupTo, nestingLevel) {
    backupFiles(
      folderToBackup,
      folderToBackupTo,
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "application/vnd.google-apps.document",
      "docx",
      nestingLevel
    );
  }

  function backupDocs_old(folderToBackup, folderToBackupTo) {
    backedFiles = getBackedFiles(
      folderToBackupTo,
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );

    var docs = folderToBackup.getFilesByType(
      "application/vnd.google-apps.document"
    );
    while (docs.hasNext()) {
      var indent = "\t".repeat(nestingLevel + 1);
      var doc = docs.next();

      var backupFileName = doc.getName() + ".docx";

      if (backedFiles.findIndex((_) => _ == backupFileName) > 0) {
        Logger.log(indent + "SKIP: " + doc.getName());
        continue;
      }

      Logger.log(indent + "BACK: " + doc.getName());

      var blob = UrlFetchApp.fetch(
        "https://docs.google.com/feeds/download/documents/export/Export?id=" +
          doc.getId() +
          "&exportFormat=docx",
        {
          headers: {
            Authorization: "Bearer " + token,
          },
        }
      ).getBlob();

      DriveApp.createFile(blob)
        .setName(backupFileName)
        .moveTo(folderToBackupTo);
    }
  }
}

function logFiles(files, prefix = "") {
  while (files.hasNext()) {
    var file = files.next();
    Logger.log(file.getName() + file.getId());
  }
}
