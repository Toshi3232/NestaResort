function backupSpreadsheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(spreadsheet.getId());
  
  // バックアップフォルダのID（事前にGoogle Driveで作成しておく）
  var backupFolderId = "1BAsqjcLzRURmdaPFhyJj48A6JSu8XAwO";
  var backupFolder = DriveApp.getFolderById(backupFolderId);
  
  // バックアップ用のファイル名（日時を付与）
  var date = new Date();
  var timestamp = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm");
  var backupFileName = spreadsheet.getName() + "_backup_" + timestamp;
  
  // ファイルをコピーしてバックアップフォルダに保存
  var newFile = file.makeCopy(backupFileName, backupFolder);
  
  // **バックアップフォルダ内の古いファイルを削除**
  cleanupOldBackups(spreadsheet.getName(), backupFolder);

  Logger.log("バックアップ完了: " + newFile.getName());
}

function cleanupOldBackups(baseName, backupFolder) {
  var files = backupFolder.getFiles();
  var backupFiles = [];

  // **バックアップフォルダ内の対象ファイルをリスト化**
  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();
    
    // 指定したスプレッドシート名のバックアップファイルのみ対象
    if (name.startsWith(baseName + "_backup_")) {
      backupFiles.push({ file: file, date: file.getLastUpdated() });
    }
  }

  // **日付の新しい順にソート**
  backupFiles.sort(function(a, b) {
    return b.date - a.date; // 降順（新しいものが先）
  });

  // **最新3つを残し、それ以外を削除**
  if (backupFiles.length > 30) {
    for (var i = 30; i < backupFiles.length; i++) {
      backupFiles[i].file.setTrashed(true);
      Logger.log("削除: " + backupFiles[i].file.getName());
    }
  }
}