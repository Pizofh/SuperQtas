function crearBackupManualQTAS(payload) {
  return crearBackupSpreadsheetQTAS_(Object.assign({
    daily: false,
    prefix: 'QTAS_BACKUP_ACTIVO'
  }, payload || {}));
}

function ejecutarBackupDiarioQTAS() {
  return crearBackupSpreadsheetQTAS_({
    daily: true,
    prefix: 'QTAS_BACKUP_ACTIVO'
  });
}

function instalarBackupDiarioQTAS(payload) {
  const settings = Object.assign({
    hour: 3
  }, payload || {});
  const hour = Math.max(0, Math.min(23, Math.floor(numero_(settings.hour) || 3)));

  limpiarTriggersBackupDiarioQTAS_();
  ScriptApp.newTrigger('ejecutarBackupDiarioQTAS')
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();

  const ss = SpreadsheetApp.getActive();
  const file = DriveApp.getFileById(ss.getId());
  const folder = asegurarCarpetaBackupsQTAS_(file);
  PropertiesService.getScriptProperties().setProperty('QTAS_BACKUP_HOUR', String(hour));

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    hour: hour,
    folderId: folder.getId(),
    folderName: folder.getName(),
    handler: 'ejecutarBackupDiarioQTAS'
  };
}

function desinstalarBackupDiarioQTAS() {
  const removed = limpiarTriggersBackupDiarioQTAS_();
  return {
    ok: true,
    removedTriggers: removed,
    handler: 'ejecutarBackupDiarioQTAS'
  };
}

function getEstadoBackupsQTAS() {
  const ss = SpreadsheetApp.getActive();
  const file = DriveApp.getFileById(ss.getId());
  const folder = asegurarCarpetaBackupsQTAS_(file);
  const triggerCount = ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'ejecutarBackupDiarioQTAS')
    .length;

  return {
    ok: true,
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    folderId: folder.getId(),
    folderName: folder.getName(),
    scheduledHour: numero_(PropertiesService.getScriptProperties().getProperty('QTAS_BACKUP_HOUR')) || 0,
    dailyTriggerInstalled: triggerCount > 0,
    triggerCount: triggerCount,
    destructiveOpsAllowed: operacionesDestructivasPermitidasQTAS_()
  };
}

function crearBackupSpreadsheetQTAS_(options) {
  const settings = Object.assign({
    daily: false,
    prefix: 'QTAS_BACKUP_ACTIVO'
  }, options || {});
  const sourceSpreadsheet = SpreadsheetApp.getActive();
  const sourceFile = DriveApp.getFileById(sourceSpreadsheet.getId());
  const folder = asegurarCarpetaBackupsQTAS_(sourceFile);
  const backupName = `${texto_(settings.prefix)}__${sourceSpreadsheet.getName()}`;
  let backupFile = buscarArchivoPorNombreExactoQTAS_(folder, backupName);
  let created = false;

  if (!backupFile) {
    backupFile = sourceFile.makeCopy(backupName, folder);
    created = true;
  }

  sincronizarArchivoBackupQTAS_(sourceSpreadsheet, SpreadsheetApp.openById(backupFile.getId()));
  return {
    ok: true,
    created: created,
    spreadsheetId: sourceSpreadsheet.getId(),
    spreadsheetName: sourceSpreadsheet.getName(),
    backupFileId: backupFile.getId(),
    backupName: backupName,
    folderId: folder.getId(),
    folderName: folder.getName()
  };
}

function asegurarCarpetaBackupsQTAS_(sourceFile) {
  const props = PropertiesService.getScriptProperties();
  const configuredId = texto_(props.getProperty('QTAS_BACKUP_FOLDER_ID'));

  if (configuredId) {
    try {
      return DriveApp.getFolderById(configuredId);
    } catch (error) {
      props.deleteProperty('QTAS_BACKUP_FOLDER_ID');
    }
  }

  const file = sourceFile || DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parentFolders = file.getParents();
  const parent = parentFolders.hasNext()
    ? parentFolders.next()
    : DriveApp.getRootFolder();
  const folderName = 'QTAS_Backups';
  const existingFolders = parent.getFoldersByName(folderName);
  const folder = existingFolders.hasNext()
    ? existingFolders.next()
    : parent.createFolder(folderName);

  props.setProperty('QTAS_BACKUP_FOLDER_ID', folder.getId());
  return folder;
}

function limpiarTriggersBackupDiarioQTAS_() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'ejecutarBackupDiarioQTAS');

  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  return triggers.length;
}

function buscarArchivoPorNombreExactoQTAS_(folder, fileName) {
  if (!folder || !fileName) return null;

  const files = folder.getFilesByName(fileName);
  return files.hasNext() ? files.next() : null;
}

function sincronizarArchivoBackupQTAS_(sourceSpreadsheet, backupSpreadsheet) {
  if (!sourceSpreadsheet || !backupSpreadsheet) return;

  const tempName = nombreHojaUnico_(backupSpreadsheet, '__TMP_QTAS_BACKUP__');
  const tempSheet = backupSpreadsheet.insertSheet(tempName);
  const backupSheets = backupSpreadsheet.getSheets().slice();

  backupSheets.forEach(sheet => {
    if (sheet.getSheetId() !== tempSheet.getSheetId()) {
      backupSpreadsheet.deleteSheet(sheet);
    }
  });

  sourceSpreadsheet.getSheets().forEach((sourceSheet, index) => {
    const copied = sourceSheet.copyTo(backupSpreadsheet);
    copied.setName(sourceSheet.getName());
    backupSpreadsheet.setActiveSheet(copied);
    backupSpreadsheet.moveActiveSheet(index + 1);
  });

  backupSpreadsheet.deleteSheet(tempSheet);
  backupSpreadsheet.setSpreadsheetTimeZone(sourceSpreadsheet.getSpreadsheetTimeZone());
}
