// Copias nativas e inmutables. La retencion solo incluye copias diarias
// verificadas de este libro; las manuales y las copias antiguas se conservan.
function crearBackupManualQTAS(payload) {
  return crearBackupSpreadsheetQTAS_(Object.assign({}, payload || {}, { daily: false }));
}

function ejecutarBackupDiarioQTAS() {
  return crearBackupSpreadsheetQTAS_({ daily: true });
}

function instalarBackupDiarioQTAS(payload) {
  return withScriptLock_('instalar backup diario', () => {
    const settings = Object.assign({ hour: 3 }, payload || {});
    const requestedHour = Number(settings.hour);
    if (!Number.isInteger(requestedHour) || requestedHour < 0 || requestedHour > 23) {
      throw new Error('La hora del backup debe ser un entero entre 0 y 23.');
    }
    const ss = obtenerSpreadsheetBackupsQTAS_();
    const folder = asegurarCarpetaBackupsQTAS_(DriveApp.getFileById(ss.getId()));
    const oldTriggers = ScriptApp.getProjectTriggers()
      .filter(trigger => trigger.getHandlerFunction() === 'ejecutarBackupDiarioQTAS');
    // Crear primero: si falla, el activador anterior sigue instalado.
    const trigger = ScriptApp.newTrigger('ejecutarBackupDiarioQTAS')
      .timeBased().everyDays(1).atHour(requestedHour)
      .inTimezone(ss.getSpreadsheetTimeZone()).create();
    PropertiesService.getScriptProperties().setProperties({
      QTAS_BACKUP_SOURCE_ID: ss.getId(),
      QTAS_BACKUP_HOUR: String(requestedHour)
    });
    oldTriggers.forEach(old => ScriptApp.deleteTrigger(old));
    return {
      ok: true, spreadsheetId: ss.getId(), spreadsheetName: ss.getName(),
      hour: requestedHour, timeZone: ss.getSpreadsheetTimeZone(),
      folderId: folder.getId(), folderName: folder.getName(),
      retainedDailyBackups: 3, triggerId: trigger.getUniqueId(),
      handler: 'ejecutarBackupDiarioQTAS'
    };
  });
}

function desinstalarBackupDiarioQTAS() {
  return { ok: true, removedTriggers: limpiarTriggersBackupDiarioQTAS_(),
    handler: 'ejecutarBackupDiarioQTAS' };
}

function getEstadoBackupsQTAS() {
  const ss = obtenerSpreadsheetBackupsQTAS_();
  const props = PropertiesService.getScriptProperties();
  const folderId = texto_(props.getProperty('QTAS_BACKUP_FOLDER_ID'));
  const triggerCount = ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'ejecutarBackupDiarioQTAS').length;
  return {
    ok: true, spreadsheetId: ss.getId(), spreadsheetName: ss.getName(),
    folderId: folderId, scheduledHour: Number(props.getProperty('QTAS_BACKUP_HOUR') || 3),
    timeZone: ss.getSpreadsheetTimeZone(), retainedDailyBackups: 3,
    dailyTriggerInstalled: triggerCount > 0, triggerCount: triggerCount,
    lastSuccessfulBackup: props.getProperty('QTAS_BACKUP_LAST_SUCCESS') || '',
    lastError: props.getProperty('QTAS_BACKUP_LAST_ERROR') || '',
    destructiveOpsAllowed: operacionesDestructivasPermitidasQTAS_()
  };
}

function obtenerSpreadsheetBackupsQTAS_() {
  const active = SpreadsheetApp.getActive();
  const sourceId = texto_(PropertiesService.getScriptProperties().getProperty('QTAS_BACKUP_SOURCE_ID'));
  if (active) {
    if (sourceId && sourceId !== active.getId()) {
      throw new Error('El libro activo no coincide con el origen configurado del backup.');
    }
    return active;
  }
  if (sourceId) return SpreadsheetApp.openById(sourceId);
  throw new Error('Instala el backup diario desde el proyecto vinculado al libro original.');
}

function crearBackupSpreadsheetQTAS_(options) {
  return withScriptLock_('crear backup', () => {
    const props = PropertiesService.getScriptProperties();
    try {
      const daily = Boolean(options && options.daily);
      const ss = obtenerSpreadsheetBackupsQTAS_();
      const sourceFile = DriveApp.getFileById(ss.getId());
      const folder = asegurarCarpetaBackupsQTAS_(sourceFile);
      const now = new Date();
      const day = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
      const managed = listarBackupsDiariosQTAS_(folder, ss.getId());
      if (daily && managed.length && managed[0].metadata.day > day) {
        throw new Error('Hay un backup de una fecha posterior. Se conserva la retencion sin crear otra copia.');
      }
      const existing = daily ? managed.find(item => item.metadata.day === day) : null;
      let file;
      let created = false;
      if (existing) {
        // Un reintento no reemplaza la foto que ya se tomo ese dia.
        validarManifestBackupQTAS_(existing.metadata.manifest,
          SpreadsheetApp.openById(existing.file.getId()));
        file = existing.file;
      } else {
        const manifest = manifestSpreadsheetBackupQTAS_(ss);
        const kind = daily ? 'DIARIO' : 'MANUAL';
        const suffix = daily ? day : Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd_HH-mm-ss')
          + '_' + Utilities.getUuid().slice(0, 8);
        const name = `QTAS_BACKUP_${kind}__${ss.getId()}__${suffix}__${ss.getName()}`;
        file = sourceFile.makeCopy(name + '__PENDIENTE', folder);
        validarManifestBackupQTAS_(manifest, SpreadsheetApp.openById(file.getId()));
        file.setName(name);
        file.setDescription(JSON.stringify({
          type: 'QTAS_BACKUP', version: 2, sourceId: ss.getId(),
          kind: daily ? 'daily' : 'manual', day: day, verified: true,
          manifest: manifest
        }));
        created = true;
      }
      // Ninguna copia previa se retira hasta que la nueva esta verificada.
      const removedIds = daily ? rotarBackupsDiariosQTAS_(folder, ss.getId(), file.getId()) : [];
      props.setProperty('QTAS_BACKUP_LAST_SUCCESS', now.toISOString());
      props.deleteProperty('QTAS_BACKUP_LAST_ERROR');
      return {
        ok: true, created: created, daily: daily, day: day,
        spreadsheetId: ss.getId(), spreadsheetName: ss.getName(),
        backupFileId: file.getId(), backupName: file.getName(),
        folderId: folder.getId(), folderName: folder.getName(),
        retainedDailyBackups: 3, removedBackupIds: removedIds
      };
    } catch (error) {
      props.setProperty('QTAS_BACKUP_LAST_ERROR', String(error.message || error));
      throw error;
    }
  });
}

function manifestSpreadsheetBackupQTAS_(ss) {
  return {
    timeZone: ss.getSpreadsheetTimeZone(),
    sheets: ss.getSheets().map(sheet => {
      const range = sheet.getDataRange();
      const values = range.getValues();
      const formulas = range.getFormulas();
      const content = values.map((row, rowIndex) => row.map((value, colIndex) => {
        const formula = formulas[rowIndex][colIndex];
        if (formula) return ['formula', formula];
        if (value instanceof Date) return ['date', value.toISOString()];
        return [typeof value, value];
      }));
      const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(content));
      return { name: sheet.getName(), rows: range.getNumRows(), columns: range.getNumColumns(),
        digest: Utilities.base64Encode(digest) };
    })
  };
}

function validarManifestBackupQTAS_(expected, copy) {
  if (!expected || JSON.stringify(manifestSpreadsheetBackupQTAS_(copy)) !== JSON.stringify(expected)) {
    throw new Error('La copia no coincide con el libro original. Se conservan los backups previos.');
  }
}

function listarBackupsDiariosQTAS_(folder, sourceId) {
  const files = folder.getFiles();
  const backups = [];
  const prefix = `QTAS_BACKUP_DIARIO__${sourceId}__`;
  while (files.hasNext()) {
    const file = files.next();
    if (file.getId() === sourceId || !file.getName().startsWith(prefix)) continue;
    let metadata;
    try { metadata = JSON.parse(file.getDescription() || '{}'); } catch (error) { continue; }
    if (metadata.type !== 'QTAS_BACKUP' || metadata.version !== 2 || !metadata.verified ||
        metadata.kind !== 'daily' || metadata.sourceId !== sourceId ||
        !/^\d{4}-\d{2}-\d{2}$/.test(metadata.day || '') || !metadata.manifest ||
        !file.getName().startsWith(prefix + metadata.day + '__') ||
        file.getName().endsWith('__PENDIENTE')) continue;
    backups.push({ file: file, metadata: metadata });
  }
  return backups.sort((a, b) => b.metadata.day.localeCompare(a.metadata.day) ||
    b.file.getDateCreated().getTime() - a.file.getDateCreated().getTime());
}

function rotarBackupsDiariosQTAS_(folder, sourceId, currentId) {
  const backups = listarBackupsDiariosQTAS_(folder, sourceId);
  if (!backups.some(item => item.file.getId() === currentId)) {
    throw new Error('La copia actual no esta verificada; no se rota ningun backup.');
  }
  // Verificar tambien las copias que se van a conservar antes de retirar otra.
  backups.slice(0, 3).forEach(item => validarManifestBackupQTAS_(item.metadata.manifest,
    SpreadsheetApp.openById(item.file.getId())));
  const removed = [];
  backups.slice(3).forEach(item => {
    if (item.file.getId() === sourceId || item.file.getId() === currentId) return;
    item.file.setTrashed(true); // Recuperable desde la papelera de Drive.
    removed.push(item.file.getId());
  });
  return removed;
}

function asegurarCarpetaBackupsQTAS_(sourceFile) {
  const props = PropertiesService.getScriptProperties();
  const configuredId = texto_(props.getProperty('QTAS_BACKUP_FOLDER_ID'));
  if (configuredId) return DriveApp.getFolderById(configuredId);
  const parentFolders = sourceFile.getParents();
  const parent = parentFolders.hasNext() ? parentFolders.next() : DriveApp.getRootFolder();
  const folders = parent.getFoldersByName('QTAS_Backups');
  const folder = folders.hasNext() ? folders.next() : parent.createFolder('QTAS_Backups');
  props.setProperty('QTAS_BACKUP_FOLDER_ID', folder.getId());
  return folder;
}

function limpiarTriggersBackupDiarioQTAS_() {
  const triggers = ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'ejecutarBackupDiarioQTAS');
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  return triggers.length;
}
