function getOrCreateCaseFolder_(rowData, caseId) {
  const existingId = getSystemValue_(rowData, '【システム】顧客フォルダID');
  if (existingId) {
    try {
      return DriveApp.getFolderById(existingId);
    } catch (error) {
      appendLog_('WARN', caseId, rowData.rowNumber, 'フォルダ再取得', error.message);
    }
  }

  const outputFolderId = getScriptProperty_(ASTRA_CONFIG.PROPERTY_KEYS.OUTPUT_FOLDER_ID, true);
  const outputFolder = DriveApp.getFolderById(outputFolderId);
  const submission = rowData.submission;
  const dateText = formatSubmissionDate_(submission.timestamp);
  const folderName = sanitizeDriveName_([
    dateText,
    submission.name,
    submission.business || '業務未判定',
    caseId
  ].join('_'));
  return outputFolder.createFolder(folderName);
}

function createCaseId_() {
  const datePart = Utilities.formatDate(new Date(), ASTRA_CONFIG.TIME_ZONE, 'yyyyMMdd');
  const randomPart = Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
  return 'AST-' + datePart + '-' + randomPart;
}

function formatSubmissionDate_(timestamp) {
  const raw = stringValue_(timestamp);
  const parsed = raw ? new Date(raw) : new Date();
  const date = isNaN(parsed.getTime()) ? new Date() : parsed;
  return Utilities.formatDate(date, ASTRA_CONFIG.TIME_ZONE, 'yyyy-MM-dd');
}
