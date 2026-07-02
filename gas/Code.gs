function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('アストラ書類システム')
    .addItem('初期設定を実行', 'setupAstraSystem')
    .addSeparator()
    .addItem('未処理を今すぐ実行', 'processPendingSubmissions')
    .addItem('選択中の回答を再生成', 'reprocessActiveRow')
    .addItem('選択中の回答のPDFを書き出す', 'exportActiveRowPdf')
    .addSeparator()
    .addItem('セルフテストを実行', 'runAstraSelfTests')
    .addToUi();
}

function setupAstraSystem() {
  setDefaultScriptProperties_();
  const responseSheet = getResponseSheet_();
  assertRequiredSourceHeaders_(responseSheet);
  ensureSystemColumns_(responseSheet);
  const legacyCount = markLegacyRowsSkipped_(responseSheet);
  ensureManagementSheets_();
  ensureDefaultTemplates_();
  installAstraTriggers_();
  writeSettingsSnapshot_(getResponseSpreadsheet_().getSheetByName(ASTRA_CONFIG.SHEETS.SETTINGS));
  const legacyMessage = legacyCount
    ? ' 既存の回答' + legacyCount + '行は「' + ASTRA_CONFIG.STATUS.SKIPPED + '」にしました。処理したい行は処理状態を「' + ASTRA_CONFIG.STATUS.PENDING + '」へ変更してください。'
    : '';
  appendLog_('INFO', '', '', '初期設定', '初期設定とトリガー作成が完了しました。' + legacyMessage);
  getResponseSpreadsheet_().toast('初期設定が完了しました。' + legacyMessage, 'アストラ書類システム', 10);
}

function processPendingSubmissions() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const sheet = getResponseSheet_();
    ensureSystemColumns_(sheet);
    markExhaustedProcessingRows_(sheet);
    const rows = findProcessableRows_(sheet, ASTRA_CONFIG.MAX_ROWS_PER_RUN);
    rows.forEach(function(rowNumber) {
      try {
        processSubmissionRow_(sheet, rowNumber, false);
      } catch (error) {
        appendLog_('ERROR', '', rowNumber, '定期処理', error.stack || error.message);
      }
    });
  } finally {
    lock.releaseLock();
  }
}

function processSubmissionWithLock_(rowNumber, force) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) {
    throw new Error('別の処理が実行中です。後ほど自動再試行します。');
  }
  try {
    processSubmissionRow_(getResponseSheet_(), rowNumber, force);
  } finally {
    lock.releaseLock();
  }
}

function processSubmissionRow_(sheet, rowNumber, force) {
  ensureSystemColumns_(sheet);
  let rowData = readSubmissionRow_(sheet, rowNumber);
  initializeAddressReviewFields_(sheet, rowNumber, rowData);
  rowData = readSubmissionRow_(sheet, rowNumber);
  const previousStatus = getSystemValue_(rowData, '【システム】処理状態');
  if (!force && previousStatus === ASTRA_CONFIG.STATUS.GENERATED) return;

  const previousAttemptCount = Number(getSystemValue_(rowData, '【システム】試行回数') || 0);
  const caseId = getSystemValue_(rowData, '【システム】案件ID') || createCaseId_();
  const validation = validateSubmission_(rowData.submission);
  const now = new Date();

  if (validation.errors.length) {
    writeSystemValues_(sheet, rowNumber, {
      '【システム】処理状態': ASTRA_CONFIG.STATUS.NEEDS_INPUT,
      '【システム】案件ID': caseId,
      '【システム】確認状態': ASTRA_CONFIG.REVIEW_STATUS.NEEDS_REVIEW,
      '【システム】確認事項': validation.errors.concat(validation.warnings).join('\n'),
      '【システム】最終処理日時': now,
      '【システム】エラー内容': ''
    });
    appendLog_('WARN', caseId, rowNumber, '入力検査', validation.errors.join(' / '));
    return;
  }

  writeSystemValues_(sheet, rowNumber, {
    '【システム】処理状態': ASTRA_CONFIG.STATUS.PROCESSING,
    '【システム】案件ID': caseId,
    '【システム】試行回数': previousAttemptCount + 1,
    '【システム】最終処理日時': now,
    '【システム】エラー内容': ''
  });

  try {
    rowData = readSubmissionRow_(sheet, rowNumber);
    const caseFolder = getOrCreateCaseFolder_(rowData, caseId);
    writeSystemValues_(sheet, rowNumber, {
      '【システム】顧客フォルダID': caseFolder.getId(),
      '【システム】顧客フォルダURL': caseFolder.getUrl()
    });
    rowData = readSubmissionRow_(sheet, rowNumber);
    const documents = generateDocuments_(rowData, caseId, caseFolder, validation, force);
    const documentUrls = documents.map(function(document) { return document.url; });
    const reviewMessages = validation.warnings.length
      ? validation.warnings
      : ['自動検査では警告なし。行政書士による最終確認は必要です。'];

    writeSystemValues_(sheet, rowNumber, {
      '【システム】処理状態': ASTRA_CONFIG.STATUS.GENERATED,
      '【システム】案件ID': caseId,
      '【システム】確認状態': ASTRA_CONFIG.REVIEW_STATUS.NEEDS_REVIEW,
      '【システム】確認事項': reviewMessages.join('\n'),
      '【システム】顧客フォルダID': caseFolder.getId(),
      '【システム】顧客フォルダURL': caseFolder.getUrl(),
      '【システム】生成書類URL': documentUrls.join('\n'),
      '【システム】最終処理日時': new Date(),
      '【システム】エラー内容': ''
    });
    appendLog_('INFO', caseId, rowNumber, '書類生成', documents.length + '件の下書きを生成しました。');
  } catch (error) {
    writeSystemValues_(sheet, rowNumber, {
      '【システム】処理状態': ASTRA_CONFIG.STATUS.ERROR,
      '【システム】確認状態': ASTRA_CONFIG.REVIEW_STATUS.NEEDS_REVIEW,
      '【システム】最終処理日時': new Date(),
      '【システム】エラー内容': stringValue_(error.message).slice(0, 1000)
    });
    notifySystemError_(caseId, rowNumber, error);
    throw error;
  }
}

function initializeAddressReviewFields_(sheet, rowNumber, rowData) {
  const currentAddress = getSystemValue_(rowData, '【システム】書類転記住所');
  const currentStatus = getSystemValue_(rowData, '【システム】住所確認状態');
  const updates = {};
  if (!currentAddress && rowData.submission.addressOriginal) {
    updates['【システム】書類転記住所'] = rowData.submission.addressOriginal;
  }
  if (!currentStatus) {
    updates['【システム】住所確認状態'] = rowData.submission.residentRecordAttachments
      ? ASTRA_CONFIG.ADDRESS_REVIEW_STATUS.NEEDS_COMPARISON
      : ASTRA_CONFIG.ADDRESS_REVIEW_STATUS.NO_ATTACHMENT;
  }
  if (Object.keys(updates).length) writeSystemValues_(sheet, rowNumber, updates);
}

function reprocessActiveRow() {
  const rowNumber = getActiveResponseRowNumber_('再生成したい行を選択してください。');
  processSubmissionWithLock_(rowNumber, true);
  getResponseSpreadsheet_().toast('選択行を再生成しました。', 'アストラ書類システム', 5);
}

function exportActiveRowPdf() {
  const rowNumber = getActiveResponseRowNumber_('PDFを書き出したい行を選択してください。');
  const sheet = getResponseSheet_();
  const rowData = readSubmissionRow_(sheet, rowNumber);
  const caseId = getSystemValue_(rowData, '【システム】案件ID');
  const documentUrls = getSystemValue_(rowData, '【システム】生成書類URL')
    .split('\n')
    .map(function(url) { return url.trim(); })
    .filter(function(url) { return Boolean(url); });
  if (!documentUrls.length) {
    throw new Error('この行には生成書類がありません。先に書類を生成してください。');
  }
  const folderId = getSystemValue_(rowData, '【システム】顧客フォルダID');
  if (!folderId) {
    throw new Error('顧客フォルダIDが未設定のためPDFを保存できません。');
  }
  const folder = DriveApp.getFolderById(folderId);
  const suffix = Utilities.formatDate(new Date(), ASTRA_CONFIG.TIME_ZONE, 'yyyyMMdd_HHmmss');
  const created = [];

  documentUrls.forEach(function(url) {
    const match = url.match(/\/d\/([-\w]+)/) || url.match(/[-\w]{25,}/);
    if (!match) return;
    const file = DriveApp.getFileById(match[1] || match[0]);
    const pdfBlob = file.getAs(MimeType.PDF).setName(file.getName() + '_' + suffix + '.pdf');
    const pdfFile = folder.createFile(pdfBlob);
    appendDocumentRecord_({
      caseId: caseId,
      rowNumber: rowNumber,
      type: 'PDF手動書き出し',
      fileId: pdfFile.getId(),
      name: pdfFile.getName(),
      url: file.getUrl(),
      pdfUrl: pdfFile.getUrl()
    });
    created.push(pdfFile.getName());
  });

  if (!created.length) {
    throw new Error('生成書類URLからファイルを特定できませんでした。');
  }
  appendLog_('INFO', caseId, rowNumber, 'PDF書き出し', created.length + '件のPDFを顧客フォルダへ保存しました。');
  getResponseSpreadsheet_().toast(created.length + '件のPDFを書き出しました。', 'アストラ書類システム', 5);
}

function getActiveResponseRowNumber_(actionMessage) {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== getResponseSheetName_()) {
    throw new Error('回答シートで' + actionMessage);
  }
  const rowNumber = sheet.getActiveRange().getRow();
  if (rowNumber <= ASTRA_CONFIG.HEADER_ROW) {
    throw new Error('ヘッダー以外の回答行を選択してください。');
  }
  return rowNumber;
}

function notifySystemError_(caseId, rowNumber, error) {
  const email = getScriptProperty_(ASTRA_CONFIG.PROPERTY_KEYS.NOTIFICATION_EMAIL, false);
  if (!email) return;
  MailApp.sendEmail({
    to: email,
    subject: '【アストラ書類システム】処理エラー ' + caseId,
    body: '回答行: ' + rowNumber + '\n案件ID: ' + caseId + '\n\n' + (error.stack || error.message)
  });
}
