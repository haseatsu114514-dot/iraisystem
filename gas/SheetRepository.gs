/**
 * 回答シートと管理シートへの読み書きを一か所に集約する。
 */
function getResponseSpreadsheet_() {
  const id = getScriptProperty_(ASTRA_CONFIG.PROPERTY_KEYS.RESPONSE_SPREADSHEET_ID, false);
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error(
    'このスクリプトが回答スプレッドシートに紐付いていません。' +
    'Script Properties に RESPONSE_SPREADSHEET_ID を設定してください。'
  );
}

function getResponseSheet_() {
  const spreadsheet = getResponseSpreadsheet_();
  const sheetName = getResponseSheetName_();
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(
      '回答シート「' + sheetName + '」が見つかりません。' +
      'シート名が異なる場合は Script Properties の RESPONSE_SHEET_NAME に実際の名前を設定してください。'
    );
  }
  return sheet;
}

function ensureSystemColumns_(sheet) {
  const headers = readHeaders_(sheet);
  let nextColumn = Math.max(headers.length, sheet.getLastColumn()) + 1;
  let addedColumns = false;

  ASTRA_CONFIG.SYSTEM_HEADERS.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      sheet.getRange(ASTRA_CONFIG.HEADER_ROW, nextColumn).setValue(header);
      headers.push(header);
      nextColumn += 1;
      addedColumns = true;
    }
  });

  const statusColumn = headers.indexOf('【システム】処理状態') + 1;
  const reviewColumn = headers.indexOf('【システム】確認状態') + 1;
  const addressReviewColumn = headers.indexOf('【システム】住所確認状態') + 1;
  if (addedColumns && sheet.getMaxRows() > ASTRA_CONFIG.HEADER_ROW) {
    const rowCount = sheet.getMaxRows() - ASTRA_CONFIG.HEADER_ROW;
    sheet.getRange(ASTRA_CONFIG.HEADER_ROW + 1, statusColumn, rowCount).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(Object.keys(ASTRA_CONFIG.STATUS).map(function(key) {
          return ASTRA_CONFIG.STATUS[key];
        }), true)
        .setAllowInvalid(true)
        .build()
    );
    sheet.getRange(ASTRA_CONFIG.HEADER_ROW + 1, reviewColumn, rowCount).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(Object.keys(ASTRA_CONFIG.REVIEW_STATUS).map(function(key) {
          return ASTRA_CONFIG.REVIEW_STATUS[key];
        }), true)
        .setAllowInvalid(true)
        .build()
    );
    sheet.getRange(ASTRA_CONFIG.HEADER_ROW + 1, addressReviewColumn, rowCount).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(Object.keys(ASTRA_CONFIG.ADDRESS_REVIEW_STATUS).map(function(key) {
          return ASTRA_CONFIG.ADDRESS_REVIEW_STATUS[key];
        }), true)
        .setAllowInvalid(true)
        .build()
    );
  }
  sheet.setFrozenRows(1);
  return createHeaderIndex_(readHeaders_(sheet));
}

function readHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) return [];
  return sheet.getRange(ASTRA_CONFIG.HEADER_ROW, 1, 1, lastColumn).getDisplayValues()[0]
    .map(function(value) { return stringValue_(value); });
}

function assertRequiredSourceHeaders_(sheet) {
  const headers = readHeaders_(sheet);
  const required = Object.keys(ASTRA_CONFIG.SOURCE_HEADERS).map(function(key) {
    return ASTRA_CONFIG.SOURCE_HEADERS[key];
  });
  const missing = required.filter(function(header) { return headers.indexOf(header) === -1; });
  if (missing.length) {
    throw new Error('回答シートに必要な列がありません: ' + missing.join(', '));
  }
}

function createHeaderIndex_(headers) {
  const index = {};
  headers.forEach(function(header, position) {
    if (header) index[header] = position + 1;
  });
  return index;
}

function readSubmissionRow_(sheet, rowNumber) {
  const headers = readHeaders_(sheet);
  const values = sheet.getRange(rowNumber, 1, 1, headers.length).getDisplayValues()[0];
  const byHeader = {};
  headers.forEach(function(header, index) {
    if (header) byHeader[header] = values[index];
  });
  return {
    rowNumber: rowNumber,
    byHeader: byHeader,
    submission: normalizeSubmission_(byHeader)
  };
}

function getSystemValue_(rowData, header) {
  return stringValue_(rowData.byHeader[header]);
}

function writeSystemValues_(sheet, rowNumber, valuesByHeader) {
  const index = createHeaderIndex_(readHeaders_(sheet));
  Object.keys(valuesByHeader).forEach(function(header) {
    if (!index[header]) throw new Error('システム列が見つかりません: ' + header);
    sheet.getRange(rowNumber, index[header]).setValue(valuesByHeader[header]);
  });
}

function markRowPending_(sheet, rowNumber) {
  ensureSystemColumns_(sheet);
  writeSystemValues_(sheet, rowNumber, {
    '【システム】処理状態': ASTRA_CONFIG.STATUS.PENDING,
    '【システム】確認状態': ASTRA_CONFIG.REVIEW_STATUS.UNREVIEWED,
    '【システム】エラー内容': ''
  });
}

function findProcessableRows_(sheet, limit) {
  const headers = readHeaders_(sheet);
  const index = createHeaderIndex_(headers);
  const statusColumn = index['【システム】処理状態'];
  const processedAtColumn = index['【システム】最終処理日時'];
  const attemptColumn = index['【システム】試行回数'];
  if (!statusColumn || sheet.getLastRow() <= ASTRA_CONFIG.HEADER_ROW) return [];

  const rowCount = sheet.getLastRow() - ASTRA_CONFIG.HEADER_ROW;
  const statuses = sheet.getRange(ASTRA_CONFIG.HEADER_ROW + 1, statusColumn, rowCount).getDisplayValues();
  const processedAts = processedAtColumn
    ? sheet.getRange(ASTRA_CONFIG.HEADER_ROW + 1, processedAtColumn, rowCount).getValues()
    : [];
  const attempts = attemptColumn
    ? sheet.getRange(ASTRA_CONFIG.HEADER_ROW + 1, attemptColumn, rowCount).getDisplayValues()
    : [];
  const now = Date.now();
  const staleMillis = ASTRA_CONFIG.PROCESSING_TIMEOUT_MINUTES * 60 * 1000;
  const rows = [];

  for (let offset = 0; offset < rowCount && rows.length < limit; offset += 1) {
    const status = stringValue_(statuses[offset][0]);
    const lastProcessed = processedAts[offset] ? processedAts[offset][0] : null;
    const attemptCount = attempts[offset] ? Number(attempts[offset][0] || 0) : 0;
    const isStale = attemptCount < ASTRA_CONFIG.MAX_ATTEMPTS &&
      status === ASTRA_CONFIG.STATUS.PROCESSING &&
      lastProcessed instanceof Date &&
      now - lastProcessed.getTime() > staleMillis;
    const canRetryError = status === ASTRA_CONFIG.STATUS.ERROR && attemptCount < ASTRA_CONFIG.MAX_ATTEMPTS;
    if (!status || status === ASTRA_CONFIG.STATUS.PENDING || canRetryError || isStale) {
      rows.push(ASTRA_CONFIG.HEADER_ROW + 1 + offset);
    }
  }
  return rows;
}

function markExhaustedProcessingRows_(sheet) {
  const index = createHeaderIndex_(readHeaders_(sheet));
  const statusColumn = index['【システム】処理状態'];
  const processedAtColumn = index['【システム】最終処理日時'];
  const attemptColumn = index['【システム】試行回数'];
  const caseIdColumn = index['【システム】案件ID'];
  if (!statusColumn || !processedAtColumn || !attemptColumn) return;
  if (sheet.getLastRow() <= ASTRA_CONFIG.HEADER_ROW) return;

  const rowCount = sheet.getLastRow() - ASTRA_CONFIG.HEADER_ROW;
  const statuses = sheet.getRange(ASTRA_CONFIG.HEADER_ROW + 1, statusColumn, rowCount).getDisplayValues();
  const processedAts = sheet.getRange(ASTRA_CONFIG.HEADER_ROW + 1, processedAtColumn, rowCount).getValues();
  const attempts = sheet.getRange(ASTRA_CONFIG.HEADER_ROW + 1, attemptColumn, rowCount).getDisplayValues();
  const caseIds = caseIdColumn
    ? sheet.getRange(ASTRA_CONFIG.HEADER_ROW + 1, caseIdColumn, rowCount).getDisplayValues()
    : [];
  const now = Date.now();
  const staleMillis = ASTRA_CONFIG.PROCESSING_TIMEOUT_MINUTES * 60 * 1000;

  for (let offset = 0; offset < rowCount; offset += 1) {
    const status = stringValue_(statuses[offset][0]);
    if (status !== ASTRA_CONFIG.STATUS.PROCESSING) continue;
    const attemptCount = Number(attempts[offset][0] || 0);
    if (attemptCount < ASTRA_CONFIG.MAX_ATTEMPTS) continue;
    const lastProcessed = processedAts[offset][0];
    if (!(lastProcessed instanceof Date) || now - lastProcessed.getTime() <= staleMillis) continue;

    const rowNumber = ASTRA_CONFIG.HEADER_ROW + 1 + offset;
    const caseId = caseIds[offset] ? stringValue_(caseIds[offset][0]) : '';
    const message = '自動再試行の上限（' + ASTRA_CONFIG.MAX_ATTEMPTS + '回）に達したため処理を中断しました。処理ログを確認し、原因解消後にメニューから再生成してください。';
    writeSystemValues_(sheet, rowNumber, {
      '【システム】処理状態': ASTRA_CONFIG.STATUS.ERROR,
      '【システム】確認状態': ASTRA_CONFIG.REVIEW_STATUS.NEEDS_REVIEW,
      '【システム】最終処理日時': new Date(),
      '【システム】エラー内容': message
    });
    appendLog_('ERROR', caseId, rowNumber, '再試行上限', message);
    notifySystemError_(caseId, rowNumber, new Error(message));
  }
}

function markLegacyRowsSkipped_(sheet) {
  const index = createHeaderIndex_(readHeaders_(sheet));
  const statusColumn = index['【システム】処理状態'];
  if (!statusColumn || sheet.getLastRow() <= ASTRA_CONFIG.HEADER_ROW) return 0;

  const rowCount = sheet.getLastRow() - ASTRA_CONFIG.HEADER_ROW;
  const range = sheet.getRange(ASTRA_CONFIG.HEADER_ROW + 1, statusColumn, rowCount);
  const statuses = range.getDisplayValues();
  let marked = 0;
  statuses.forEach(function(row) {
    if (!stringValue_(row[0])) {
      row[0] = ASTRA_CONFIG.STATUS.SKIPPED;
      marked += 1;
    }
  });
  if (marked) range.setValues(statuses);
  return marked;
}

function ensureManagementSheets_() {
  const spreadsheet = getResponseSpreadsheet_();
  ensureSheetWithHeaders_(spreadsheet, ASTRA_CONFIG.SHEETS.LOG, [
    '日時', 'レベル', '案件ID', '回答行', '処理', 'メッセージ'
  ]);
  ensureSheetWithHeaders_(spreadsheet, ASTRA_CONFIG.SHEETS.DOCUMENTS, [
    '作成日時', '案件ID', '回答行', '書類種別', 'ファイルID', 'ファイル名', 'URL', 'PDF URL'
  ]);
  const settingsSheet = ensureSheetWithHeaders_(spreadsheet, ASTRA_CONFIG.SHEETS.SETTINGS, [
    '設定キー', '値', '説明'
  ]);
  writeSettingsSnapshot_(settingsSheet);
}

function ensureSheetWithHeaders_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0 || !stringValue_(sheet.getRange(1, 1).getDisplayValue())) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#d9ead3');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function writeSettingsSnapshot_(sheet) {
  const properties = PropertiesService.getScriptProperties().getProperties();
  const rows = [
    ['VERSION', ASTRA_CONFIG.VERSION, 'ローカルコードの版'],
    ['ROOT_FOLDER_ID', properties.ROOT_FOLDER_ID || '', '親フォルダ'],
    ['OUTPUT_FOLDER_ID', properties.OUTPUT_FOLDER_ID || '', '顧客別フォルダの保存先（必須）'],
    ['TEMPLATE_FOLDER_ID', properties.TEMPLATE_FOLDER_ID || '', '差込テンプレート保存先（必須）'],
    ['RESPONSE_SHEET_NAME', properties.RESPONSE_SHEET_NAME || ASTRA_CONFIG.RESPONSE_SHEET_NAME, '回答シート名（変更時のみ設定）'],
    ['TEMPLATE_CONFIRMATION_ID', properties.TEMPLATE_CONFIRMATION_ID || '', '申請内容確認書テンプレート'],
    ['TEMPLATE_REQUIREMENTS_ID', properties.TEMPLATE_REQUIREMENTS_ID || '', '必要書類リストテンプレート'],
    ['PDF_ENABLED', properties.PDF_ENABLED || 'false', 'true のとき生成時にPDFも作成（メニューから手動書き出しも可）'],
    ['NOTIFICATION_EMAIL', properties.NOTIFICATION_EMAIL || '', '処理エラー時の通知先メール（空なら通知しない）'],
    ['注意', '住所は自動変換しない', '顧客入力を暫定転記し、行政書士が必要に応じて修正・照合する'],
    ['住民票等', '提出は任意', '問い合わせ段階の入力負担を増やさない']
  ];
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  sheet.autoResizeColumns(1, 3);
}

function appendLog_(level, caseId, rowNumber, action, message) {
  const sheet = getResponseSpreadsheet_().getSheetByName(ASTRA_CONFIG.SHEETS.LOG);
  if (!sheet) return;
  sheet.appendRow([new Date(), level, caseId || '', rowNumber || '', action || '', stringValue_(message).slice(0, 1000)]);
}

function appendDocumentRecord_(record) {
  const sheet = getResponseSpreadsheet_().getSheetByName(ASTRA_CONFIG.SHEETS.DOCUMENTS);
  if (!sheet) return;
  sheet.appendRow([
    new Date(), record.caseId, record.rowNumber, record.type, record.fileId,
    record.name, record.url, record.pdfUrl || ''
  ]);
}
