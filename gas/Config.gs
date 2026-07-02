/**
 * アストラ書類自動作成システムの設定。
 * フォルダ等のIDはリポジトリへ書かず、Script Properties で設定する。
 * DEFAULT_IDS は導入先ごとに埋める場合の初期値枠（空のままでよい）。
 */
const ASTRA_CONFIG = Object.freeze({
  VERSION: '0.2.0',
  TIME_ZONE: 'Asia/Tokyo',
  RESPONSE_SHEET_NAME: 'フォームの回答 1',
  HEADER_ROW: 1,
  MAX_ROWS_PER_RUN: 10,
  MAX_ATTEMPTS: 3,
  PROCESSING_TIMEOUT_MINUTES: 15,
  WORKER_INTERVAL_MINUTES: 5,
  DEFAULT_IDS: Object.freeze({
    ROOT_FOLDER_ID: '',
    RESPONSE_SPREADSHEET_ID: '',
    OUTPUT_FOLDER_ID: '',
    TEMPLATE_FOLDER_ID: '',
    GAS_FOLDER_ID: ''
  }),
  PROPERTY_KEYS: Object.freeze({
    ROOT_FOLDER_ID: 'ROOT_FOLDER_ID',
    RESPONSE_SPREADSHEET_ID: 'RESPONSE_SPREADSHEET_ID',
    RESPONSE_SHEET_NAME: 'RESPONSE_SHEET_NAME',
    OUTPUT_FOLDER_ID: 'OUTPUT_FOLDER_ID',
    TEMPLATE_FOLDER_ID: 'TEMPLATE_FOLDER_ID',
    GAS_FOLDER_ID: 'GAS_FOLDER_ID',
    TEMPLATE_CONFIRMATION_ID: 'TEMPLATE_CONFIRMATION_ID',
    TEMPLATE_REQUIREMENTS_ID: 'TEMPLATE_REQUIREMENTS_ID',
    PDF_ENABLED: 'PDF_ENABLED',
    NOTIFICATION_EMAIL: 'NOTIFICATION_EMAIL'
  }),
  STATUS: Object.freeze({
    PENDING: '未処理',
    PROCESSING: '処理中',
    GENERATED: '生成済み',
    NEEDS_INPUT: '入力不備',
    ERROR: 'エラー',
    SKIPPED: '対象外（導入前）'
  }),
  REVIEW_STATUS: Object.freeze({
    UNREVIEWED: '未確認',
    NEEDS_REVIEW: '要確認',
    CONFIRMED: '確認済み'
  }),
  ADDRESS_REVIEW_STATUS: Object.freeze({
    UNREVIEWED: '未確認',
    NO_ATTACHMENT: '住民票等未提出',
    NEEDS_COMPARISON: '要照合',
    CONFIRMED: '確認済み'
  }),
  SHEETS: Object.freeze({
    LOG: '処理ログ',
    DOCUMENTS: '生成書類管理',
    SETTINGS: 'システム設定'
  }),
  SYSTEM_HEADERS: Object.freeze([
    '【システム】処理状態',
    '【システム】案件ID',
    '【システム】確認状態',
    '【システム】確認事項',
    '【システム】書類転記住所',
    '【システム】住所確認状態',
    '【システム】顧客フォルダID',
    '【システム】顧客フォルダURL',
    '【システム】生成書類URL',
    '【システム】試行回数',
    '【システム】最終処理日時',
    '【システム】エラー内容'
  ]),
  SOURCE_HEADERS: Object.freeze({
    TIMESTAMP: 'タイムスタンプ',
    APPLICANT_TYPE: '申込区分',
    NAME: '氏名または法人名',
    NAME_KANA: '氏名または法人名のフリガナ',
    REPRESENTATIVE: '代表者名（法人・個人事業主の場合）',
    POSTAL_CODE: '郵便番号',
    ADDRESS: '住所または本店所在地',
    PHONE: '電話番号',
    EMAIL: 'メールアドレス',
    BUSINESS: '依頼したい業務',
    DETAILS: '依頼内容の詳細',
    CONSENT: '個人情報の利用目的への同意',
    NOTES: '備考・連絡事項'
  }),
  OPTIONAL_SOURCE_HEADERS: Object.freeze({
    RESIDENT_RECORD_ATTACHMENTS: Object.freeze([
      '住民票等の画像（任意）',
      '住民票画像（任意）',
      '住民票・登記簿等（任意）'
    ])
  })
});

function setDefaultScriptProperties_() {
  const properties = PropertiesService.getScriptProperties();
  const current = properties.getProperties();
  const updates = {};

  Object.keys(ASTRA_CONFIG.DEFAULT_IDS).forEach(function(key) {
    if (!current[key] && ASTRA_CONFIG.DEFAULT_IDS[key]) {
      updates[key] = ASTRA_CONFIG.DEFAULT_IDS[key];
    }
  });

  if (!current.PDF_ENABLED) updates.PDF_ENABLED = 'false';
  if (Object.keys(updates).length) properties.setProperties(updates, false);
  return properties.getProperties();
}

function getScriptProperty_(key, required) {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (required && !value) {
    throw new Error('Script Properties に ' + key + ' が設定されていません。');
  }
  return value || '';
}

function isPdfEnabled_() {
  return getScriptProperty_(ASTRA_CONFIG.PROPERTY_KEYS.PDF_ENABLED, false) === 'true';
}

function getResponseSheetName_() {
  return getScriptProperty_(ASTRA_CONFIG.PROPERTY_KEYS.RESPONSE_SHEET_NAME, false) ||
    ASTRA_CONFIG.RESPONSE_SHEET_NAME;
}
