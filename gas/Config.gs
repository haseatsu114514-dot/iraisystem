/**
 * アストラ書類自動作成システムの設定。
 * 本番値は Script Properties を優先し、DEFAULT_IDS は初回設定の既定値にだけ使う。
 */
const ASTRA_CONFIG = Object.freeze({
  VERSION: '0.1.0',
  TIME_ZONE: 'Asia/Tokyo',
  RESPONSE_SHEET_NAME: 'フォームの回答 1',
  HEADER_ROW: 1,
  MAX_ROWS_PER_RUN: 10,
  PROCESSING_TIMEOUT_MINUTES: 15,
  WORKER_INTERVAL_MINUTES: 5,
  DEFAULT_IDS: Object.freeze({
    ROOT_FOLDER_ID: '12UA4QRlVYNIm2XJ8TG5SkM56vYWGIj0G',
    RESPONSE_SPREADSHEET_ID: '14GZqV-axuZbNv54TOyv73P94gUIZ19D0uMYCQ7gb8AA',
    OUTPUT_FOLDER_ID: '18Fanr4EfWlWf8n6SC0UQjSc234CXms18',
    TEMPLATE_FOLDER_ID: '1hH-DwV9SR1-cfMig4xH6rPZpEXdsTBVE',
    GAS_FOLDER_ID: '1b8d470IAkndIyY5hgKhISIvvNbUbXhOP'
  }),
  PROPERTY_KEYS: Object.freeze({
    ROOT_FOLDER_ID: 'ROOT_FOLDER_ID',
    RESPONSE_SPREADSHEET_ID: 'RESPONSE_SPREADSHEET_ID',
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
    ERROR: 'エラー'
  }),
  REVIEW_STATUS: Object.freeze({
    UNREVIEWED: '未確認',
    NEEDS_REVIEW: '要確認',
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
  })
});

function setDefaultScriptProperties_() {
  const properties = PropertiesService.getScriptProperties();
  const current = properties.getProperties();
  const updates = {};

  Object.keys(ASTRA_CONFIG.DEFAULT_IDS).forEach(function(key) {
    if (!current[key]) updates[key] = ASTRA_CONFIG.DEFAULT_IDS[key];
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
