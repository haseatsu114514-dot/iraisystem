/**
 * 回答値を正規化する。住所の原文は絶対に書き換えない。
 */
function normalizeSubmission_(rowValues) {
  const headers = ASTRA_CONFIG.SOURCE_HEADERS;
  return {
    timestamp: stringValue_(rowValues[headers.TIMESTAMP]),
    applicantType: stringValue_(rowValues[headers.APPLICANT_TYPE]),
    name: stringValue_(rowValues[headers.NAME]),
    nameKana: stringValue_(rowValues[headers.NAME_KANA]),
    representative: stringValue_(rowValues[headers.REPRESENTATIVE]),
    postalCode: stringValue_(rowValues[headers.POSTAL_CODE]),
    addressOriginal: stringValue_(rowValues[headers.ADDRESS]),
    phone: stringValue_(rowValues[headers.PHONE]),
    email: stringValue_(rowValues[headers.EMAIL]),
    businessOriginal: stringValue_(rowValues[headers.BUSINESS]),
    business: normalizeBusiness_(rowValues[headers.BUSINESS]),
    details: stringValue_(rowValues[headers.DETAILS]),
    consent: stringValue_(rowValues[headers.CONSENT]),
    notes: stringValue_(rowValues[headers.NOTES])
  };
}

function validateSubmission_(submission) {
  const errors = [];
  const warnings = [];

  if (!submission.applicantType) errors.push('申込区分が未入力です。');
  if (!submission.name) errors.push('氏名または法人名が未入力です。');
  if (!submission.addressOriginal) errors.push('住所または本店所在地が未入力です。');
  if (!submission.phone) errors.push('電話番号が未入力です。');
  if (!submission.email) errors.push('メールアドレスが未入力です。');
  if (!submission.business) errors.push('依頼したい業務を判定できません。');
  if (!isAffirmative_(submission.consent)) {
    errors.push('個人情報の利用目的への同意を確認できません。');
  }

  if (submission.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(submission.email)) {
    warnings.push('メールアドレスの形式を確認してください。');
  }

  const postalDigits = normalizeDigits_(submission.postalCode).replace(/[^0-9]/g, '');
  if (submission.postalCode && postalDigits.length !== 7) {
    warnings.push('郵便番号が7桁ではありません。');
  }

  if (submission.phone) {
    const phoneDigits = normalizeDigits_(submission.phone).replace(/[^0-9]/g, '');
    if (phoneDigits.length < 10 || phoneDigits.length > 11) {
      warnings.push('電話番号の桁数を確認してください。');
    }
  }

  detectAddressWarnings_(submission.addressOriginal).forEach(function(message) {
    warnings.push(message);
  });

  if (/法人|会社|個人事業主/.test(submission.applicantType) && !submission.representative) {
    warnings.push('法人・個人事業主の場合の代表者名が未入力です。');
  }

  if (!submission.details) {
    warnings.push('依頼内容の詳細が未入力です。初回確認時に聞き取りが必要です。');
  }

  return { errors: uniqueStrings_(errors), warnings: uniqueStrings_(warnings) };
}

function detectAddressWarnings_(address) {
  const value = stringValue_(address);
  if (!value) return [];

  const warnings = [];
  const normalized = normalizeDigits_(value);
  const hasHyphenNumberSequence = /\d+[\-－ー―‐‑–−]\d+/.test(normalized);
  if (hasHyphenNumberSequence) {
    warnings.push('住所がハイフン省略表記の可能性があります。住民票・登記簿どおりの表記を確認してください（自動変換はしていません）。');
  }
  if (!/[都道府県]/.test(value)) {
    warnings.push('住所に都道府県が見当たりません。省略されていないか確認してください。');
  }
  if (/[\-－ー―‐‑–−]$/.test(value)) {
    warnings.push('住所がハイフンで終わっています。入力途中でないか確認してください。');
  }
  return warnings;
}

function normalizeBusiness_(value) {
  const text = stringValue_(value).replace(/\s/g, '');
  if (!text) return '';
  if (/補助金/.test(text)) return '補助金業務';
  if (/建設業/.test(text)) return '建設業許可';
  if (/風営|深夜酒類/.test(text)) return '風営法';
  if (/その他/.test(text)) return 'その他';
  return '';
}

function isAffirmative_(value) {
  const text = stringValue_(value);
  if (/同意しない|同意しません|いいえ|false|^0$/i.test(text)) return false;
  return /同意|はい|確認しました|true|^1$/i.test(text);
}

function normalizeDigits_(value) {
  return stringValue_(value).replace(/[０-９]/g, function(character) {
    return String.fromCharCode(character.charCodeAt(0) - 0xFEE0);
  });
}

function stringValue_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function uniqueStrings_(items) {
  return items.filter(function(item, index, source) {
    return item && source.indexOf(item) === index;
  });
}

function sanitizeDriveName_(value) {
  const sanitized = stringValue_(value)
    .replace(/[\\/:*?"<>|\r\n\t]/g, '_')
    .replace(/\s+/g, ' ')
    .replace(/_+/g, '_')
    .trim();
  return (sanitized || '名称未設定').slice(0, 80);
}
