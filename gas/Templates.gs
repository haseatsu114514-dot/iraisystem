const ASTRA_TEMPLATE_DEFINITIONS = Object.freeze([
  Object.freeze({
    key: 'confirmation',
    type: '申請内容確認書',
    propertyKey: 'TEMPLATE_CONFIRMATION_ID',
    filePrefix: '申請内容確認書'
  }),
  Object.freeze({
    key: 'requirements',
    type: '必要書類リスト',
    propertyKey: 'TEMPLATE_REQUIREMENTS_ID',
    filePrefix: '必要書類リスト'
  })
]);

const ASTRA_REQUIREMENT_ITEMS = Object.freeze({
  '補助金業務': Object.freeze([
    '対象となる補助金名・公募回の確認',
    '法人・事業者の基本資料',
    '直近の決算・確定申告関係資料',
    '事業計画・導入予定経費の資料',
    '見積書・相見積もりの要否確認',
    '公募要領と最新様式の確認'
  ]),
  '建設業許可': Object.freeze([
    '申請区分（新規・更新・業種追加等）の確認',
    '法人・役員・営業所に関する資料',
    '経営業務管理責任者に関する確認資料',
    '専任技術者に関する確認資料',
    '財務・社会保険関係資料',
    '申請先自治体の最新様式確認'
  ]),
  '風営法': Object.freeze([
    '営業区分と管轄警察署の確認',
    '店舗・物件の使用権原に関する資料',
    '飲食店営業許可等の確認',
    '店舗平面図・求積図・設備図の確認',
    '管理者・誓約関係資料の確認',
    '都道府県警察の最新様式確認'
  ]),
  'その他': Object.freeze([
    '相談内容の具体化',
    '提出先・期限の確認',
    '本人・法人確認資料の要否確認',
    '正式な必要書類を行政書士が確定'
  ])
});

function ensureDefaultTemplates_() {
  const properties = PropertiesService.getScriptProperties();
  const templateFolder = DriveApp.getFolderById(
    getScriptProperty_(ASTRA_CONFIG.PROPERTY_KEYS.TEMPLATE_FOLDER_ID, true)
  );

  if (!properties.getProperty('TEMPLATE_CONFIRMATION_ID')) {
    const document = DocumentApp.create('TEMPLATE_申請内容確認書_共通');
    const body = document.getBody();
    body.appendParagraph('申請内容確認書').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('この書類は顧客入力を転記した下書きです。行政書士が原本と照合し、確認済みにしてから使用してください。');
    body.appendParagraph('案件ID: {{案件ID}}');
    body.appendParagraph('作成日: {{作成日}}');
    body.appendParagraph('申込区分: {{申込区分}}');
    body.appendParagraph('氏名または法人名: {{氏名または法人名}}');
    body.appendParagraph('フリガナ: {{フリガナ}}');
    body.appendParagraph('代表者名: {{代表者名}}');
    body.appendParagraph('郵便番号: {{郵便番号}}');
    body.appendParagraph('顧客入力住所（原文）: {{住所または本店所在地}}');
    body.appendParagraph('電話番号: {{電話番号}}');
    body.appendParagraph('メールアドレス: {{メールアドレス}}');
    body.appendParagraph('依頼したい業務: {{依頼したい業務}}');
    body.appendParagraph('依頼内容の詳細: {{依頼内容の詳細}}');
    body.appendParagraph('備考・連絡事項: {{備考・連絡事項}}');
    body.appendParagraph('自動確認事項').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('{{確認事項}}');
    body.appendParagraph('行政書士確認欄: □ 住所原本照合　□ 氏名照合　□ 依頼内容確認　□ 使用可');
    document.saveAndClose();
    const file = DriveApp.getFileById(document.getId());
    file.moveTo(templateFolder);
    properties.setProperty('TEMPLATE_CONFIRMATION_ID', document.getId());
  }

  if (!properties.getProperty('TEMPLATE_REQUIREMENTS_ID')) {
    const document = DocumentApp.create('TEMPLATE_必要書類リスト_共通');
    const body = document.getBody();
    body.appendParagraph('必要書類リスト（初回確認用）').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('案件ID: {{案件ID}}');
    body.appendParagraph('依頼者: {{氏名または法人名}}');
    body.appendParagraph('業務: {{依頼したい業務}}');
    body.appendParagraph('注意: この一覧は初回確認用です。正式な必要書類は、管轄・申請区分・最新様式を行政書士が確認して確定してください。');
    body.appendParagraph('{{必要書類リスト}}');
    body.appendParagraph('確認事項').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph('{{確認事項}}');
    document.saveAndClose();
    const file = DriveApp.getFileById(document.getId());
    file.moveTo(templateFolder);
    properties.setProperty('TEMPLATE_REQUIREMENTS_ID', document.getId());
  }
}

function getEnabledTemplates_() {
  const properties = PropertiesService.getScriptProperties();
  return ASTRA_TEMPLATE_DEFINITIONS.map(function(definition) {
    return {
      key: definition.key,
      type: definition.type,
      filePrefix: definition.filePrefix,
      templateId: properties.getProperty(definition.propertyKey) || ''
    };
  }).filter(function(template) {
    return Boolean(template.templateId);
  });
}

function getRequirementListText_(business) {
  const items = ASTRA_REQUIREMENT_ITEMS[business] || ASTRA_REQUIREMENT_ITEMS['その他'];
  return items.map(function(item) { return '□ ' + item; }).join('\n');
}

function buildTemplateData_(submission, caseId, validation) {
  const warnings = validation.warnings.length
    ? validation.warnings.map(function(item) { return '□ ' + item; }).join('\n')
    : '□ 自動検査では警告なし（最終確認は必須）';
  return {
    '案件ID': caseId,
    '作成日': Utilities.formatDate(new Date(), ASTRA_CONFIG.TIME_ZONE, 'yyyy年M月d日'),
    '申込区分': submission.applicantType,
    '氏名または法人名': submission.name,
    'フリガナ': submission.nameKana,
    '代表者名': submission.representative,
    '郵便番号': submission.postalCode,
    '住所または本店所在地': submission.addressOriginal,
    '電話番号': submission.phone,
    'メールアドレス': submission.email,
    '依頼したい業務': submission.business,
    '依頼内容の詳細': submission.details,
    '備考・連絡事項': submission.notes,
    '確認事項': warnings,
    '必要書類リスト': getRequirementListText_(submission.business)
  };
}
