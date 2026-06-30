function runAstraSelfTests() {
  const tests = [];
  tests.push(runTest_('ハイフン住所を要確認にする', function() {
    const warnings = detectAddressWarnings_('東京都新宿区西新宿1-2-1');
    assertTrue_(warnings.some(function(item) { return item.indexOf('ハイフン') !== -1; }));
  }));
  tests.push(runTest_('正式表記らしい住所はハイフン警告しない', function() {
    const warnings = detectAddressWarnings_('東京都新宿区西新宿一丁目2番地1号');
    assertTrue_(!warnings.some(function(item) { return item.indexOf('ハイフン') !== -1; }));
  }));
  tests.push(runTest_('業務名を正規化する', function() {
    assertEqual_('建設業許可', normalizeBusiness_('建設業許可の相談'));
    assertEqual_('風営法', normalizeBusiness_('深夜酒類'));
  }));
  tests.push(runTest_('同意なしをエラーにする', function() {
    const sample = createValidTestSubmission_();
    sample.consent = '';
    assertTrue_(validateSubmission_(sample).errors.length > 0);
  }));

  const failed = tests.filter(function(test) { return !test.ok; });
  const message = tests.map(function(test) {
    return (test.ok ? 'PASS' : 'FAIL') + ': ' + test.name + (test.error ? ' - ' + test.error : '');
  }).join('\n');
  Logger.log(message);
  if (failed.length) throw new Error(failed.length + '件のセルフテストが失敗しました。\n' + message);
  SpreadsheetApp.getUi().alert('セルフテスト成功', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

function runTest_(name, callback) {
  try {
    callback();
    return { name: name, ok: true, error: '' };
  } catch (error) {
    return { name: name, ok: false, error: error.message };
  }
}

function assertTrue_(condition) {
  if (!condition) throw new Error('条件が false です。');
}

function assertEqual_(expected, actual) {
  if (expected !== actual) throw new Error('expected=' + expected + ', actual=' + actual);
}

function createValidTestSubmission_() {
  return {
    timestamp: '2026/07/01 10:00:00',
    applicantType: '個人',
    name: '山田太郎',
    nameKana: 'ヤマダタロウ',
    representative: '',
    postalCode: '160-0023',
    addressOriginal: '東京都新宿区西新宿一丁目2番地1号',
    phone: '090-1234-5678',
    email: 'test@example.com',
    businessOriginal: '建設業許可',
    business: '建設業許可',
    details: '新規許可について相談したい。',
    consent: '同意する',
    notes: ''
  };
}
