import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const context = vm.createContext({ console });
for (const file of ['Config.gs', 'Validation.gs', 'Templates.gs']) {
  const source = fs.readFileSync(path.join(root, 'gas', file), 'utf8');
  vm.runInContext(source, context, { filename: file });
}

function evaluate(expression) {
  return vm.runInContext(expression, context);
}

function validSubmission(overrides = {}) {
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
    consent: '個人情報の利用目的に同意します',
    notes: '',
    ...overrides
  };
}

test('ハイフン住所を勝手に変換せず警告する', () => {
  const warnings = evaluate(`detectAddressWarnings_('東京都新宿区西新宿1-2-1')`);
  assert.equal(warnings.length, 1);
  assert.match(warnings[0], /自動変換はしていません/);
});

test('丁目があっても残りがハイフンなら警告する', () => {
  const warnings = evaluate(`detectAddressWarnings_('東京都新宿区西新宿1丁目2-1')`);
  assert.ok(warnings.some((warning) => warning.includes('ハイフン')));
});

test('住民票表記を想定した住所にはハイフン警告を出さない', () => {
  const warnings = evaluate(`detectAddressWarnings_('東京都新宿区西新宿一丁目2番地1号')`);
  assert.equal(warnings.length, 0);
});

test('業務名の揺れを正規化する', () => {
  assert.equal(evaluate(`normalizeBusiness_('建設業許可の相談')`), '建設業許可');
  assert.equal(evaluate(`normalizeBusiness_('深夜酒類営業')`), '風営法');
  assert.equal(evaluate(`normalizeBusiness_('補助金について')`), '補助金業務');
});

test('同意文が長くても同意として扱う', () => {
  assert.equal(evaluate(`isAffirmative_('個人情報の利用目的に同意します')`), true);
  assert.equal(evaluate(`isAffirmative_('同意しません')`), false);
});

test('必須項目が揃った回答はエラーなし', () => {
  context.sample = validSubmission();
  const result = evaluate('validateSubmission_(sample)');
  assert.deepEqual([...result.errors], []);
});

test('同意なしは生成停止対象', () => {
  context.sample = validSubmission({ consent: '' });
  const result = evaluate('validateSubmission_(sample)');
  assert.ok(result.errors.some((error) => error.includes('同意')));
});

test('Drive名から危険文字を取り除く', () => {
  const name = evaluate(`sanitizeDriveName_('株式会社/A:B*C?')`);
  assert.equal(name, '株式会社_A_B_C_');
});
