import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const releaseRoot = path.join(root, 'release', 'astra-gas-mvp-0.1.0');
const gasOutput = path.join(releaseRoot, 'gas');
const docsOutput = path.join(releaseRoot, 'docs');
const orderedGasFiles = [
  'Config.gs',
  'Validation.gs',
  'SheetRepository.gs',
  'FolderManager.gs',
  'Templates.gs',
  'DocumentGenerator.gs',
  'Triggers.gs',
  'Code.gs',
  'Tests.gs'
];

await fs.rm(releaseRoot, { recursive: true, force: true });
await fs.mkdir(gasOutput, { recursive: true });
await fs.mkdir(docsOutput, { recursive: true });

const bundleParts = [];
for (const file of orderedGasFiles) {
  const sourcePath = path.join(root, 'gas', file);
  const source = await fs.readFile(sourcePath, 'utf8');
  await fs.copyFile(sourcePath, path.join(gasOutput, file));
  bundleParts.push(`// ===== ${file} =====\n${source.trim()}\n`);
}

await fs.writeFile(
  path.join(releaseRoot, 'AstraSystem.gs'),
  bundleParts.join('\n'),
  'utf8'
);
await fs.copyFile(path.join(root, 'gas', 'appsscript.json'), path.join(releaseRoot, 'appsscript.json'));
await fs.copyFile(path.join(root, 'README.md'), path.join(releaseRoot, 'README.md'));
await fs.copyFile(
  path.join(root, 'docs', 'EXISTING_SYSTEM_REVIEW.md'),
  path.join(docsOutput, 'EXISTING_SYSTEM_REVIEW.md')
);
await fs.copyFile(
  path.join(root, 'docs', 'ACCEPTANCE_CHECKLIST.md'),
  path.join(docsOutput, 'ACCEPTANCE_CHECKLIST.md')
);

console.log(releaseRoot);
