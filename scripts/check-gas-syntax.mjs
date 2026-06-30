import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const gasDir = path.join(root, 'gas');
const files = fs.readdirSync(gasDir)
  .filter((name) => name.endsWith('.gs'))
  .sort();

for (const file of files) {
  const source = fs.readFileSync(path.join(gasDir, file), 'utf8');
  new vm.Script(source, { filename: file });
  console.log(`PASS ${file}`);
}
