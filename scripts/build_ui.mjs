import {readFile, writeFile} from 'node:fs/promises';
import {fileURLToPath} from 'node:url';
import {dirname, join} from 'node:path';

const ui = join(dirname(dirname(fileURLToPath(import.meta.url))), 'ui');
const modules = ["state.mjs","api.mjs","file-list.mjs","grid-view.mjs","preview-view.mjs","workbook-controller.mjs","week-picker.mjs"];
const sources = [];
for (const name of modules) {
  sources.push((await readFile(join(ui, 'modules', name), 'utf8')).replace(/^export /gm, ''));
}
const entry = (await readFile(join(ui, 'app.entry.js'), 'utf8')).replace(/^import .*;\r?\n/gm, '');
await writeFile(join(ui, 'app.js'), "'use strict';\r\n" + sources.join('\r\n') + '\r\n' + entry);
