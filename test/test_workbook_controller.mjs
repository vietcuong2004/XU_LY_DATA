import {test} from 'node:test';
import assert from 'node:assert/strict';
import {createWorkbookController} from '../ui/modules/workbook-controller.mjs';
import {createWorkspaceState} from '../ui/modules/state.mjs';

function harness() {
  const state = createWorkspaceState();
  state.job = {id: 'job'};
  state.report = {source_file: {id: -1}, families: [{id: 0}, {id: 1}, {id: 2}]};
  const pending = [], shown = [], errors = [], viewed = [];
  const controller = createWorkbookController({state,
    api(url, data, options) { return new Promise((resolve, reject) => pending.push({url, ...options, resolve, reject})); },
    view: {loading() {}, setBusy() {}, show: result => shown.push(result), error: error => errors.push(error)},
    markViewed: id => viewed.push(id), renderFiles() {}
  });
  return {controller, state, pending, shown, errors, viewed};
}

test('late preview cannot replace the most recently selected family', async () => {
  const h = harness();
  const first = h.controller.selectFile(1);
  const second = h.controller.selectFile(2);
  assert.equal(h.pending[0].signal.aborted, true);
  h.pending[1].resolve({name: 'second'});
  await second;
  h.pending[0].resolve({name: 'first'});
  await first;
  assert.deepEqual(h.shown, [{name: 'second'}]);
  assert.deepEqual(h.viewed, [2]);
  assert.equal(h.state.file, 2);
});

test('late errors are ignored and failed current request can be retried', async () => {
  const h = harness();
  const first = h.controller.selectFile(1);
  const second = h.controller.selectFile(-1);
  h.pending[0].reject(new Error('old request'));
  await first;
  assert.equal(h.errors.length, 0);
  h.pending[1].reject(new Error('current request'));
  await second;
  assert.equal(h.errors.length, 1);
  const retry = h.controller.selectFile(-1);
  assert.match(h.pending[2].url, /file=-1&sheet=3$/);
  h.pending[2].resolve({name: 'source'});
  await retry;
  assert.equal(h.state.preview.name, 'source');
});

test('switching sheets invalidates the previous sheet request', async () => {
  const h = harness();
  const first = h.controller.selectSheet(1);
  const second = h.controller.selectSheet(2);
  h.pending[1].resolve({sheet: 2}); await second;
  h.pending[0].resolve({sheet: 1}); await first;
  assert.equal(h.state.preview.sheet, 2);
});

test('selecting the source opens SUM but an explicit Capsule tab remains selectable', async () => {
  const h = harness();
  const opening = h.controller.selectFile(-1);
  assert.match(h.pending[0].url, /file=-1&sheet=3$/);
  h.pending[0].resolve({file: {is_source: true}, sheets: ['Capsule', 'Injection', 'Production', 'SUM'], sheet: 'SUM'});
  await opening;
  assert.equal(h.state.sheet, 3);
  assert.equal(h.state.preview.sheet, 'SUM');

  const capsule = h.controller.selectSheet(0);
  h.pending[1].resolve({file: {is_source: true}, sheets: ['Capsule', 'Injection', 'Production', 'SUM'], sheet: 'Capsule'});
  await capsule;
  assert.equal(h.state.sheet, 0);
  assert.equal(h.state.preview.sheet, 'Capsule');
});

class Element extends EventTarget {}
function eventFor(target) {
  const event = new Event('click');
  Object.defineProperty(event, 'target', {value: target});
  return event;
}
test('rebinding attaches one listener; downloads do not switch file', async () => {
  const h = harness(), files = new Element(), sheets = new Element();
  const downloads = [];
  const options = {files, sheets, chooseSource() {}, download: id => downloads.push(id)};
  h.controller.bind(options); h.controller.bind(options);
  files.dispatchEvent(eventFor({closest: selector => selector === '[data-file]' ? {dataset: {file: '1'}} : null}));
  assert.equal(h.pending.length, 1);
  h.pending[0].resolve({name: 'one'});
  await Promise.resolve();
  files.dispatchEvent(eventFor({closest: selector => selector === '[data-download-file]' ? {dataset: {downloadFile: '2'}} : null}));
  assert.deepEqual(downloads, [2]);
  assert.equal(h.pending.length, 1);
  assert.equal(h.state.file, 1);
});

test('workspace instances do not share column widths or selection', () => {
  const a = createWorkspaceState(), b = createWorkspaceState();
  a.colWidths[0] = 123; a.file = 9;
  assert.deepEqual(b.colWidths, {});
  assert.equal(b.file, 0);
});
