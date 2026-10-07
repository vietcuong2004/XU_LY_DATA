import {test} from 'node:test';
import assert from 'node:assert/strict';
import {formulaBarText, meaningfulRowCount} from '../ui/modules/grid-view.mjs';

test('preview removes only trailing formatted rows without data', () => {
  const rows = [
    [{value: '2027/52', formula: null, comment: ''}],
    [{value: '2027/53', formula: null, comment: ''}],
    [{value: null, formula: '=A2', comment: ''}],
    [{value: null, formula: null, comment: ''}],
    [{value: '', formula: null, comment: ''}]
  ];
  assert.equal(meaningfulRowCount(rows), 3);
  assert.equal(meaningfulRowCount([[{value: null}]]), 1);
});

test('formula bar shows the real formula and regular cell values', () => {
  assert.equal(formulaBarText({value: 405.6, formula: '=SUM(B14:B130)'}), '=SUM(B14:B130)');
  assert.equal(formulaBarText({value: true, formula: null}), 'TRUE');
  assert.equal(formulaBarText({value: 405.6, formula: null}), '405.6');
  assert.equal(formulaBarText({value: 405.6, formula: null}, {raw: '404.8'}), '404.8');
  assert.equal(formulaBarText(null), '');
});
