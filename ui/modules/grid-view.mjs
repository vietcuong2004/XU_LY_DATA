export function colIndex(text) { return [...text].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1; }
export function coords(address) { const m = /^([A-Z]+)(\d+)$/.exec(address); return { col: colIndex(m[1]), row: Number(m[2]) - 1 }; }
export function meaningfulRowCount(rows) {
  let count = rows.length;
  while (count > 1 && rows[count - 1].every(cell =>
    (cell.value === null || cell.value === '') && !cell.formula && !cell.comment)) count--;
  return count;
}
export function formulaBarText(cell, pending) {
  if (!cell) return '';
  if (pending) return String(pending.raw ?? '');
  if (cell.formula) return String(cell.formula);
  if (cell.value === null || cell.value === undefined) return '';
  if (typeof cell.value === 'boolean') return cell.value ? 'TRUE' : 'FALSE';
  return String(cell.value);
}
export function createGridView({state, $, escapeHTML, cellDisplay, pendingEdit, updateEditToolbar, updateFormulaBar}) {
function renderGrid() {
  const data = state.preview, covered = new Set(), merges = new Map();
  for (const merge of data.merges) { const [start, end] = merge.split(':').map(coords); merges.set(`${start.row},${start.col}`, { rows: end.row - start.row + 1, cols: end.col - start.col + 1 }); for (let r = start.row; r <= end.row; r++)for (let c = start.col; c <= end.col; c++)if (r !== start.row || c !== start.col) covered.add(`${r},${c}`); }

  // Thu thập các ô có lỗi/cảnh báo trên sheet này để đánh dấu đỏ (Lỗi dữ liệu / Chỗ cần sửa)
  const problemCells = new Set(
    (data.file.problems || [])
      .filter(p => (p.sheet || '').trim() === (data.sheet || '').trim())
      .map(p => p.cell)
  );

  const customWidths = state.colWidths?.[state.file]?.[state.sheet] || {};

  let colgroupHtml = '<colgroup>';
  const idxWidth = customWidths['col_0'] || 38;
  colgroupHtml += `<col class="index-col" data-col-idx="0" style="width: ${idxWidth}px;">`;

  data.columns.forEach((c, i) => {
    const colIdx = i + 1;
    const defWidth = customWidths[`col_${colIdx}`] || (data.col_widths?.[c] || (i === 0 ? 140 : 100));
    colgroupHtml += `<col class="${i === 0 ? 'label-col' : 'value-col'}" data-col-idx="${colIdx}" style="width: ${defWidth}px;">`;
  });
  colgroupHtml += '</colgroup>';

  let theadHtml = '<thead><tr>';
  theadHtml += `<th class="th-index th-resizable" data-col-idx="0"><div class="col-resizer" data-col-idx="0" title="Kéo để đổi độ rộng"></div></th>`;
  data.columns.forEach((c, i) => {
    const colIdx = i + 1;
    theadHtml += `<th scope="col" class="th-col th-resizable" data-col-idx="${colIdx}"><span class="th-text">${c}</span><div class="col-resizer" data-col-idx="${colIdx}" title="Kéo để đổi độ rộng cột ${c}"></div></th>`;
  });
  theadHtml += '</tr></thead>';

  let html = `<table class="sheet-grid" aria-label="${escapeHTML(data.sheet)}">${colgroupHtml}${theadHtml}<tbody>`;
  const visibleRows = data.rows.slice(0, meaningfulRowCount(data.rows));
  visibleRows.forEach((row, r) => {
    html += `<tr><td class="row-number">${r + 1}</td>`;
    row.forEach((cell, c) => {
      if (covered.has(`${r},${c}`)) return;
      const merge = merges.get(`${r},${c}`);
      const header = (state.sheet === 0 && r >= 1 && r < 8) || (state.sheet === 1 && r >= 8 && r < 13);
      const isProblem = cell.error || problemCells.has(cell.address);
      const pending = pendingEdit(state.sheet, cell.address);
      const isChanged = cell.changed;
      const classes = [
        'data-cell',
        c === 0 ? 'label-cell' : '',
        header ? 'header-cell' : '',
        typeof cell.value === 'number' ? 'numeric' : '',
        isChanged ? 'changed' : '',
        pending ? 'draft-cell' : '',
        isProblem ? 'error-cell' : '',
        cell.formula ? 'has-formula' : '',
        cell.editable ? 'is-editable' : ''
      ];
      const text = pending ? pending.raw : cell.value === null ? '' : cellDisplay(cell);
      // Màn hình chỉnh sửa chính không áp dụng màu sắc trang trí của Excel.
      // Chỉ phân biệt theo trạng thái: Khớp nguồn (bình thường), Đã thay đổi (xanh), Lỗi/cần sửa (đỏ).
      const titleHint = isProblem ? 'Lỗi/cần kiểm tra · Nhấn để sửa' : isChanged ? 'Đã chỉnh sửa · Nhấn để xem/sửa tiếp' : cell.editable ? 'Nhấn để sửa' : cell.formula || 'Khớp nguồn';
      html += `<td tabindex="0" role="button" data-address="${cell.address}" aria-label="${cell.address}: ${escapeHTML(text || 'trống')}" class="${classes.filter(Boolean).join(' ')}" ${merge ? `rowspan="${merge.rows}" colspan="${merge.cols}"` : ''} title="${escapeHTML(cell.address + ' · ' + titleHint)}">${escapeHTML(text)}</td>`;
    });
    html += '</tr>';
  });
  $('grid-container').innerHTML = html + '</tbody></table>';
  $('grid-info').textContent = `${visibleRows.length} dòng có dữ liệu · ${data.columns.length} cột · Đã lưu${data.revision ? ' · Lần sửa ' + data.revision : ''}`;
  state.cell = null;
  updateFormulaBar(null);
  updateEditToolbar();
}

return {renderGrid};
}
