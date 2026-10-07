export function createPreviewView({state, $, fmt, escapeHTML, renderGrid, renderProblems, renderAudit, focusCell, openCell, toast}) {
  const setBusy = busy => {
    $('grid-container').setAttribute('aria-busy', String(busy));
    if (busy) $('download-one').hidden = true;
    if ($('btn-preview-excel')) $('btn-preview-excel').disabled = busy;
  };
  return {
    setBusy,
    loading() { setBusy(true); $('grid-container').innerHTML = '<div class="empty-state">Đang mở bảng dữ liệu…</div>'; },
    error(error) { $('grid-container').textContent = error.message; $('download-one').hidden = true; toast(error.message, true); },
    show(result, selection, focusAddress) {
    $('download-one').hidden = false;
    let displayName = result.file.is_source ? (result.file.original_name || result.file.file || result.file.name) : result.file.name;
    if (result.file.is_source && (!displayName || displayName.toLowerCase() === 'input.xlsx')) {
      displayName = state.job?.source || state.config?.local_input || '2026 INTERNAL SCHEDULE FERRERO-WK40.xlsx';
    }
    $('workbook-name').textContent = displayName;
    if (result.file.is_source) {
      $('workbook-meta').textContent = `${result.sheets.length} sheet · ${result.source_editing?.message || 'File nguồn'} · Tải file này nhận nguồn hiện tại.`;
    } else {
      $('workbook-meta').textContent = `${result.file.items} item · ${result.file.markets} thị trường · ${result.file.source_checks ? fmt.format(result.file.source_checks) + ' đối chiếu nguồn' : 'Lịch nguồn: ' + state.report.first_week + ' → ' + state.report.last_week}`;
    }
    $('download-one').href = `/api/jobs/${selection.job}/download?file=${selection.file}`;
    $('download-one').setAttribute('download', displayName);
    if ($('excel-preview-download')) {
      $('excel-preview-download').href = $('download-one').href;
      $('excel-preview-download').setAttribute('download', displayName);
    }
    $('sheet-tabs').innerHTML = result.sheets.map((name, index) => `<button role="tab" aria-selected="${index === selection.sheet}" class="sheet-tab ${index === selection.sheet ? 'selected' : ''}" data-sheet="${index}">${index === 0 ? '▦ ' : index === 1 ? '▤ ' : '▥ '}${escapeHTML(name)}</button>`).join('');
    renderGrid(); renderProblems(); renderAudit();
    if (focusAddress) { focusCell(focusAddress); openCell(focusAddress); }

    }
  };
}
