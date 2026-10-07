export function createFileList({state, $, storage, normalize, escapeHTML}) {
function getSavedFileStates() {
  if (!state.job) return {};
  try {
    return JSON.parse(storage.getItem(`job_${state.job.id}_file_states`) || '{}');
  } catch (e) { return {}; }
}

function saveFileState(fileId, updates) {
  if (!state.job) return;
  const all = getSavedFileStates();
  const current = all[String(fileId)] || { viewed: false, downloaded: false, downloaded_rev: -1 };
  all[String(fileId)] = { ...current, ...updates };
  try {
    storage.setItem(`job_${state.job.id}_file_states`, JSON.stringify(all));
  } catch (e) { }
}

function getFileStatus(fileId, fileObj) {
  const all = getSavedFileStates();
  const st = all[String(fileId)] || { viewed: false, downloaded: false, downloaded_rev: -1 };
  const isEdited = (fileObj.changed_cells > 0) || (fileObj.edits && fileObj.edits.length > 0);
  const rev = fileObj.revision || 0;

  if (isEdited) {
    if (st.downloaded && st.downloaded_rev === rev) {
      return { key: 'downloaded', dotClass: 'dot-downloaded', label: 'Đã tải xuống', textClass: 'text-downloaded' };
    }
    return { key: 'edited', dotClass: 'dot-edited', label: 'Đã sửa', textClass: 'text-edited' };
  }
  if (st.downloaded) {
    return { key: 'downloaded', dotClass: 'dot-downloaded', label: 'Đã tải xuống', textClass: 'text-downloaded' };
  }
  if (st.viewed) {
    return { key: 'viewed', dotClass: 'dot-viewed', label: 'Đã xem', textClass: 'text-viewed' };
  }
  return { key: 'unviewed', dotClass: 'dot-unviewed', label: 'Chưa xem', textClass: 'text-unviewed' };
}

function fileMatchesFilter(fileId, fileObj, filter, search) {
  if (search && !normalize(fileObj.name || '').includes(normalize(search))) {
    return false;
  }
  if (filter === 'all') return true;
  const status = getFileStatus(fileId, fileObj);
  if (filter === 'warning') return Boolean(fileObj.problems && fileObj.problems.length);
  return status.key === filter;
}

function renderFamilies() {
  if (!state.report) return;
  const search = state.search;
  const filter = state.filter;

  const totalCount = state.report.families.length + (state.report.source_file ? 1 : 0);
  $('family-count').textContent = totalCount;

  let html = '';

  // 1. Source file at top
  const src = state.report.source_file;
  if (src && fileMatchesFilter(-1, src, filter, search)) {
    const srcStatus = getFileStatus(-1, src);
    const isSelected = state.file === -1;
    let name = src.original_name || src.file || src.name;
    if (!name || name.toLowerCase() === 'input.xlsx') {
      name = state.job?.source || state.config?.local_input || '2026 INTERNAL SCHEDULE FERRERO-WK40.xlsx';
    }
    html += `
      <div class="family-item source-item ${isSelected ? 'selected' : ''}" data-file="-1" role="button" tabindex="0" title="Nhấn để xem file nguồn gốc">
        <div class="family-item-content">
          <div class="family-item-name">
            <span class="status-dot ${srcStatus.dotClass}" title="${srcStatus.label}"></span>
            <span class="source-tag">Nguồn</span>
            <span class="name-text">${escapeHTML(name)}</span>
          </div>
          <small>Sửa nguồn · tính lại workbook và tạo lại mọi Family</small>
        </div>
        <div class="source-actions">
          <button type="button" class="btn-source-replace" data-choose-source title="Tải lên file Excel nguồn mới từ máy tính">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="17 8 12 3 7 8"></polyline>
              <line x1="12" y1="3" x2="12" y2="15"></line>
            </svg>
            <span>Tải lên file khác</span>
          </button>
        </div>
      </div>
    `;
  }

  // 2. Family items
  const matchedFamilies = state.report.families.filter(f => fileMatchesFilter(f.id, f, filter, search));
  if (matchedFamilies.length > 0) {
    if (src && fileMatchesFilter(-1, src, filter, search)) {
      html += `<div class="source-divider"><span>CÁC FILE FAMILY OUTPUT (${matchedFamilies.length})</span></div>`;
    }
    html += matchedFamilies.map(f => {
      const status = getFileStatus(f.id, f);
      const isSelected = f.id === state.file;
      return `
        <div class="family-item ${isSelected ? 'selected' : ''}" data-file="${f.id}" role="button" tabindex="0">
          <div class="family-item-content">
            <div class="family-item-name">
              <span class="status-dot ${status.dotClass}" title="${status.label}"></span>
              <span class="name-text">${escapeHTML(f.name)}</span>
            </div>
            <small>${f.markets} thị trường · ${f.items} item · <span class="status-text ${status.textClass}">${status.label}</span>${(status.key === 'unviewed' && f.problems && f.problems.length) ? ' · <b style="color:#d97706">cần xem</b>' : ''}</small>
          </div>
          <button type="button" class="btn-item-download" data-download-file="${f.id}" title="Tải xuống file ${escapeHTML(f.name)}.xlsx">
            <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
            <span>Tải xuống</span>
          </button>
        </div>
      `;
    }).join('');
  }

  if (!html) {
    html = '<div class="empty-state">Không tìm thấy file phù hợp.</div>';
  }

  $('family-list').innerHTML = html;
}

return {saveFileState, getFileStatus, renderFamilies};
}
