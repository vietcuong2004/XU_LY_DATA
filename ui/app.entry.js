import {createWorkspaceState} from './modules/state.mjs';
import {createWeekPicker} from './modules/week-picker.mjs';
import {api} from './modules/api.mjs';
import {createFileList} from './modules/file-list.mjs';
import {createGridView, coords, colIndex, formulaBarText} from './modules/grid-view.mjs';
import {createPreviewView} from './modules/preview-view.mjs';
import {createWorkbookController} from './modules/workbook-controller.mjs';
const workbookDrafts = new Map();
const $ = id => document.getElementById(id);
const state = createWorkspaceState();
var previewModalSheet = 0;
const fmt = new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 6 });
const escapeHTML = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const display = value => value === null || value === undefined || value === '' ? '—' : typeof value === 'number' ? fmt.format(value) : String(value);
const cellDisplay = (cell, value = cell.value) => typeof value === 'number' && cell.format === 'General' ? String(value) : display(value);
const statusNames = { warning: 'Cần kiểm tra', edited: 'Đã chỉnh sửa', matched: 'Khớp nguồn' };
const badge = status => `<span class="badge ${status}">${statusNames[status]}</span>`;
const normalize = text => String(text).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

const {saveFileState, getFileStatus, renderFamilies} = createFileList({
  state, $, storage: {getItem: key => localStorage.getItem(key), setItem: (key, value) => localStorage.setItem(key, value)},
  normalize, escapeHTML
});
const {renderGrid} = createGridView({state, $, escapeHTML, cellDisplay, pendingEdit, updateEditToolbar, updateFormulaBar});
const previewView = createPreviewView({
  state, $, fmt, escapeHTML, renderGrid, renderProblems, renderAudit, focusCell, openCell, toast
});
const workbookController = createWorkbookController({
  state, api, view: previewView, markViewed: saveFileState, renderFiles: renderFamilies
});
const {loadPreview} = workbookController;
const weekPicker = createWeekPicker({$, document, normalize});
const {parseWeekInput, populateWeekOptions, resetWeekOptions} = weekPicker;

let toastTimer;
function toast(message, error = false) { $('toast').textContent = message; $('toast').className = 'show' + (error ? ' error' : ''); clearTimeout(toastTimer); toastTimer = setTimeout(() => $('toast').className = '', 5500); }

function formatDateTime(isoString) {
  if (!isoString) return '';
  try {
    const d = new Date(isoString);
    if (isNaN(d.getTime())) return String(isoString);
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    return `${day}/${month}/${year} ${hours}:${minutes}`;
  } catch (e) {
    return String(isoString);
  }
}

function formatFileSize(bytes) {
  if (!bytes || bytes <= 0) return '12.4 MB';
  const mb = bytes / (1024 * 1024);
  return `${mb.toFixed(1)} MB`;
}

function updateUploadStateUI() {
  const hasJob = Boolean(state.job && state.report);
  const cardGuide = $('card-guide');
  const cardRecent = $('card-recent');

  if (cardGuide) cardGuide.hidden = hasJob;
  if (cardRecent) cardRecent.hidden = !hasJob;

  if (hasJob) {
    const job = state.job;
    const report = state.report;
    const src = report.source_file;
    let fileName = src?.original_name || job?.source || src?.name || '2026 INTERNAL SCHEDULE FERRERO-WK40.xlsx';
    if (fileName && fileName.toLowerCase() === 'input.xlsx') {
      fileName = (job?.source && job.source.toLowerCase() !== 'input.xlsx') ? job.source : (state.config?.local_input || '2026 INTERNAL SCHEDULE FERRERO-WK40.xlsx');
    }
    setFileSelectedUI(fileName, src?.size);

    const famCount = report.families ? report.families.length : (job.total || 26);
    const badge = $('tab-review-badge');
    if (badge) {
      badge.textContent = `${famCount} Family`;
      badge.hidden = false;
    }
  } else {
    const badge = $('tab-review-badge');
    if (badge) badge.hidden = true;
  }
}

function switchTab(tabId) {
  const isUpload = tabId === 'upload';
  $('tab-btn-upload').classList.toggle('active', isUpload);
  $('tab-btn-upload').setAttribute('aria-selected', isUpload ? 'true' : 'false');
  $('tab-btn-review').classList.toggle('active', !isUpload);
  $('tab-btn-review').setAttribute('aria-selected', !isUpload ? 'true' : 'false');

  $('tab-content-upload').hidden = !isUpload;
  $('tab-content-review').hidden = isUpload;

  if (!isUpload) {
    const hasReport = Boolean(state.report);
    $('view-review').hidden = !hasReport;
    $('results-empty').hidden = hasReport;
  }
}

$('tab-btn-upload').addEventListener('click', () => switchTab('upload'));
$('tab-btn-review').addEventListener('click', () => switchTab('review'));
const btnGoto = $('btn-goto-upload');
if (btnGoto) btnGoto.addEventListener('click', () => switchTab('upload'));

function setExportState(enabled, href = '') {
  const btn = $('go-export');
  if (!btn) return;
  if (enabled && href) {
    btn.removeAttribute('aria-disabled');
    btn.href = href;
    btn.title = 'Tải toàn bộ kết quả đã lưu (.zip)';
  } else {
    btn.setAttribute('aria-disabled', 'true');
    btn.removeAttribute('href');
    btn.title = 'Chưa có file family nào được tạo để xuất';
  }
}
const btnExport = $('go-export');
if (btnExport) {
  btnExport.addEventListener('click', e => {
    if (btnExport.getAttribute('aria-disabled') === 'true' || !btnExport.getAttribute('href')) {
      e.preventDefault();
      e.stopPropagation();
      toast('Chưa hoàn tất tạo các file Family để xuất kết quả.', true);
      return;
    }
    if (state.report) {
      if (state.report.source_file) {
        saveFileState(-1, { downloaded: true, downloaded_rev: state.report.source_file.revision || 0 });
      }
      if (state.report.families) {
        state.report.families.forEach(f => {
          saveFileState(f.id, { downloaded: true, downloaded_rev: f.revision || 0 });
        });
      }
      renderFamilies();
    }
  });
}
setExportState(false);

// Expand the actual editor, retaining sheet tabs, cell dialogs and save events.
function setEditorFullscreen(expanded) {
  document.querySelector('.workbook-panel').classList.toggle('editor-fullscreen', expanded);
  document.body.classList.toggle('editor-expanded', expanded);
  $('fullscreen-editor').setAttribute('aria-pressed', String(expanded));
  $('fullscreen-editor').textContent = expanded ? '⛶ Thu nhỏ' : '⛶ Toàn màn hình';
}
$('fullscreen-editor').addEventListener('click', () => setEditorFullscreen(
  $('fullscreen-editor').getAttribute('aria-pressed') !== 'true'));
document.addEventListener('keydown', e => {
  if (e.key === 'Escape' && !document.querySelector('dialog[open]')) setEditorFullscreen(false);
});

function showView(view) {
  if (view !== 'progress') stopProgressClock();
  state.view = view;
  const busy = view === 'progress';
  if ($('view-upload')) $('view-upload').hidden = false;
  if ($('tab-content-upload')) $('tab-content-upload').hidden = false;
  $('view-progress').hidden = !busy;
  const card = $('card-unloaded') || $('card-loaded') || $('step-upload-card') || $('step-week-card') || document.querySelector('.studio-card');
  if (card) card.inert = busy;
  const hist = $('history');
  if (hist) hist.inert = busy;
  const hasFile = Boolean(state.input || state.useLocal || (state.job && state.report));
  if ($('process-button')) $('process-button').disabled = busy || !hasFile;
  if ($('clear-input')) $('clear-input').disabled = busy;
  if ($('shipping-week')) $('shipping-week').disabled = busy || !hasFile;
  if ($('shipping-year')) $('shipping-year').disabled = busy || !hasFile;
  if ($('btn-toggle-week')) $('btn-toggle-week').disabled = busy || !hasFile;
  if (view === 'review') {
    switchTab('review');
  } else if (view === 'progress' || view === 'upload') {
    switchTab('upload');
  }
}

function getCurrentUploadTimeStr() {
  const now = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`;
}

function setFileSelectedUI(fileName, fileSize) {
  if ($('drop-input')) $('drop-input').hidden = true;
  if ($('selected-file-display')) $('selected-file-display').hidden = false;
  if ($('selected-file-name')) $('selected-file-name').textContent = fileName;

  const sizeStr = fileSize ? formatFileSize(fileSize) : '18.6 MB';
  const timeStr = getCurrentUploadTimeStr();
  if ($('selected-file-meta')) {
    $('selected-file-meta').textContent = `Dung lượng: ${sizeStr} | Ngày tải lên: ${timeStr}`;
  }

  const wkMatch = fileName.match(/(?:WK|Wk|wk)[-_ ]?(\d+)/i);
  const yearMatch = fileName.match(/20\d{2}/);
  const yr = yearMatch ? parseInt(yearMatch[0], 10) : (new Date().getFullYear() || 2026);
  const wk = wkMatch && Number(wkMatch[1]) >= 1 && Number(wkMatch[1]) <= 53 ? Number(wkMatch[1]) : 40;

  populateWeekOptions(wk, yr);

  if ($('process-button')) $('process-button').disabled = false;
}

function resetFileSelection() {
  state.input = null;
  state.useLocal = false;
  if ($('input-file')) $('input-file').value = '';
  if ($('drop-input')) $('drop-input').hidden = false;
  if ($('selected-file-display')) $('selected-file-display').hidden = true;

  resetWeekOptions();
  if ($('process-button')) $('process-button').disabled = true;
}

function outputOptions(fileName) {
  const inpVal = $('shipping-week')?.value;
  const yearSelVal = $('shipping-year')?.value;
  const chosenYear = yearSelVal ? (parseInt(yearSelVal, 10) || 2026) : 2026;
  const parsed = parseWeekInput(inpVal);

  const wkMatch = fileName ? fileName.match(/(?:WK|Wk|wk)[-_ ]?(\d+)/i) : null;
  const fallbackWeek = wkMatch && Number(wkMatch[1]) >= 1 && Number(wkMatch[1]) <= 53 ? Number(wkMatch[1]) : 40;

  let week = fallbackWeek;
  let shipping_week = '';

  if (parsed && parsed.isAll) {
    // Chế độ xuất toàn bộ 26 Family: không truyền shipping_week để giữ trọn vẹn tất cả Family
    shipping_week = '';
  } else if (parsed && parsed.week) {
    week = parsed.week;
    shipping_week = `${chosenYear}/${String(week).padStart(2, '0')}`;
  } else {
    // Mặc định xuất toàn bộ 26 Family nếu là Tất cả hoặc để trống
    shipping_week = '';
  }

  const names = `${fileName} ${state.config?.template || ''}`;
  const seasonMatch = names.match(/(?:^|[^0-9])(?!20\d{2})(\d{4})(?:[^0-9]|$)/);
  return {
    week,
    shipping_week,
    season: seasonMatch?.[1] || '2728',
    start: '',
    end: '',
    use_local: state.useLocal,
    original_filename: fileName
  };
}

function clearInputFile(e) {
  if (e) { e.preventDefault(); e.stopPropagation(); }
  resetFileSelection();
  $('upload-error').hidden = true;
}
const btnClearInput = $('clear-input');
if (btnClearInput) btnClearInput.addEventListener('click', clearInputFile);

function inputChanged(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.xlsx')) { return toast('Vui lòng chọn file Excel .xlsx.', true); }
  if (file.size > 50 * 1024 * 1024) { return toast('File vượt quá 50 MB.', true); }
  state.input = file;
  state.useLocal = false;
  $('upload-error').hidden = true;
  setFileSelectedUI(file.name, file.size);
  toast(`Đã nhận file: ${file.name}. Hãy kiểm tra tuần và bấm "Tạo kết quả theo tuần".`);
}

$('input-file').addEventListener('change', e => inputChanged(e.target.files[0]));

const btnSelectExcel = $('btn-select-file');
if (btnSelectExcel) {
  btnSelectExcel.addEventListener('click', e => {
    e.stopPropagation();
    $('input-file').click();
  });
}

const btnChangeFile = $('btn-change-file');
if (btnChangeFile) {
  btnChangeFile.addEventListener('click', e => {
    e.stopPropagation();
    resetFileSelection();
    $('input-file').click();
  });
}

const btnGotoFamilyView = $('btn-goto-family-view');
if (btnGotoFamilyView) {
  btnGotoFamilyView.addEventListener('click', () => {
    switchTab('review');
  });
}

const dropZone = $('drop-input');
if (dropZone) {
  dropZone.addEventListener('click', e => {
    if (e.target.closest('#btn-select-file')) return;
    $('input-file').click();
  });
  for (const event of ['dragenter', 'dragover']) dropZone.addEventListener(event, e => {
    e.preventDefault();
    dropZone.classList.add('dragging');
  });
  for (const event of ['dragleave', 'drop']) dropZone.addEventListener(event, e => {
    e.preventDefault();
    dropZone.classList.remove('dragging');
  });
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('dragging');
    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
      inputChanged(e.dataTransfer.files[0]);
    }
  });
}

const guideHeader = document.querySelector('#card-guide .side-card-header');
if (guideHeader) {
  guideHeader.style.cursor = 'pointer';
  guideHeader.addEventListener('click', () => {
    const timeline = document.querySelector('.guide-timeline');
    const alertBox = document.querySelector('.guide-alert-box');
    const chevron = document.querySelector('.chevron-up-icon');
    if (timeline && alertBox) {
      const isHidden = timeline.hidden;
      timeline.hidden = !isHidden;
      alertBox.hidden = !isHidden;
      if (chevron) chevron.textContent = isHidden ? '^' : 'v';
    }
  });
}

const btnUseLocal = $('use-local');
if (btnUseLocal) {
  btnUseLocal.addEventListener('click', () => {
    state.useLocal = true; state.input = null;
    $('upload-error').hidden = true;
    startProcessing();
  });
}
let currentUploadRequest = null;
let progressClock = null;
let progressStarted = 0;
function stopProgressClock() {
  clearInterval(progressClock); progressClock = null;
}
function startProgressClock() {
  stopProgressClock(); progressStarted = Date.now();
  const tick = () => { $('progress-time').textContent = `Đã chạy: ${Math.floor((Date.now() - progressStarted) / 1000)} giây`; };
  tick(); progressClock = setInterval(tick, 1000);
}
function renderJobProgress(job) {
  if (isCancelled) return;
  $('progress-message').textContent = job.message;
  const known = job.total > 0;
  const ready = job.status === 'ready';
  const checking = job.phase === 'validate';
  // Reserve the final portion for the actual source comparison pass.
  const pct = ready ? 100 : known ? Math.min(99, Math.floor(checking ?
    80 + 19 * (job.checked || 0) / job.total : 80 * job.completed / job.total)) : 0;
  document.querySelector('.progress-track')?.classList.toggle('indeterminate', !known && !ready);
  $('progress-fill').style.width = known || ready ? `${pct}%` : '25%';
  if ($('progress-percent')) $('progress-percent').textContent = known || ready ? `${pct}%` : 'Đang xử lý';
  $('progress-count').textContent = known ? `(${checking ? job.checked || 0 : job.completed} / ${job.total} Family${checking ? ' đối chiếu' : ''})` : 'Đọc và kiểm tra cấu trúc';
  if ($('progress-eta')) $('progress-eta').textContent = ready ? 'Hoàn tất!' : checking ? 'Đang đối chiếu…' : 'Đang xử lý…';
}

function progressDecoder(onJob) {
  let offset = 0, pending = '';
  return (text, final = false) => {
    pending += text.slice(offset); offset = text.length;
    const lines = pending.split('\n'); pending = lines.pop();
    if (final && pending.trim()) { lines.push(pending); pending = ''; }
    for (const line of lines) if (line.trim()) onJob(JSON.parse(line));
  };
}
let isCancelled = false;

function cancelCurrentProcess() {
  isCancelled = true;
  if (currentUploadRequest) {
    try { currentUploadRequest.abort(); } catch (e) { }
    currentUploadRequest = null;
  }
  if (state.poll) {
    clearTimeout(state.poll);
    state.poll = null;
  }
  state.job = null;
  state.report = null;
  setExportState(false);
  showView('upload');
  $('process-button').disabled = !state.input && !state.useLocal;
  $('progress-message').textContent = 'Đang chuẩn bị dữ liệu…';
  $('progress-fill').style.width = '0%';
  if ($('progress-percent')) $('progress-percent').textContent = '0%';
  if ($('progress-eta')) $('progress-eta').textContent = 'Đang tính toán...';
  const track = document.querySelector('.progress-track');
  if (track) track.classList.remove('indeterminate');
  const confirmDialog = $('cancel-confirm-dialog');
  if (confirmDialog && confirmDialog.open) confirmDialog.close();
  toast('Đã hủy tải lên.');
}

const btnCancelProgress = $('btn-cancel-progress');
if (btnCancelProgress) {
  btnCancelProgress.addEventListener('click', e => {
    e.preventDefault();
    e.stopPropagation();
    const dialog = $('cancel-confirm-dialog');
    if (dialog) dialog.showModal();
  });
}

const btnAbortDismiss = $('btn-abort-dismiss');
if (btnAbortDismiss) {
  btnAbortDismiss.addEventListener('click', () => {
    const dialog = $('cancel-confirm-dialog');
    if (dialog) dialog.close();
  });
}

const btnAbortConfirm = $('btn-abort-confirm');
if (btnAbortConfirm) {
  btnAbortConfirm.addEventListener('click', () => {
    cancelCurrentProcess();
  });
}

const cancelDialog = $('cancel-confirm-dialog');
if (cancelDialog) {
  cancelDialog.addEventListener('click', e => {
    if (e.target === cancelDialog) cancelDialog.close();
  });
}

async function uploadJob(options) {
  let uploadFile = state.input;
  if (uploadFile) {
    $('progress-message').textContent = 'Đang tải toàn bộ workbook để giữ đủ các sheet và công thức…';
  }
  const form = new FormData();
  form.append('options', JSON.stringify(options));
  if (uploadFile) form.append('input', uploadFile, uploadFile.name);
  return new Promise((resolve, reject) => {
    const request = new XMLHttpRequest();
    currentUploadRequest = request;
    request.open('POST', '/api/jobs');
    request.setRequestHeader('Accept', 'application/x-ndjson');
    let latest = null, streamError = null;
    const decode = progressDecoder(job => { latest = job; renderJobProgress(job); });
    const streaming = () => request.getResponseHeader('Content-Type')?.includes('application/x-ndjson');
    request.onprogress = () => {
      if (isCancelled || !streaming()) return;
      try { decode(request.responseText); } catch (error) { streamError = error; }
    };
    request.upload.onprogress = event => {
      if (!event.lengthComputable || isCancelled) return;
      const percent = Math.round(event.loaded / event.total * 100);
      $('progress-message').textContent = percent < 100 ? `Đang tải file lên… ${percent}%` : 'Đã gửi file. Đang chờ máy chủ tiếp nhận…';
      $('progress-count').textContent = `${(event.loaded / 1024 / 1024).toFixed(1)} / ${(event.total / 1024 / 1024).toFixed(1)} MB`;
    };
    request.onload = () => {
      currentUploadRequest = null;
      if (isCancelled) { reject(new Error('ABORTED')); return; }
      try {
        if (streaming()) {
          decode(request.responseText, true);
          if (streamError) throw streamError;
          if (latest?.status === 'error') throw new Error(latest.message);
          if (latest?.status !== 'ready') throw new Error('Kết nối kết thúc trước khi xử lý xong. Hãy kiểm tra phiên trong lịch sử.');
        } else latest = JSON.parse(request.responseText);
        if (request.status >= 200 && request.status < 300 && latest) { resolve(latest); return; }
        throw new Error(latest?.error || 'Máy chủ chưa tiếp nhận được file.');
      } catch (error) { reject(new Error(request.status === 413 ? 'File vượt giới hạn tải lên của máy chủ.' : error.message)); }
    };
    request.onabort = () => {
      currentUploadRequest = null;
      reject(new Error('ABORTED'));
    };
    request.onerror = () => {
      currentUploadRequest = null;
      if (isCancelled) { reject(new Error('ABORTED')); return; }
      reject(new Error('Mất kết nối khi tải file lên. Vui lòng thử lại.'));
    };
    request.send(form);
  });
}
async function startProcessing() {
  $('upload-error').hidden = true;
  isCancelled = false;
  try {
    if (!state.input && !state.useLocal) throw new Error('Hãy chọn file nguồn.');
    if ($('process-button')) $('process-button').disabled = true;
    setExportState(false);
    showView('progress'); $('progress-message').textContent = 'Đang tải file lên…'; $('progress-count').textContent = 'Khởi tạo'; $('progress-fill').style.width = '5%';
    startProgressClock();
    if ($('progress-percent')) $('progress-percent').textContent = '0%';
    if ($('progress-eta')) $('progress-eta').textContent = 'Đang ước tính...';
    $('progress-time').textContent = 'Đã chạy: 0 giây';
    const initialTrack = document.querySelector('.progress-track');
    if (initialTrack) initialTrack.classList.add('indeterminate');
    const payload = outputOptions(state.input?.name || state.config?.local_input || '');
    const job = await uploadJob(payload);
    if (isCancelled) return;
    state.job = job; state.report = null; state.preview = null;
    await pollJob(job.id, job);
  } catch (error) {
    if (error.message === 'ABORTED' || isCancelled) return;
    showView('upload'); $('upload-error').textContent = error.message; $('upload-error').hidden = false;
  }
  finally { if (!isCancelled && $('process-button')) $('process-button').disabled = state.view === 'progress'; }
}
const btnProcess = $('process-button');
if (btnProcess) btnProcess.addEventListener('click', startProcessing);

async function pollJob(id, receivedJob = null) {
  if (isCancelled) return;
  clearTimeout(state.poll);
  const job = receivedJob || await api(`/api/jobs/${id}`);
  if (isCancelled) return;
  state.job = job;
  renderJobProgress(job);
  if (job.status === 'ready') {
    await openJob(id, 'review');
    loadConfig().catch(error => toast(error.message, true));
    return;
  }
  if (job.status === 'error') { setExportState(false); showView('upload'); $('upload-error').textContent = job.message; $('upload-error').hidden = false; await loadConfig(); return; }
  if (!isCancelled) {
    state.poll = setTimeout(() => pollJob(id).catch(error => { if (!isCancelled) { toast(error.message, true); showView('upload'); } }), 1200);
  }
}
async function loadConfig() {
  state.config = await api('/api/config');
  if ($('use-local')) $('use-local').hidden = !state.config.local_input;
  if ($('local-name')) $('local-name').textContent = state.config.local_input || '';
  const clearBtn = $('btn-clear-all-history');
  if (clearBtn) clearBtn.hidden = !state.config.jobs.length;

  const historyEl = $('history');
  if (historyEl) {
    if (state.config.jobs && state.config.jobs.length) {
      historyEl.innerHTML = state.config.jobs.map(job => `
        <div class="recent-file-item" data-job="${job.id}">
          <button type="button" class="btn-open-recent" data-open-job="${job.id}" title="Mở lại phiên này">
            <div class="recent-file-icon">X</div>
            <div class="recent-file-info">
              <span class="recent-file-name" title="${escapeHTML(job.source)}">${escapeHTML(job.source)}</span>
              <span class="recent-file-sub">${formatDateTime(job.created)} | ${job.status === 'ready' ? `${job.total || 26} Family` : job.status === 'error' ? 'Bị gián đoạn' : 'Đang xử lý'}</span>
            </div>
          </button>
          <button type="button" class="btn-delete-recent" data-delete-job="${job.id}" title="Xóa phiên này khỏi lịch sử" aria-label="Xóa">
            ✕
          </button>
        </div>
      `).join('');
    } else {
      historyEl.innerHTML = '<p class="muted" style="padding:14px 10px;font-size:12px;color:#94a3b8;">Chưa có phiên xử lý.</p>';
    }
  }
  updateUploadStateUI();
}

$('history')?.addEventListener('click', async e => {
  const deleteBtn = e.target.closest('[data-delete-job]');
  if (deleteBtn) {
    e.stopPropagation();
    const jobId = deleteBtn.dataset.deleteJob;
    if (!confirm('Bạn có chắc chắn muốn xóa phiên xử lý này khỏi lịch sử?')) return;
    try {
      await api(`/api/jobs/${jobId}/delete`, {});
      try { localStorage.removeItem(`job_${jobId}_file_states`); } catch (err) { }
      if (state.job && state.job.id === jobId) {
        state.job = null;
        state.report = null;
        state.preview = null;
        setExportState(false);
        updateUploadStateUI();
        showView('upload');
      }
      await loadConfig();
      toast('Đã xóa phiên khỏi lịch sử.');
    } catch (err) {
      toast(err.message, true);
    }
    return;
  }
  const button = e.target.closest('[data-open-job]');
  if (!button) return;
  try {
    const job = await api(`/api/jobs/${button.dataset.openJob}`);
    if (job.status === 'ready') await openJob(job.id, 'upload');
    else if (job.status === 'error') toast(job.message, true);
    else { showView('progress'); startProgressClock(); await pollJob(job.id); }
  } catch (error) {
    toast(error.message, true);
  }
});

const btnClearAllHistory = $('btn-clear-all-history');
if (btnClearAllHistory) {
  btnClearAllHistory.addEventListener('click', async e => {
    e.preventDefault();
    e.stopPropagation();
    if (!state.config?.jobs?.length) {
      toast('Chưa có lịch sử nào để xóa.');
      return;
    }
    if (!confirm(`Bạn có chắc chắn muốn xóa TOÀN BỘ ${state.config.jobs.length} phiên trong lịch sử?`)) return;
    try {
      await api('/api/jobs/clear', {});
      if (state.job) {
        state.job = null;
        state.report = null;
        state.preview = null;
        setExportState(false);
        updateUploadStateUI();
        showView('upload');
      }
      await loadConfig();
      toast('Đã xóa toàn bộ lịch sử.');
    } catch (err) {
      toast(err.message, true);
    }
  });
}
function currentEntry() {
  if (!state.report) return null;
  if (state.file === -1) {
    if (!state.report.source_file) {
      let srcName = state.job?.source || state.report.source;
      if (!srcName || srcName.toLowerCase() === 'input.xlsx') {
        srcName = state.config?.local_input || '2026 INTERNAL SCHEDULE FERRERO-WK40.xlsx';
      }
      state.report.source_file = { id: -1, name: srcName, file: srcName, original_name: srcName, items: 0, markets: 0, revision: 0, edits: [], is_source: true, changed_cells: 0, status_label: 'matched', problems: [] };
    } else if (state.report.source_file.original_name === 'input.xlsx' || state.report.source_file.name === 'input.xlsx') {
      const real = state.job?.source || state.config?.local_input || '2026 INTERNAL SCHEDULE FERRERO-WK40.xlsx';
      state.report.source_file.original_name = real;
      state.report.source_file.name = real;
      state.report.source_file.file = real;
    }
    return state.report.source_file;
  }
  return state.report.families.find(f => f.id === state.file) || state.report.families[0];
}

async function openJob(id, targetTab = 'upload') {
  clearTimeout(state.poll);
  state.job = await api(`/api/jobs/${id}`);
  state.report = await api(`/api/jobs/${id}/report`);
  state.filter = 'all'; state.search = ''; if ($('family-search')) $('family-search').value = '';
  state.file = state.report.families.find(f => f.name === 'BARBIE 2728')?.id ?? state.report.families[0]?.id ?? -1; state.sheet = 0;
  saveFileState(state.file, { viewed: true });
  if ($('source-caption')) $('source-caption').textContent = state.job.source;
  setExportState(state.report.families.length > 0, `/api/jobs/${state.job.id}/zip`);
  document.querySelectorAll('[data-filter]').forEach(e => e.classList.toggle('active', e.dataset.filter === 'all'));
  renderMetrics(); renderFamilies();
  updateUploadStateUI();
  showView(targetTab);
  await loadPreview();
  history.replaceState(null, '', `?job=${id}`);
  switchTab(targetTab);
  if (targetTab === 'review') {
    $('view-review').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

function renderMetrics() {
  const families = state.report.families;
  const filter = state.report.shipping_filter;
  const summary = $('shipping-filter-summary');
  if (summary) {
    summary.hidden = !filter;
    summary.textContent = filter ? `Tuần ${filter.week}: xuất ${filter.selected}/${filter.detected} Family. ${filter.excluded.length} Family không có số lượng giao > 0 nên không xuất. Giữ đủ ${state.report.weeks} tuần trong mỗi file.` : '';
  }
  const warnings = families.filter(f => f.problems.length).length;
  const edited = families.filter(f => f.changed_cells > 0).length;
  const metrics = [[families.length, 'File Family đã tạo', '▦', ''], [state.report.weeks, 'Tuần trong kế hoạch', '▤', ''], [warnings, 'Family cần kiểm tra', '◎', 'amber'], [edited, 'Family đã chỉnh sửa', '✎', 'blue']];
  if ($('metrics')) $('metrics').innerHTML = metrics.map(([value, label, icon, color]) => `<div class="metric"><div><strong>${value}</strong><label>${label}</label></div><span class="metric-icon ${color}">${icon}</span></div>`).join('');
  const badge = $('tab-review-badge');
  if (badge) {
    badge.textContent = `${families.length} Family`;
    badge.hidden = false;
  }
}

function downloadSingleFile(fileId) {
  if (!state.job) return;
  const cur = fileId === -1 ? state.report?.source_file : state.report?.families?.find(f => f.id === fileId);
  let fileName = cur ? (cur.original_name || cur.file || cur.name) : 'export.xlsx';
  if (fileId === -1 && (!fileName || fileName.toLowerCase() === 'input.xlsx')) {
    fileName = state.job?.source || state.config?.local_input || '2026 INTERNAL SCHEDULE FERRERO-WK40.xlsx';
  }
  const a = document.createElement('a');
  a.href = `/api/jobs/${state.job.id}/download?file=${fileId}`;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
  trackDownload(fileId);
  toast(`Đang tải xuống: ${fileName}`);
}

function downloadAllFiles() {
  if (!state.job) return;
  const a = document.createElement('a');
  a.href = `/api/jobs/${state.job.id}/zip`;
  a.download = `cac_file_family_${state.job.id.slice(0, 8)}.zip`;
  document.body.appendChild(a);
  a.click();
  a.remove();

  if (state.report) {
    // Chỉ tải tất cả các file Family output, không tải và không đánh dấu file nguồn
    (state.report.families || []).forEach(f => {
      saveFileState(f.id, { downloaded: true, downloaded_rev: f.revision || 0 });
    });
    renderFamilies();
  }
  toast('Đang tải xuống tất cả các file Family output (.zip)...');
}

if ($('btn-download-all-files')) {
  $('btn-download-all-files').addEventListener('click', downloadAllFiles);
}

$('family-search').addEventListener('input', e => { state.search = e.target.value; renderFamilies(); });
document.querySelectorAll('[data-filter]').forEach(button => button.addEventListener('click', () => {
  state.filter = button.dataset.filter;
  document.querySelectorAll('[data-filter]').forEach(el => el.classList.toggle('active', el === button));
  renderFamilies();
}));
async function openExcelPreview(sheetIdx) {
  if (sheetIdx === undefined) previewModalSheet = state.sheet;
  else previewModalSheet = sheetIdx;
  const dialog = $('excel-preview-dialog');
  if (!dialog) return;
  if (!state.preview || !state.preview.file) {
    toast('Chưa có dữ liệu để xem trước.', true);
    return;
  }
  $('excel-preview-title').textContent = `Xem trước: ${state.preview.file.name}.xlsx`;
  if ($('excel-preview-download')) $('excel-preview-download').href = `/api/jobs/${state.job.id}/download?file=${state.file}`;

  $('excel-preview-tabs').innerHTML = state.preview.sheets.map((name, idx) =>
    `<button role="tab" aria-selected="${idx === previewModalSheet}" class="sheet-tab ${idx === previewModalSheet ? 'selected' : ''}" data-prev-sheet="${idx}">
       ${idx === 0 ? '▦ ' : idx === 1 ? '▤ ' : '▥ '}${escapeHTML(name)}
     </button>`
  ).join('');

  dialog.showModal();
  await renderExcelWysiwygGrid(previewModalSheet);
}

async function renderExcelWysiwygGrid(sheetIdx) {
  const container = $('excel-preview-grid');
  container.innerHTML = '<div class="empty-state">Đang tải cấu trúc và định dạng Excel…</div>';
  try {
    const data = await api(`/api/jobs/${state.job.id}/preview?file=${state.file}&sheet=${sheetIdx}`);
    const covered = new Set(), merges = new Map();
    for (const merge of data.merges) {
      const [start, end] = merge.split(':').map(coords);
      merges.set(`${start.row},${start.col}`, { rows: end.row - start.row + 1, cols: end.col - start.col + 1 });
      for (let r = start.row; r <= end.row; r++) {
        for (let c = start.col; c <= end.col; c++) {
          if (r !== start.row || c !== start.col) covered.add(`${r},${c}`);
        }
      }
    }

    let html = `<table class="wysiwyg-table" aria-label="${escapeHTML(data.sheet)}">`;
    html += '<colgroup><col style="width: 42px;">';
    data.columns.forEach((c) => {
      const width = (data.col_widths && data.col_widths[c]) ? data.col_widths[c] : 85;
      html += `<col style="width: ${width}px;">`;
    });
    html += '</colgroup><thead><tr><th></th>';
    data.columns.forEach(c => {
      html += `<th scope="col">${c}</th>`;
    });
    html += '</tr></thead><tbody>';

    data.rows.forEach((row, r) => {
      const rowHeight = (data.row_heights && data.row_heights[String(r + 1)]) ? `height: ${data.row_heights[String(r + 1)]}px;` : '';
      html += `<tr style="${rowHeight}"><td class="wysiwyg-row-num">${r + 1}</td>`;
      row.forEach((cell, c) => {
        if (covered.has(`${r},${c}`)) return;
        const merge = merges.get(`${r},${c}`);
        const styles = [];
        if (cell.bg) styles.push(`background-color: ${cell.bg} !important`);
        if (cell.fg) styles.push(`color: ${cell.fg} !important`);
        if (cell.bold) styles.push('font-weight: 700 !important');
        if (cell.italic) styles.push('font-style: italic !important');
        if (cell.font_size) styles.push(`font-size: ${Math.max(10, Math.round(cell.font_size * 1.15))}px !important`);
        if (cell.align) styles.push(`text-align: ${cell.align}`);
        else if (typeof cell.value === 'number') styles.push('text-align: right');
        if (cell.valign) styles.push(`vertical-align: ${cell.valign === 'center' ? 'middle' : cell.valign}`);
        if (cell.wrap) styles.push('white-space: pre-wrap; word-break: break-word');
        else styles.push('white-space: nowrap; text-overflow: ellipsis');

        if (cell.borders) {
          for (const [side, b] of Object.entries(cell.borders)) {
            const width = b.style === 'double' ? '3px double' : (b.style === 'medium' || b.style === 'thick') ? '2px solid' : '1px solid';
            styles.push(`border-${side}: ${width} ${b.color} !important`);
          }
        }

        const styleAttr = styles.length ? ` style="${styles.join('; ')}"` : '';
        const spanAttr = merge ? ` rowspan="${merge.rows}" colspan="${merge.cols}"` : '';
        const text = cell.value === null ? '' : cellDisplay(cell);
        html += `<td${spanAttr}${styleAttr} title="${escapeHTML(cell.address + (text ? ': ' + text : ''))}">${escapeHTML(text)}</td>`;
      });
      html += '</tr>';
    });

    html += '</tbody></table>';
    container.innerHTML = html;
    $('excel-preview-info').textContent = `Sheet: ${data.sheet.trim()} · ${data.rows.length} dòng × ${data.columns.length} cột`;
  } catch (err) {
    container.innerHTML = `<div class="empty-state">${escapeHTML(err.message)}</div>`;
  }
}

function toggleExcelPreviewFullscreen() {
  const dialog = $('excel-preview-dialog');
  if (!dialog) return;
  const isFs = dialog.classList.toggle('is-fullscreen');
  const btn = $('btn-fullscreen-preview');
  if (btn) {
    btn.title = isFs ? 'Thu nhỏ' : 'Toàn màn hình';
    btn.setAttribute('aria-label', btn.title);
  }
  const iconExpand = $('fullscreen-icon-expand');
  const iconCompress = $('fullscreen-icon-compress');
  if (iconExpand) iconExpand.hidden = isFs;
  if (iconCompress) iconCompress.hidden = !isFs;
}

function closeExcelPreviewModal() {
  const dialog = $('excel-preview-dialog');
  if (!dialog) return;
  dialog.close();
  dialog.classList.remove('is-fullscreen');
  const btn = $('btn-fullscreen-preview');
  if (btn) {
    btn.title = 'Toàn màn hình';
    btn.setAttribute('aria-label', 'Toàn màn hình');
  }
  if ($('fullscreen-icon-expand')) $('fullscreen-icon-expand').hidden = false;
  if ($('fullscreen-icon-compress')) $('fullscreen-icon-compress').hidden = true;
}

if ($('btn-preview-excel')) $('btn-preview-excel').addEventListener('click', () => openExcelPreview());
if ($('btn-fullscreen-preview')) $('btn-fullscreen-preview').addEventListener('click', toggleExcelPreviewFullscreen);
if ($('close-excel-preview')) $('close-excel-preview').addEventListener('click', closeExcelPreviewModal);
if ($('excel-preview-dialog')) $('excel-preview-dialog').addEventListener('click', e => {
  if (e.target === $('excel-preview-dialog')) closeExcelPreviewModal();
});
if ($('excel-preview-tabs')) $('excel-preview-tabs').addEventListener('click', async e => {
  const btn = e.target.closest('[data-prev-sheet]');
  if (!btn) return;
  const idx = Number(btn.dataset.prevSheet);
  previewModalSheet = idx;
  document.querySelectorAll('#excel-preview-tabs .sheet-tab').forEach(b => {
    const isSel = (b === btn);
    b.classList.toggle('selected', isSel);
    b.setAttribute('aria-selected', isSel ? 'true' : 'false');
  });
  await renderExcelWysiwygGrid(idx);
});
function draftKey() { return state.job ? state.job.id + ':' + state.file : ''; }
function currentDraft() { return workbookDrafts.get(draftKey()); }
function pendingEdit(sheet, address) { return currentDraft()?.edits.get(sheet + ':' + address); }
function rawCell(cell) { return String((state.file === -1 ? cell.formula : null) ?? cell.value ?? ''); }
function updateFormulaBar(cell = state.cell) {
  const address = $('formula-cell-address'), expression = $('formula-cell-expression');
  if (!address || !expression) return;
  address.value = cell?.address || '';
  expression.value = formulaBarText(cell, cell ? pendingEdit(state.sheet, cell.address) : null);
  expression.title = expression.value;
}
function updateEditToolbar() {
  const draft = currentDraft(), count = draft?.edits.size || 0;
  $('edit-workbook').hidden = Boolean(draft);
  $('edit-workbook').disabled = true;
  $('edit-workbook').title = 'Chức năng chỉnh sửa tạm thời bị vô hiệu hóa';
  $('save-workbook').hidden = !draft;
  $('cancel-workbook').hidden = !draft;
  $('save-workbook').disabled = state.saving || !count;
  $('cancel-workbook').disabled = state.saving;
  $('save-workbook').textContent = state.saving ? 'Đang lưu và tính lại…' : `Lưu${count ? ' (' + count + ' ô)' : ''}`;
  document.querySelector('.workbook-panel').classList.toggle('inline-editing', Boolean(draft));
}
function stageCell(cell, raw, sheet = state.sheet) {
  const draft = currentDraft(); if (!draft || state.saving || !cell.editable) return;
  const key = sheet + ':' + cell.address;
  if (raw === rawCell(cell)) draft.edits.delete(key);
  else {
    if (draft.edits.size >= 1000 && !draft.edits.has(key)) throw new Error('Tối đa 1000 ô mỗi lần lưu.');
    let kind = 'text', value = raw;
    if (state.file === -1) {
      if (raw === '') kind = 'blank';
      else if (raw.startsWith("'")) { kind = 'text'; value = raw.slice(1); }
      else if (raw.startsWith('=')) kind = 'formula';
      else if (/^(true|false)$/i.test(raw)) kind = 'boolean';
      else if (cell.input_type === 'date') kind = 'date';
      else if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(raw)) kind = 'number';
    }
    draft.edits.set(key, { sheet, cell: cell.address, value, value_type: kind, raw });
  }
  updateEditToolbar();
}
function focusCell(address) {
  const element = $('grid-container').querySelector(`[data-address="${address}"]`); if (!element) return;
  $('grid-container').querySelectorAll('.selected-cell').forEach(el => el.classList.remove('selected-cell'));
  element.classList.add('selected-cell'); element.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}
function openCell(address) {
  const { row, col } = coords(address), cell = state.preview?.rows[row]?.[col]; if (!cell) return;
  focusCell(address); state.cell = cell;
  updateFormulaBar(cell);
  if (!currentDraft() || state.saving) return;
  if (!cell.editable) { toast('Ô này được tính tự động hoặc thuộc cấu trúc bảng.'); return; }
  const td = $('grid-container').querySelector(`[data-address="${address}"]`);
  if (!td || td.querySelector('input')) return;
  const input = document.createElement('input');
  const before = pendingEdit(state.sheet, address)?.raw ?? rawCell(cell);
  input.className = 'inline-cell-input'; input.value = before; input.setAttribute('aria-label', `Sửa ${state.preview.sheet}!${address}`);
  input.autocomplete = 'off'; input.spellcheck = false;
  td.replaceChildren(input);
  input.addEventListener('input', () => { try { stageCell(cell, input.value); td.classList.toggle('draft-cell', Boolean(pendingEdit(state.sheet, address))); updateFormulaBar(cell); } catch (error) { toast(error.message, true); } });
  input.addEventListener('blur', () => {
    const pending = pendingEdit(state.sheet, address);
    if (input.isConnected) td.textContent = pending ? pending.raw : cell.value === null ? '' : cellDisplay(cell);
    td.classList.toggle('draft-cell', Boolean(pending));
  });
  input.addEventListener('keydown', e => {
    if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); stageCell(cell, before); input.blur(); return; }
    if (e.key === 'Tab' || e.key === 'Enter') {
      e.preventDefault(); e.stopPropagation(); input.blur();
      const cells = [...$('grid-container').querySelectorAll('[data-address].is-editable')];
      const next = cells[cells.indexOf(td) + (e.shiftKey ? -1 : 1)];
      if (next) openCell(next.dataset.address);
    }
  });
  input.addEventListener('paste', e => {
    const text = e.clipboardData.getData('text/plain'); if (!/[\t\r\n]/.test(text)) return;
    e.preventDefault();
    let skipped = 0;
    const rows = text.replace(/\r/g, '').replace(/\n$/, '').split('\n');
    try {
      rows.forEach((line, dr) => line.split('\t').forEach((raw, dc) => {
        const dest = state.preview.rows[row + dr]?.[col + dc];
        if (dest?.editable) stageCell(dest, raw); else skipped++;
      }));
    } catch (error) { toast(error.message, true); }
    input.blur(); renderGrid(); updateEditToolbar();
    if (skipped) toast(`Đã bỏ qua ${skipped} ô ngoài bảng hoặc chỉ đọc.`);
  });
  input.focus(); input.select();
}
$('grid-container').addEventListener('click', e => {
  if (e.target.closest('input')) return;
  const cell = e.target.closest('[data-address]'); if (cell) openCell(cell.dataset.address);
});
$('grid-container').addEventListener('keydown', e => {
  if (e.target.closest('input')) return;
  if (e.key === 'Enter' || e.key === 'F2') { const cell = e.target.closest('[data-address]'); if (cell) { e.preventDefault(); openCell(cell.dataset.address); } }
});
$('edit-workbook').addEventListener('click', () => {
  return;
  if (!state.preview) return;
  workbookDrafts.set(draftKey(), { revision: state.preview.revision, edits: new Map() });
  updateEditToolbar();
});
$('cancel-workbook').addEventListener('click', () => {
  if (state.saving) return;
  workbookDrafts.delete(draftKey()); renderGrid(); updateEditToolbar();
  toast('Đã hủy các thay đổi chưa lưu của file này.');
});
async function saveWorkbook() {
  document.activeElement?.blur();
  const draft = currentDraft(); if (state.saving || !draft?.edits.size) return;
  const job = state.job.id, file = state.file, sheet = state.sheet, key = draftKey();
  state.saving = true; updateEditToolbar();
  try {
    const edits = [...draft.edits.values()].map(({ raw, ...edit }) => edit);
    const result = await api(`/api/jobs/${job}/edit/${file}`, { revision: draft.revision, edits });
    workbookDrafts.delete(key);
    if (state.job?.id !== job) return;
    await openJob(result.new_job_id);
    state.file = file; state.sheet = sheet; renderFamilies(); await loadPreview(); await loadConfig();
    toast('Đã lưu toàn bộ thay đổi vào phiên mới. Phiên cũ vẫn được giữ trong lịch sử.');
  } catch (error) { toast(error.message + ' Các ô chưa lưu vẫn được giữ trên bảng.', true); }
  finally { state.saving = false; updateEditToolbar(); }
}
$('save-workbook').addEventListener('click', saveWorkbook);
document.addEventListener('keydown', e => { if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && currentDraft()) { e.preventDefault(); saveWorkbook(); } });
window.addEventListener('beforeunload', e => { if ([...workbookDrafts.values()].some(d => d.edits.size)) { e.preventDefault(); e.returnValue = ''; } });
function getProblemSuggestion(p) {
  if (p && p.suggestion) return p.suggestion;
  if (!p) return '';
  if (p.kind === 'week') {
    const raw = (p.message || '').match(/([0-9]{4})\/([0-9]+)/);
    if (raw) {
      const year = parseInt(raw[1], 10);
      const week = parseInt(raw[2], 10);
      const nextYear = year + 1;
      if (week === 53) {
        return `Năm ${year} theo lịch ISO chỉ có 52 tuần (không có tuần 53). Gợi ý sửa: đổi thành "${nextYear}/01" (tuần đầu năm sau) hoặc kiểm tra lại file gốc nếu thuộc tuần ${year}/52.`;
      }
      if (week > 53) {
        return `Số tuần vượt quá giới hạn năm ${year}. Gợi ý kiểm tra lại hoặc chuyển sang năm ${nextYear}.`;
      }
      return `Nhập lại tuần theo định dạng YYYY/WW hợp lệ (ví dụ: ${year}/01 hoặc 2028/01).`;
    }
    return 'Định dạng tuần chuẩn ISO là YYYY/WW (ví dụ: 2028/01). Nhấn vào đây để đến ô và nhập lại.';
  }
  if (p.kind === 'error') {
    const val = (p.message || '').toUpperCase();
    if (val.includes('#REF')) return 'Lỗi #REF!: Công thức tham chiếu ô không tồn tại hoặc bị xóa trong nguồn. Kiểm tra lại dữ liệu gốc.';
    if (val.includes('#VAL')) return 'Lỗi #VALUE!: Sai kiểu dữ liệu (chữ/số kết hợp). Kiểm tra lại giá trị đầu vào.';
    if (val.includes('#DIV')) return 'Lỗi #DIV/0!: Phép chia cho 0. Kiểm tra lại mẫu số trong công thức.';
    return 'Lỗi công thức Excel. Nhấn vào đây để đến ô và sửa lại dữ liệu chính xác.';
  }
  return 'Nhấn vào đây để chuyển đến ô cần kiểm tra và chỉnh sửa.';
}

function renderProblems() {
  const entry = currentEntry();
  if (!entry) { $('problem-count').textContent = '0'; $('problems').innerHTML = ''; return; }
  const list = entry.problems || [];
  const sorted = [...list].sort((a, b) => (a.sheet === 'Breakdown ' ? -1 : 1) - (b.sheet === 'Breakdown ' ? -1 : 1));
  $('problem-count').textContent = list.length;
  $('jump-warning').hidden = !list.length;
  if (state.file === -1) {
    $('problems').innerHTML = '<div class="empty-state">✓ File nguồn Excel gốc tải lên.<br>Không có điểm cần kiểm tra.</div>';
    return;
  }
  $('problems').innerHTML = sorted.length ? sorted.map(p => {
    const sugg = getProblemSuggestion(p);
    return `<button class="problem-row" data-problem-sheet="${escapeHTML(p.sheet)}" data-problem-cell="${p.cell}" type="button">
      <span class="problem-icon">!</span>
      <div class="problem-body">
        <div class="problem-title">${escapeHTML(p.message)}</div>
        <div class="problem-meta"><small>${escapeHTML(p.sheet)} · <strong>${p.cell}</strong>${p.kind === 'week' ? ' · cần xác nhận lại tuần' : ''}</small></div>
        ${sugg ? `<div class="problem-suggestion"><span class="sugg-icon">💡</span><div class="sugg-content"><strong>Gợi ý cách sửa:</strong> ${escapeHTML(sugg)}</div></div>` : ''}
      </div>
      <span class="problem-arrow" title="Đến ô và sửa">↗</span>
    </button>`;
  }).join('') : '<div class="empty-state">✓ Không có ô lỗi hoặc tuần ISO không hợp lệ.<br>Đối chiếu với nguồn không thay thế xác nhận nghiệp vụ.</div>';
}
$('problems').addEventListener('click', async e => {
  const button = e.target.closest('[data-problem-cell]');
  if (!button) return;
  const targetSheet = button.dataset.problemSheet;
  const targetCell = button.dataset.problemCell;
  const sheetIndex = state.preview.sheets.indexOf(targetSheet);
  if (sheetIndex >= 0 && sheetIndex !== state.sheet) {
    state.sheet = sheetIndex;
  }
  await loadPreview(targetCell);
  openCell(targetCell);
});
$('jump-warning').addEventListener('click', () => { $('problems').scrollIntoView({ behavior: 'smooth', block: 'center' }); });
function renderAudit() {
  const entry = currentEntry();
  if (!entry) { $('audit-list').innerHTML = ''; return; }
  const edits = entry.edits || [];
  $('audit-list').innerHTML = edits.length ? [...edits].reverse().map(e => `<div class="audit-row"><b>${escapeHTML(e.sheet)} · ${e.cell}</b><div class="audit-values"><del>${escapeHTML(display(e.before))}</del><span>→</span><span>${escapeHTML(display(e.after))}</span></div><small>${new Date(e.time).toLocaleString('vi-VN')}${e.restored ? ' · Khôi phục' : ''}${e.reason ? ' · ' + escapeHTML(e.reason) : ''}</small></div>`).join('') : '<div class="empty-state">Chưa có chỉnh sửa.<br>Mỗi thay đổi sẽ được lưu lại tại đây.</div>';
}

function trackDownload(fileId) {
  const cur = currentEntry();
  saveFileState(fileId, { downloaded: true, downloaded_rev: cur ? (cur.revision || 0) : 0 });
  renderFamilies();
}
$('download-one').addEventListener('click', () => trackDownload(state.file));
if ($('excel-preview-download')) {
  $('excel-preview-download').addEventListener('click', () => trackDownload(state.file));
}

function initWorkspaceResizer() {
  const resizer = $('workspace-resizer');
  const workspace = document.querySelector('.review-workspace');
  if (!resizer || !workspace) return;

  const savedWidth = localStorage.getItem('workspace_sidebar_width');
  if (savedWidth) {
    workspace.style.setProperty('--sidebar-width', `${savedWidth}px`);
  }

  let isDragging = false;
  let startX = 0;
  let startWidth = 0;

  resizer.addEventListener('mousedown', e => {
    e.preventDefault();
    isDragging = true;
    startX = e.clientX;
    const currentWidth = parseFloat(getComputedStyle(workspace).getPropertyValue('--sidebar-width')) || 255;
    startWidth = currentWidth;
    resizer.classList.add('is-dragging');
    document.body.classList.add('is-resizing-col');

    const onMouseMove = ev => {
      if (!isDragging) return;
      const delta = ev.clientX - startX;
      const newWidth = Math.min(650, Math.max(160, Math.round(startWidth + delta)));
      workspace.style.setProperty('--sidebar-width', `${newWidth}px`);
    };

    const onMouseUp = ev => {
      if (!isDragging) return;
      isDragging = false;
      resizer.classList.remove('is-dragging');
      document.body.classList.remove('is-resizing-col');
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
      const finalWidth = parseFloat(getComputedStyle(workspace).getPropertyValue('--sidebar-width')) || 255;
      try {
        localStorage.setItem('workspace_sidebar_width', finalWidth);
      } catch (err) { }
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  });
}

function initTableColumnResizer() {
  const container = $('grid-container');
  if (!container) return;

  let isResizing = false;
  let startX = 0;
  let startWidth = 0;
  let targetCol = null;
  let targetResizer = null;
  let activeColIdx = null;

  container.addEventListener('mousedown', e => {
    const resizer = e.target.closest('.col-resizer');
    if (!resizer) return;
    e.preventDefault();
    e.stopPropagation();

    activeColIdx = resizer.dataset.colIdx;
    targetResizer = resizer;
    const table = container.querySelector('.sheet-grid');
    if (!table) return;

    targetCol = table.querySelector(`col[data-col-idx="${activeColIdx}"]`);
    const th = resizer.closest('th');
    if (!targetCol || !th) return;

    isResizing = true;
    startX = e.clientX;
    startWidth = th.getBoundingClientRect().width;
    resizer.classList.add('is-resizing');
    document.body.classList.add('is-resizing-col');

    const onMouseMove = ev => {
      if (!isResizing || !targetCol) return;
      const delta = ev.clientX - startX;
      const newWidth = Math.max(35, Math.round(startWidth + delta));
      targetCol.style.width = `${newWidth}px`;
    };

    const onMouseUp = ev => {
      if (!isResizing) return;
      isResizing = false;
      if (targetResizer) targetResizer.classList.remove('is-resizing');
      document.body.classList.remove('is-resizing-col');
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);

      if (targetCol) {
        const finalWidth = parseFloat(targetCol.style.width);
        if (!state.colWidths) state.colWidths = {};
        if (!state.colWidths[state.file]) state.colWidths[state.file] = {};
        if (!state.colWidths[state.file][state.sheet]) state.colWidths[state.file][state.sheet] = {};
        state.colWidths[state.file][state.sheet][`col_${activeColIdx}`] = finalWidth;
      }
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  });
}

workbookController.bind({
  files: $('family-list'), sheets: $('sheet-tabs'),
  chooseSource() { state.choosingReplacement = true; $('input-file').value = ''; $('input-file').click(); },
  download: downloadSingleFile
});

weekPicker.bind();
initWorkspaceResizer();
initTableColumnResizer();
resetWeekOptions();

loadConfig().then(async () => {
  const params = new URLSearchParams(location.search), job = params.get('job');
  if (job && /^[a-f0-9]{32}$/.test(job)) {
    const cell = params.get('cell');
    await openJob(job, cell ? 'review' : 'upload');
    if (cell && /^[A-Z]+[1-9][0-9]*$/.test(cell)) openCell(cell);
  }
}).catch(error => toast('Không kết nối được ứng dụng: ' + error.message, true));
