// Owns selection and request lifetime. Views only receive the latest result.
export function createWorkbookController({state, api, view, markViewed, renderFiles}) {
  let sequence = 0;
  let activeRequest;

  async function loadPreview(focusAddress, preferSourceSummary = false) {
    if (!state.job) return false;
    const selection = {job: state.job.id, file: state.file, sheet: state.sheet};
    const request = ++sequence;
    activeRequest?.abort();
    activeRequest = new AbortController();
    const current = () => request === sequence && selection.job === state.job?.id &&
      selection.file === state.file && selection.sheet === state.sheet;
    state.preview = null;
    view.loading();
    try {
      const result = await api(`/api/jobs/${selection.job}/preview?file=${selection.file}&sheet=${selection.sheet}`,
        undefined, {signal: activeRequest.signal});
      if (!current()) return false;
      if (preferSourceSummary && result.file?.is_source) {
        const summary = result.sheets.findIndex(name => String(name).trim().toUpperCase() === 'SUM');
        if (summary >= 0 && summary !== selection.sheet) {
          state.sheet = summary;
          return loadPreview(focusAddress, false);
        }
      }
      state.preview = result;
      view.show(result, selection, focusAddress);
      markViewed(selection.file, {viewed: true});
      renderFiles();
      return true;
    } catch (error) {
      if (current() && error.name !== 'AbortError') view.error(error);
      return false;
    } finally {
      if (current()) view.setBusy(false);
    }
  }

  async function selectFile(file) {
    if (!Number.isInteger(file) || state.saving) return false;
    if (file === -1 ? !state.report?.source_file : !state.report?.families.some(f => f.id === file)) return false;
    state.file = file;
    // Master schedules use SUM as the review sheet. Index 3 avoids first
    // downloading Capsule; the response still verifies and corrects the hint.
    state.sheet = file === -1 ? (state.report.source_file?.summary_sheet_index ?? 3) : 0;
    renderFiles();
    return loadPreview(undefined, file === -1);
  }

  async function selectSheet(sheet) {
    if (!Number.isInteger(sheet) || sheet < 0 || state.saving) return false;
    state.sheet = sheet;
    return loadPreview();
  }

  let detach;
  function bind({files, sheets, chooseSource, download}) {
    detach?.(); // Reinitialization must never attach a second handler.
    const click = event => {
      if (event.target.closest('[data-choose-source]')) { chooseSource(); return; }
      const downloadButton = event.target.closest('[data-download-file]');
      if (downloadButton) { download(Number(downloadButton.dataset.downloadFile)); return; }
      const card = event.target.closest('[data-file]');
      if (card) void selectFile(Number(card.dataset.file));
    };
    const keydown = event => {
      if (!['Enter', ' '].includes(event.key) || !event.target.matches('[data-file]')) return;
      event.preventDefault();
      void selectFile(Number(event.target.dataset.file));
    };
    const sheetClick = event => {
      const tab = event.target.closest('[data-sheet]');
      if (tab) void selectSheet(Number(tab.dataset.sheet));
    };
    files.addEventListener('click', click);
    files.addEventListener('keydown', keydown);
    sheets.addEventListener('click', sheetClick);
    detach = () => {
      files.removeEventListener('click', click);
      files.removeEventListener('keydown', keydown);
      sheets.removeEventListener('click', sheetClick);
    };
    return detach;
  }
  return {loadPreview, selectFile, selectSheet, bind};
}
