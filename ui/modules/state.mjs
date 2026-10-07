export function createWorkspaceState() {
  return { job: null, report: null, file: 0, sheet: 0, preview: null, filter: 'all', search: '', input: null, useLocal: false, config: null, poll: null, cell: null, view: 'upload', saving: false, colWidths: {}, choosingReplacement: false };
}
