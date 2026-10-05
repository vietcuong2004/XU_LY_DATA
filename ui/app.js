'use strict';
const $ = id => document.getElementById(id);
const state = {job:null, report:null, file:0, sheet:0, preview:null, filter:'all', search:'', input:null, template:null, useLocal:false, config:null, poll:null, cell:null, view:'upload', saving:false};
const fmt = new Intl.NumberFormat('vi-VN',{maximumFractionDigits:6});
const escapeHTML = value => String(value ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const display = value => value === null || value === undefined || value === '' ? '—' : typeof value === 'number' ? fmt.format(value) : String(value);
const cellDisplay = (cell,value=cell.value) => typeof value==='number' && cell.format==='General' ? String(value) : display(value);
const statusNames = {warning:'Cần kiểm tra',edited:'Đã chỉnh sửa',matched:'Khớp nguồn'};
const badge = status => `<span class="badge ${status}">${statusNames[status]}</span>`;
const normalize = text => String(text).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');

async function api(path, data){
  const response = await fetch(path, data === undefined ? {} : {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
  const result = await response.json();
  if(!response.ok) throw new Error(result.error || 'Không thể kết nối. Vui lòng thử lại.');
  return result;
}
let toastTimer;
function toast(message,error=false){$('toast').textContent=message;$('toast').className='show'+(error?' error':'');clearTimeout(toastTimer);toastTimer=setTimeout(()=>$('toast').className='',5500);}

function switchTab(tabId){
  const isUpload = tabId === 'upload';
  $('tab-btn-upload').classList.toggle('active', isUpload);
  $('tab-btn-upload').setAttribute('aria-selected', isUpload ? 'true' : 'false');
  $('tab-btn-review').classList.toggle('active', !isUpload);
  $('tab-btn-review').setAttribute('aria-selected', !isUpload ? 'true' : 'false');

  $('tab-content-upload').hidden = !isUpload;
  $('tab-content-review').hidden = isUpload;

  if(!isUpload){
    const hasReport = Boolean(state.report);
    $('view-review').hidden = !hasReport;
    $('results-empty').hidden = hasReport;
  }
}

$('tab-btn-upload').addEventListener('click', ()=>switchTab('upload'));
$('tab-btn-review').addEventListener('click', ()=>switchTab('review'));
const btnGoto = $('btn-goto-upload');
if(btnGoto) btnGoto.addEventListener('click', ()=>switchTab('upload'));

function setExportState(enabled, href=''){
  const btn = $('go-export');
  if(!btn) return;
  if(enabled && href){
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
if(btnExport){
  btnExport.addEventListener('click', e => {
    if(btnExport.getAttribute('aria-disabled') === 'true' || !btnExport.getAttribute('href')){
      e.preventDefault();
      e.stopPropagation();
      toast('Chưa hoàn tất tạo các file Family để xuất kết quả.', true);
    }
  });
}
setExportState(false);

function showView(view){
  state.view=view;
  const busy=view==='progress';
  $('view-upload').hidden=false;
  $('view-progress').hidden=!busy;
  const card=document.querySelector('.modern-card') || document.querySelector('.upload-card');
  if(card) card.inert=busy;
  const hist=$('history');
  if(hist) hist.inert=busy;
  $('process-button').disabled=busy || (!state.input && !state.useLocal);
  $('clear-input').disabled=busy;
  if(view==='review'){
    switchTab('review');
  } else if(view==='progress'||view==='upload'){
    switchTab('upload');
  }
}

function updateFileInfo(fileName){
  if(!fileName) return;
  const wkMatch = fileName.match(/(?:WK|Wk|wk)[-_ ]?(\d+)/i);
  const yrMatch = fileName.match(/(20\d\d)/);
  const weekNum = wkMatch ? Number(wkMatch[1]) : (Number($('report-week')?.value) || 40);
  const yearNum = yrMatch ? Number(yrMatch[1]) : 2026;
  
  if(wkMatch && $('report-week')) $('report-week').value = weekNum;
  
  if($('info-week')) $('info-week').textContent = `WK${weekNum}`;
  if($('info-year')) $('info-year').textContent = `${yearNum}`;
  if($('info-families')) $('info-families').textContent = '26';
  if($('info-destinations')) $('info-destinations').textContent = '78';
  if($('info-items')) $('info-items').textContent = '312';
  
  const startWk = $('start-week')?.value?.trim() || `${yearNum}/${weekNum + 2}`;
  const endWk = $('end-week')?.value?.trim() || `${yearNum + 1}/12`;
  if($('info-range')) $('info-range').textContent = `${startWk} – ${endWk}`;
}

function clearInputFile(e){
  if(e){e.preventDefault();e.stopPropagation();}
  state.input=null;
  state.useLocal=false;
  $('input-file').value='';
  $('input-title').textContent='Chưa chọn file';
  const pill=$('file-pill');
  if(pill) pill.hidden=true;
  $('drop-input').classList.remove('loaded');
  $('process-button').disabled=true;
  $('upload-error').hidden=true;
  
  if($('info-week')) $('info-week').textContent='—';
  if($('info-year')) $('info-year').textContent='—';
  if($('info-families')) $('info-families').textContent='—';
  if($('info-destinations')) $('info-destinations').textContent='—';
  if($('info-items')) $('info-items').textContent='—';
  if($('info-range')) $('info-range').textContent='—';
}
$('clear-input').addEventListener('click',clearInputFile);

function inputChanged(file){
  if(!file)return;
  if(!file.name.toLowerCase().endsWith('.xlsx'))return toast('Vui lòng chọn file Excel .xlsx.',true);
  if(file.size>45*1024*1024)return toast('File vượt quá 45 MB.',true);
  state.input=file;state.useLocal=false;
  $('input-title').textContent=file.name;
  const pill=$('file-pill');
  if(pill) pill.hidden=false;
  $('drop-input').classList.add('loaded');
  $('process-button').disabled=false;
  $('upload-error').hidden=true;
  updateFileInfo(file.name);
}
$('input-file').addEventListener('change',e=>inputChanged(e.target.files[0]));
$('drop-input').addEventListener('click',e=>{
  if(e.target.closest('#clear-input')||e.target.closest('#process-button')||e.target.closest('#file-pill'))return;
  if(!state.input&&!state.useLocal){$('input-file').click();}
});
$('template-file').addEventListener('change',e=>{
  const file=e.target.files[0];
  if(!file)return;
  if(!file.name.toLowerCase().endsWith('.xlsx')||file.size>45*1024*1024)return toast('Chọn file mẫu .xlsx tối đa 45 MB.',true);
  state.template=file;
  $('template-name').textContent=`${file.name} (mẫu riêng)`;
  if($('clear-template')) $('clear-template').hidden=false;
  toast('Đã chọn mẫu riêng thành công.');
});
const btnClearTpl = $('clear-template');
if(btnClearTpl){
  btnClearTpl.addEventListener('click',e=>{
    e.preventDefault();
    e.stopPropagation();
    state.template=null;
    $('template-file').value='';
    const defaultTpl=state.config?.template || 'BARBIE 2728 _Weekly shipment schedule 2728_WK39.xlsx';
    $('template-name').textContent=defaultTpl;
    btnClearTpl.hidden=true;
    toast('Đã khôi phục về mẫu mặc định.');
  });
}
for(const event of ['dragenter','dragover'])$('drop-input').addEventListener(event,e=>{e.preventDefault();$('drop-input').classList.add('dragging');});
for(const event of ['dragleave','drop'])$('drop-input').addEventListener(event,e=>{e.preventDefault();$('drop-input').classList.remove('dragging');});
$('drop-input').addEventListener('drop',e=>inputChanged(e.dataTransfer.files[0]));
const btnUseLocal = $('use-local');
if (btnUseLocal) {
  btnUseLocal.addEventListener('click',()=>{
    state.useLocal=true;state.input=null;
    $('input-title').textContent=state.config.local_input;
    const pill=$('file-pill');
    if(pill) pill.hidden=false;
    $('drop-input').classList.add('loaded');
    $('process-button').disabled=false;
    $('upload-error').hidden=true;
    updateFileInfo(state.config.local_input);
  });
}
let currentUploadRequest = null;
let isCancelled = false;

function cancelCurrentProcess(){
  isCancelled = true;
  if(currentUploadRequest){
    try { currentUploadRequest.abort(); } catch(e){}
    currentUploadRequest = null;
  }
  if(state.poll){
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
  if($('progress-percent')) $('progress-percent').textContent = '0%';
  if($('progress-eta')) $('progress-eta').textContent = 'Đang tính toán...';
  const track = document.querySelector('.progress-track');
  if(track) track.classList.remove('indeterminate');
  const confirmDialog = $('cancel-confirm-dialog');
  if(confirmDialog && confirmDialog.open) confirmDialog.close();
  toast('Đã hủy tải lên.');
}

const btnCancelProgress = $('btn-cancel-progress');
if(btnCancelProgress){
  btnCancelProgress.addEventListener('click', e => {
    e.preventDefault();
    e.stopPropagation();
    const dialog = $('cancel-confirm-dialog');
    if(dialog) dialog.showModal();
  });
}

const btnAbortDismiss = $('btn-abort-dismiss');
if(btnAbortDismiss){
  btnAbortDismiss.addEventListener('click', () => {
    const dialog = $('cancel-confirm-dialog');
    if(dialog) dialog.close();
  });
}

const btnAbortConfirm = $('btn-abort-confirm');
if(btnAbortConfirm){
  btnAbortConfirm.addEventListener('click', () => {
    cancelCurrentProcess();
  });
}

const cancelDialog = $('cancel-confirm-dialog');
if(cancelDialog){
  cancelDialog.addEventListener('click', e => {
    if(e.target === cancelDialog) cancelDialog.close();
  });
}

function uploadJob(options){
  const form=new FormData();
  form.append('options',JSON.stringify(options));
  if(state.input)form.append('input',state.input);
  if(state.template)form.append('template',state.template);
  return new Promise((resolve,reject)=>{
    const request=new XMLHttpRequest();
    currentUploadRequest=request;
    request.open('POST','/api/jobs');
    request.responseType='json';
    request.upload.onprogress=event=>{
      if(!event.lengthComputable || isCancelled)return;
      const percent=Math.round(event.loaded/event.total*100);
      $('progress-message').textContent=percent<100?`Đang tải file lên… ${percent}%`:'Đã gửi file. Đang chờ máy chủ tiếp nhận…';
      $('progress-count').textContent=`${(event.loaded/1024/1024).toFixed(1)} / ${(event.total/1024/1024).toFixed(1)} MB`;
    };
    request.onload=()=>{
      currentUploadRequest=null;
      if(isCancelled){reject(new Error('ABORTED'));return;}
      if(request.status>=200&&request.status<300&&request.response){resolve(request.response);return;}
      reject(new Error(request.response?.error || (request.status===413?'File vượt giới hạn tải lên của máy chủ.':'Máy chủ chưa tiếp nhận được file. Vui lòng thử lại.')));
    };
    request.onabort=()=>{
      currentUploadRequest=null;
      reject(new Error('ABORTED'));
    };
    request.onerror=()=>{
      currentUploadRequest=null;
      if(isCancelled){reject(new Error('ABORTED'));return;}
      reject(new Error('Mất kết nối khi tải file lên. Vui lòng thử lại.'));
    };
    request.send(form);
  });
}
$('process-button').addEventListener('click',async()=>{
  $('upload-error').hidden=true;
  isCancelled=false;
  try{
    const week=Number($('report-week').value),season=$('season').value.trim(),start=$('start-week').value.trim(),end=$('end-week').value.trim();
    if(!Number.isInteger(week)||week<1||week>53)throw new Error('Tuần báo cáo phải từ 1 đến 53.');
    if(!/^[a-zA-Z0-9_-]{1,20}$/.test(season))throw new Error('Mùa chỉ gồm chữ, số, dấu _ hoặc -.');
    if([start,end].some(v=>v&&!/^\d{4}\/\d{1,2}$/.test(v)))throw new Error('Khoảng tuần cần dạng YYYY/WW, ví dụ 2026/42.');
    if(!state.input&&!state.useLocal)throw new Error('Hãy chọn file nguồn.');
    $('process-button').disabled=true;
    setExportState(false);
    showView('progress');$('progress-message').textContent='Đang tải file lên…';$('progress-count').textContent='Khởi tạo';$('progress-fill').style.width='5%';
    if($('progress-percent')) $('progress-percent').textContent='0%';
    if($('progress-eta')) $('progress-eta').textContent='Đang ước tính...';
    $('progress-time').textContent='Đã chạy: 0 giây';
    const initialTrack=document.querySelector('.progress-track');
    if(initialTrack) initialTrack.classList.add('indeterminate');
    const payload={week,season,start,end,use_local:state.useLocal};
    const job=await uploadJob(payload);
    if(isCancelled) return;
    state.job=job;state.report=null;state.preview=null;
    await pollJob(job.id);
  }catch(error){
    if(error.message==='ABORTED'||isCancelled)return;
    showView('upload');$('upload-error').textContent=error.message;$('upload-error').hidden=false;
  }
  finally{if(!isCancelled)$('process-button').disabled=state.view==='progress';}
});
async function pollJob(id){
  if(isCancelled) return;
  clearTimeout(state.poll);
  const job=await api(`/api/jobs/${id}`);
  if(isCancelled) return;
  state.job=job;
  $('progress-message').textContent=job.message;
  const elapsedSeconds=Math.max(0,Math.floor((Date.now()-new Date(job.created))/1000));
  $('progress-time').textContent=`Đã chạy: ${elapsedSeconds} giây`;
  const hasProgress=job.total>0;
  const track=document.querySelector('.progress-track');
  if(track) track.classList.toggle('indeterminate',!hasProgress);
  const pct = hasProgress ? Math.min(100, Math.round((job.completed/job.total)*100)) : 0;
  $('progress-fill').style.width=hasProgress?`${(job.completed/job.total)*100}%`:'25%';
  if($('progress-percent')) $('progress-percent').textContent=hasProgress?`${pct}%`:'0%';
  $('progress-count').textContent=hasProgress?`(${job.completed} / ${job.total} Family)`:'Đọc và kiểm tra cấu trúc';
  if($('progress-eta')){
    if(hasProgress && job.completed > 0){
      const remaining=job.total - job.completed;
      if(remaining <= 0){
        $('progress-eta').textContent='Hoàn tất!';
      } else {
        const rate=job.completed / Math.max(1, elapsedSeconds);
        const etaSec=Math.max(1, Math.round(remaining / rate));
        if(etaSec < 60){
          $('progress-eta').textContent=`~${etaSec} giây`;
        } else {
          const m=Math.floor(etaSec / 60);
          const s=etaSec % 60;
          $('progress-eta').textContent=`~${m}p ${s}s`;
        }
      }
    } else {
      $('progress-eta').textContent='Đang ước tính...';
    }
  }
  if(job.status==='ready'){await openJob(id);await loadConfig();return;}
  if(job.status==='error'){setExportState(false);showView('upload');$('upload-error').textContent=job.message;$('upload-error').hidden=false;await loadConfig();return;}
  if(!isCancelled){
    state.poll=setTimeout(()=>pollJob(id).catch(error=>{if(!isCancelled){toast(error.message,true);showView('upload');}}),1200);
  }
}
async function loadConfig(){
  state.config=await api('/api/config');
  const defaultTpl = state.config?.template || 'templates/BARBIE 2728 _Weekly shipment schedule 2728_WK39.xlsx';
  if(!state.template){
    $('template-name').textContent=defaultTpl;
    if($('clear-template')) $('clear-template').hidden=true;
  }
  if($('use-local')) $('use-local').hidden=!state.config.local_input;
  if($('local-name')) $('local-name').textContent=state.config.local_input || '';
  $('history').innerHTML=state.config.jobs.length?state.config.jobs.map(job=>`<button class="history-item" data-job="${job.id}"><strong>WK${String(job.week).padStart(2,'0')} · ${escapeHTML(job.source)}</strong><small>${new Date(job.created).toLocaleDateString('vi-VN')} · ${job.status==='ready'?`${job.total} Family`:job.status==='error'?'Bị gián đoạn':'Đang xử lý'}</small></button>`).join(''):'<p class="muted">Chưa có phiên xử lý.</p>';
}
$('history').addEventListener('click',async e=>{const button=e.target.closest('[data-job]');if(!button)return;try{const job=await api(`/api/jobs/${button.dataset.job}`);if(job.status==='ready')await openJob(job.id);else if(job.status==='error')toast(job.message,true);else{showView('progress');await pollJob(job.id);}}catch(error){toast(error.message,true);}});
async function openJob(id){
  clearTimeout(state.poll);
  state.job=await api(`/api/jobs/${id}`);
  state.report=await api(`/api/jobs/${id}/report`);
  state.filter='all';state.search='';$('family-search').value='';
  state.file=state.report.families.find(f=>f.name==='BARBIE 2728')?.id ?? 0;state.sheet=0;
  if($('source-caption')) $('source-caption').textContent=state.job.source;
  setExportState(true, `/api/jobs/${state.job.id}/zip`);
  document.querySelectorAll('[data-filter]').forEach(e=>e.classList.toggle('active',e.dataset.filter==='all'));
  renderMetrics();renderFamilies();showView('review');await loadPreview();
  history.replaceState(null,'',`?job=${id}`);
  $('view-review').scrollIntoView({behavior:'smooth',block:'start'});
}
function renderMetrics(){
  const families=state.report.families;
  const warnings=families.filter(f=>f.problems.length).length;
  const edited=families.filter(f=>f.changed_cells>0).length;
  const metrics=[[families.length,'File Family đã tạo','▦',''],[state.report.weeks,'Tuần trong kế hoạch','▤',''],[warnings,'Family cần kiểm tra','◎','amber'],[edited,'Family đã chỉnh sửa','✎','blue']];
  if($('metrics')) $('metrics').innerHTML=metrics.map(([value,label,icon,color])=>`<div class="metric"><div><strong>${value}</strong><label>${label}</label></div><span class="metric-icon ${color}">${icon}</span></div>`).join('');
  const badge = $('tab-review-badge');
  if(badge){
    badge.textContent = `${families.length} Family`;
    badge.hidden = false;
  }
}
function renderFamilies(){
  const list=state.report.families.filter(f=>normalize(f.name).includes(normalize(state.search))&&(state.filter==='all'||(state.filter==='warning'?f.problems.length:f.changed_cells>0)));
  $('family-count').textContent=state.report.families.length;
  $('family-list').innerHTML=list.length?list.map(f=>`<button class="family-item ${f.id===state.file?'selected':''}" data-file="${f.id}"><span class="family-item-name"><span class="status-dot ${f.status_label}"></span>${escapeHTML(f.name)}</span><small>${f.markets} thị trường · ${f.items} item${f.problems.length?' · cần xem':''}</small></button>`).join(''):'<div class="empty-state">Không tìm thấy Family.</div>';
}
$('family-search').addEventListener('input',e=>{state.search=e.target.value;renderFamilies();});
document.querySelectorAll('[data-filter]').forEach(button=>button.addEventListener('click',()=>{state.filter=button.dataset.filter;document.querySelectorAll('[data-filter]').forEach(el=>el.classList.toggle('active',el===button));renderFamilies();}));
$('family-list').addEventListener('click',async e=>{const button=e.target.closest('[data-file]');if(!button)return;state.file=Number(button.dataset.file);state.sheet=0;renderFamilies();await loadPreview();});
let previewRequest=0;
async function loadPreview(focusAddress){
  const request=++previewRequest;
  $('grid-container').innerHTML='<div class="empty-state">Đang mở bảng dữ liệu…</div>';
  try{
    const result=await api(`/api/jobs/${state.job.id}/preview?file=${state.file}&sheet=${state.sheet}`);
    if(request!==previewRequest)return;
    state.preview=result;
    $('workbook-name').textContent=result.file.name;
    $('workbook-meta').textContent=`${result.file.items} item · ${result.file.markets} thị trường · ${result.file.source_checks ? fmt.format(result.file.source_checks)+' đối chiếu nguồn' : 'Lịch nguồn: '+state.report.first_week+' → '+state.report.last_week}`;
    $('download-one').href=`/api/jobs/${state.job.id}/download?file=${state.file}`;
    if($('excel-preview-download')) $('excel-preview-download').href=$('download-one').href;
    $('sheet-tabs').innerHTML=result.sheets.map((name,index)=>`<button role="tab" aria-selected="${index===state.sheet}" class="sheet-tab ${index===state.sheet?'selected':''}" data-sheet="${index}">${index===0?'▦ ':index===1?'▤ ':'▥ '}${escapeHTML(name)}</button>`).join('');
    renderGrid();renderProblems();renderAudit();
    if(focusAddress){focusCell(focusAddress);openCell(focusAddress);}
  }catch(error){$('grid-container').innerHTML=`<div class="empty-state">${escapeHTML(error.message)}</div>`;toast(error.message,true);}
}
$('sheet-tabs').addEventListener('click',async e=>{const button=e.target.closest('[data-sheet]');if(!button)return;state.sheet=Number(button.dataset.sheet);await loadPreview();});
function colIndex(text){return [...text].reduce((n,c)=>n*26+c.charCodeAt(0)-64,0)-1;}
function coords(address){const m=/^([A-Z]+)(\d+)$/.exec(address);return{col:colIndex(m[1]),row:Number(m[2])-1};}
function renderGrid(){
  const data=state.preview,covered=new Set(),merges=new Map();
  for(const merge of data.merges){const [start,end]=merge.split(':').map(coords);merges.set(`${start.row},${start.col}`,{rows:end.row-start.row+1,cols:end.col-start.col+1});for(let r=start.row;r<=end.row;r++)for(let c=start.col;c<=end.col;c++)if(r!==start.row||c!==start.col)covered.add(`${r},${c}`);}
  let html='<table class="sheet-grid" aria-label="'+escapeHTML(data.sheet)+'"><colgroup><col class="index-col">'+data.columns.map((c,i)=>`<col class="${i===0?'label-col':'value-col'}">`).join('')+'</colgroup><thead><tr><th></th>'+data.columns.map(c=>`<th scope="col">${c}</th>`).join('')+'</tr></thead><tbody>';
  data.rows.forEach((row,r)=>{
    html+=`<tr><td class="row-number">${r+1}</td>`;
    row.forEach((cell,c)=>{
      if(covered.has(`${r},${c}`))return;
      const merge=merges.get(`${r},${c}`);
      const header=(state.sheet===0&&r>=1&&r<8)||(state.sheet===1&&r>=8&&r<13);
      const classes=['data-cell',c===0?'label-cell':'',header?'header-cell':'',typeof cell.value==='number'?'numeric':'',cell.changed?'changed':'',cell.error?'error-cell':'',cell.formula?'has-formula':''];
      const text=cell.value===null?'':cellDisplay(cell);
      const inlineStyles=[];
      if(cell.bg && cell.bg!=='#ffffff') inlineStyles.push(`background-color:${cell.bg}`);
      if(cell.fg) inlineStyles.push(`color:${cell.fg}`);
      const styleAttr=inlineStyles.length?` style="${inlineStyles.join(';')}"`:'';
      html+=`<td tabindex="0" role="button" data-address="${cell.address}" aria-label="${cell.address}: ${escapeHTML(text||'trống')}" class="${classes.join(' ')}"${styleAttr} ${merge?`rowspan="${merge.rows}" colspan="${merge.cols}"`:''} title="${escapeHTML(cell.address+' · '+(cell.editable?'Nhấn để sửa':cell.formula?'Công thức tự tính':'Xem đối chiếu'))}">${escapeHTML(text)}</td>`;
    });
    html+='</tr>';
  });
  $('grid-container').innerHTML=html+'</tbody></table>';
  $('grid-info').textContent=`${data.rows.length} dòng · ${data.columns.length} cột · Đã lưu${data.revision?' · Lần sửa '+data.revision:''}`;
}

let previewModalSheet = 0;
async function openExcelPreview(sheetIdx){
  if(sheetIdx === undefined) previewModalSheet = state.sheet;
  else previewModalSheet = sheetIdx;
  const dialog = $('excel-preview-dialog');
  if(!dialog) return;
  if(!state.preview || !state.preview.file){
    toast('Chưa có dữ liệu để xem trước.', true);
    return;
  }
  $('excel-preview-title').textContent = `Xem trước: ${state.preview.file.name}.xlsx`;
  if($('excel-preview-download')) $('excel-preview-download').href = `/api/jobs/${state.job.id}/download?file=${state.file}`;
  
  $('excel-preview-tabs').innerHTML = state.preview.sheets.map((name, idx)=>
    `<button role="tab" aria-selected="${idx===previewModalSheet}" class="sheet-tab ${idx===previewModalSheet?'selected':''}" data-prev-sheet="${idx}">
       ${idx===0?'▦ ':idx===1?'▤ ':'▥ '}${escapeHTML(name)}
     </button>`
  ).join('');

  dialog.showModal();
  await renderExcelWysiwygGrid(previewModalSheet);
}

async function renderExcelWysiwygGrid(sheetIdx){
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

if($('btn-preview-excel')) $('btn-preview-excel').addEventListener('click', () => openExcelPreview());
if($('close-excel-preview')) $('close-excel-preview').addEventListener('click', () => $('excel-preview-dialog').close());
if($('excel-preview-dialog')) $('excel-preview-dialog').addEventListener('click', e => {
  if (e.target === $('excel-preview-dialog')) $('excel-preview-dialog').close();
});
if($('excel-preview-tabs')) $('excel-preview-tabs').addEventListener('click', async e => {
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
$('grid-container').addEventListener('click',e=>{const cell=e.target.closest('[data-address]');if(cell)openCell(cell.dataset.address);});
$('grid-container').addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){const cell=e.target.closest('[data-address]');if(cell){e.preventDefault();openCell(cell.dataset.address);}}});
function focusCell(address){const element=$('grid-container').querySelector(`[data-address="${address}"]`);if(!element)return;document.querySelectorAll('.selected-cell').forEach(el=>el.classList.remove('selected-cell'));element.classList.add('selected-cell');element.scrollIntoView({block:'center',inline:'center',behavior:'smooth'});}
function openCell(address){
  const {row,col}=coords(address),cell=state.preview.rows[row]?.[col];if(!cell)return;
  state.cell=cell;focusCell(address);
  $('cell-heading').textContent=`${state.preview.sheet.trim()} · ${address}`;
  let context=cell.comment;
  if(!context){const week=row>=13&&state.sheet===0?state.preview.rows[row][0].value:row>=18&&state.sheet===1?state.preview.rows[row][0].value:null;context=week?`Tuần ${week} · ${state.preview.file.name}`:state.preview.file.name;}
  if(state.sheet===0&&col>0&&row>=8)context=`${state.preview.rows[4][col]?.value||''} · ${state.preview.rows[5][col]?.value||''}\n${context}`;
  $('cell-context').textContent=context;
  $('original-value').textContent=cellDisplay(cell,cell.original);$('current-value').textContent=cellDisplay(cell);
  $('formula-info').hidden=!cell.formula;$('formula-info').textContent=cell.formula || '';
  $('edit-fields').hidden=!cell.editable;$('readonly-info').hidden=Boolean(cell.editable);
  $('readonly-info').textContent=cell.formula?'Ô này được tính tự động. Sửa dữ liệu gốc trong Breakdown để cập nhật các tổng liên quan.':state.sheet===2?'Ngày được cập nhật theo tuần trong Breakdown. Hãy sửa nhãn tuần ở sheet Breakdown.':'Ô thuộc cấu trúc hoặc dữ liệu tổng hợp. Chỉnh số lượng, phiên bản, destination và capsule tại Breakdown; mã item/MPG giữ cố định để không trộn nhóm.';
  const isDate=cell.editable?.startsWith('date');
  $('new-value').type=isDate?'date':'text';$('new-value').inputMode=cell.editable==='number'?'decimal':'text';
  $('new-value').value=cell.error?'':cell.value??'';
  $('new-value').placeholder=cell.editable==='week'?'YYYY/WW':cell.editable==='number'?'Ví dụ: 202.8':'Nhập giá trị mới';
  $('edit-hint').textContent=cell.editable==='week'?'Đổi tuần sẽ cập nhật ngày thứ Hai và Release Qty trong file này.':cell.editable==='number'?'Dùng dấu chấm thập phân. Để trống nếu chưa có số liệu; các tổng tự tính lại.':'Thay đổi chỉ áp dụng cho file kết quả này, không sửa file nguồn.';
  $('edit-reason').value='';$('edit-error').hidden=true;
  $('save-cell').hidden=!cell.editable;$('restore-cell').hidden=!cell.editable||!cell.changed;
  $('trace-cell').hidden=!cell.formula&&state.sheet!==2;
  $('cell-dialog').showModal();if(cell.editable)setTimeout(()=>$('new-value').focus(),70);
}
$('close-dialog').addEventListener('click',()=>$('cell-dialog').close());
$('cell-dialog').addEventListener('click',e=>{if(e.target===$('cell-dialog')){const rect=$('cell-dialog').getBoundingClientRect();if(e.clientX<rect.left||e.clientX>rect.right||e.clientY<rect.top||e.clientY>rect.bottom)$('cell-dialog').close();}});
$('edit-form').addEventListener('submit',e=>{e.preventDefault();saveCell(false);});
$('restore-cell').addEventListener('click',()=>saveCell(true));
async function saveCell(restore){
  if(state.saving)return;state.saving=true;$('save-cell').disabled=true;$('restore-cell').disabled=true;$('edit-error').hidden=true;
  const address=state.cell.address;
  try{
    const entry=await api(`/api/jobs/${state.job.id}/edit/${state.file}`,{sheet:state.sheet,cell:address,revision:state.preview.revision,value:$('new-value').value,reason:$('edit-reason').value,restore});
    state.report.families[entry.id]=entry;$('cell-dialog').close();renderMetrics();renderFamilies();await loadPreview();focusCell(address);toast(restore?'Đã khôi phục giá trị gốc và tính lại các tổng.':'Đã lưu chỉnh sửa và cập nhật các tổng liên quan.');
  }catch(error){$('edit-error').textContent=error.message;$('edit-error').hidden=false;}
  finally{state.saving=false;$('save-cell').disabled=false;$('restore-cell').disabled=false;}
}
$('trace-cell').addEventListener('click',async()=>{
  const cell=state.cell;
  let address;
  if(state.sheet===2)address=`A${coords(cell.address).row+12}`;
  else{
    const refs=[...(cell.formula||'').matchAll(/'Breakdown '!([A-Z]+\d+)/g)].map(m=>m[1]);
    address=refs[0];
    if(!address){const r=coords(cell.address).row+1;address=state.preview.file.problems.find(p=>p.sheet==='Breakdown '&&p.kind==='error'&&coords(p.cell).col<=state.preview.file.items)?.cell || `B${r>=19?r-5:9}`;}
  }
  $('cell-dialog').close();state.sheet=0;await loadPreview(address);
});
function renderProblems(){
  const entry=state.report.families[state.file],list=entry.problems;
  // Input errors first; dependent formula errors remain visible below.
  const sorted=[...list].sort((a,b)=>(a.sheet==='Breakdown '?-1:1)-(b.sheet==='Breakdown '?-1:1));
  $('problem-count').textContent=list.length;
  $('jump-warning').hidden=!list.length;
  $('problems').innerHTML=sorted.length?sorted.map(p=>`<button class="problem-row" data-problem-sheet="${escapeHTML(p.sheet)}" data-problem-cell="${p.cell}"><span class="problem-icon">!</span><div>${escapeHTML(p.message)}<small>${escapeHTML(p.sheet)} · ${p.cell}${p.kind==='week'?' · cần xác nhận lại tuần':''}</small></div><span>↗</span></button>`).join(''):'<div class="empty-state">✓ Không có ô lỗi hoặc tuần ISO không hợp lệ.<br>Đối chiếu với nguồn không thay thế xác nhận nghiệp vụ.</div>';
}
$('problems').addEventListener('click',async e=>{const button=e.target.closest('[data-problem-cell]');if(!button)return;state.sheet=state.preview.sheets.indexOf(button.dataset.problemSheet);await loadPreview(button.dataset.problemCell);});
$('jump-warning').addEventListener('click',()=>{$('problems').scrollIntoView({behavior:'smooth',block:'center'});});
function renderAudit(){
  const edits=state.report.families[state.file].edits;
  $('audit-list').innerHTML=edits.length?[...edits].reverse().map(e=>`<div class="audit-row"><b>${escapeHTML(e.sheet)} · ${e.cell}</b><div class="audit-values"><del>${escapeHTML(display(e.before))}</del><span>→</span><span>${escapeHTML(display(e.after))}</span></div><small>${new Date(e.time).toLocaleString('vi-VN')}${e.restored?' · Khôi phục':''}${e.reason?' · '+escapeHTML(e.reason):''}</small></div>`).join(''):'<div class="empty-state">Chưa có chỉnh sửa.<br>Mỗi thay đổi sẽ được lưu lại tại đây.</div>';
}
loadConfig().then(async()=>{
  const params=new URLSearchParams(location.search),job=params.get('job');
  if(job&&/^[a-f0-9]{32}$/.test(job)){
    await openJob(job);
    const cell=params.get('cell');
    if(cell&&/^[A-Z]+[1-9][0-9]*$/.test(cell))openCell(cell);
  }
}).catch(error=>toast('Không kết nối được ứng dụng: '+error.message,true));
