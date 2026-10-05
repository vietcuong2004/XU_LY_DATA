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

function showView(view){
  state.view=view;
  const busy=view==='progress';
  $('view-upload').hidden=false;
  $('view-progress').hidden=!busy;
  document.querySelector('.upload-card').inert=busy;
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

function clearInputFile(e){
  if(e){e.preventDefault();e.stopPropagation();}
  state.input=null;
  state.useLocal=false;
  $('input-file').value='';
  $('input-title').textContent='Kéo file Excel vào đây';
  $('input-subtitle').textContent='hoặc nhấn để chọn file từ máy tính · tối đa 45 MB';
  $('drop-input').classList.remove('loaded');
  $('choose-file-btn').hidden=false;
  $('process-button').hidden=true;
  $('process-button').disabled=true;
  $('upload-error').hidden=true;
}
$('clear-input').addEventListener('click',clearInputFile);

function inputChanged(file){
  if(!file)return;
  if(!file.name.toLowerCase().endsWith('.xlsx'))return toast('Vui lòng chọn file Excel .xlsx.',true);
  if(file.size>45*1024*1024)return toast('File vượt quá 45 MB.',true);
  state.input=file;state.useLocal=false;
  $('input-title').textContent=file.name;
  $('input-subtitle').textContent=`${(file.size/1024/1024).toFixed(2)} MB · Đã chọn file nguồn`;
  $('drop-input').classList.add('loaded');
  $('choose-file-btn').hidden=true;
  $('process-button').hidden=false;
  $('process-button').disabled=false;
  $('upload-error').hidden=true;
}
$('input-file').addEventListener('change',e=>inputChanged(e.target.files[0]));
$('drop-input').addEventListener('click',e=>{
  if(e.target.closest('#clear-input')||e.target.closest('#process-button'))return;
  if(!state.input&&!state.useLocal){$('input-file').click();}
});
$('template-file').addEventListener('change',e=>{const file=e.target.files[0];if(!file)return;if(!file.name.toLowerCase().endsWith('.xlsx')||file.size>45*1024*1024)return toast('Chọn file mẫu .xlsx tối đa 45 MB.',true);state.template=file;$('template-name').textContent=file.name;});
for(const event of ['dragenter','dragover'])$('drop-input').addEventListener(event,e=>{e.preventDefault();$('drop-input').classList.add('dragging');});
for(const event of ['dragleave','drop'])$('drop-input').addEventListener(event,e=>{e.preventDefault();$('drop-input').classList.remove('dragging');});
$('drop-input').addEventListener('drop',e=>inputChanged(e.dataTransfer.files[0]));
$('use-local').addEventListener('click',()=>{
  state.useLocal=true;state.input=null;
  $('input-title').textContent=state.config.local_input;
  $('input-subtitle').textContent='Đã chọn file có sẵn trong thư mục làm việc';
  $('drop-input').classList.add('loaded');
  $('choose-file-btn').hidden=true;
  $('process-button').hidden=false;
  $('process-button').disabled=false;
  $('upload-error').hidden=true;
});
function readFile(file){return new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve({name:file.name,data:reader.result.split(',')[1]});reader.onerror=()=>reject(new Error('Không đọc được file.'));reader.readAsDataURL(file);});}
$('process-button').addEventListener('click',async()=>{
  $('upload-error').hidden=true;
  try{
    const week=Number($('report-week').value),season=$('season').value.trim(),start=$('start-week').value.trim(),end=$('end-week').value.trim();
    if(!Number.isInteger(week)||week<1||week>53)throw new Error('Tuần báo cáo phải từ 1 đến 53.');
    if(!/^[a-zA-Z0-9_-]{1,20}$/.test(season))throw new Error('Mùa chỉ gồm chữ, số, dấu _ hoặc -.');
    if([start,end].some(v=>v&&!/^\d{4}\/\d{1,2}$/.test(v)))throw new Error('Khoảng tuần cần dạng YYYY/WW, ví dụ 2026/42.');
    if(!state.input&&!state.useLocal)throw new Error('Hãy chọn file nguồn.');
    if(!state.template&&!state.config.template)throw new Error('Hãy chọn file mẫu BARBIE.');
    $('process-button').disabled=true;
    showView('progress');$('progress-message').textContent='Đang tải file lên…';$('progress-count').textContent='Chuẩn bị file';$('progress-fill').style.width='4%';
    document.querySelector('.progress-track').classList.add('indeterminate');
    const payload={week,season,start,end,use_local:state.useLocal};
    if(state.input)payload.input=await readFile(state.input);
    if(state.template)payload.template=await readFile(state.template);
    const job=await api('/api/jobs',payload);
    state.job=job;state.report=null;state.preview=null;
    await pollJob(job.id);
  }catch(error){showView('upload');$('upload-error').textContent=error.message;$('upload-error').hidden=false;}
  finally{$('process-button').disabled=state.view==='progress';}
});
async function pollJob(id){
  clearTimeout(state.poll);
  const job=await api(`/api/jobs/${id}`);state.job=job;
  $('progress-message').textContent=job.message;
  $('progress-time').textContent=`${Math.max(0,Math.floor((Date.now()-new Date(job.created))/1000))} giây`;
  const hasProgress=job.total>0;
  document.querySelector('.progress-track').classList.toggle('indeterminate',!hasProgress);
  $('progress-fill').style.width=hasProgress?`${job.completed/job.total*100}%`:'25%';
  $('progress-count').textContent=hasProgress?`${job.completed} / ${job.total} Family`:'Đọc và kiểm tra cấu trúc';
  if(job.status==='ready'){await openJob(id);await loadConfig();return;}
  if(job.status==='error'){showView('upload');$('upload-error').textContent=job.message;$('upload-error').hidden=false;await loadConfig();return;}
  state.poll=setTimeout(()=>pollJob(id).catch(error=>{toast(error.message,true);showView('upload');}),1200);
}
async function loadConfig(){
  state.config=await api('/api/config');
  if(!state.template)$('template-name').textContent=state.config.template || 'Chưa có mẫu mặc định. Vui lòng chọn file mẫu.';
  $('use-local').hidden=!state.config.local_input;$('local-name').textContent=state.config.local_input || '';
  $('history').innerHTML=state.config.jobs.length?state.config.jobs.map(job=>`<button class="history-item" data-job="${job.id}"><strong>WK${String(job.week).padStart(2,'0')} · ${escapeHTML(job.source)}</strong><small>${new Date(job.created).toLocaleDateString('vi-VN')} · ${job.status==='ready'?`${job.total} Family`:job.status==='error'?'Bị gián đoạn':'Đang xử lý'}</small></button>`).join(''):'<p class="muted">Chưa có phiên xử lý.</p>';
}
$('history').addEventListener('click',async e=>{const button=e.target.closest('[data-job]');if(!button)return;try{const job=await api(`/api/jobs/${button.dataset.job}`);if(job.status==='ready')await openJob(job.id);else if(job.status==='error')toast(job.message,true);else{showView('progress');await pollJob(job.id);}}catch(error){toast(error.message,true);}});
async function openJob(id){
  clearTimeout(state.poll);
  state.job=await api(`/api/jobs/${id}`);
  state.report=await api(`/api/jobs/${id}/report`);
  state.filter='all';state.search='';$('family-search').value='';
  state.file=state.report.families.find(f=>f.name==='BARBIE 2728')?.id ?? 0;state.sheet=0;
  $('source-caption').textContent=state.job.source;
  $('go-export').href=`/api/jobs/${state.job.id}/zip`;
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
  $('metrics').innerHTML=metrics.map(([value,label,icon,color])=>`<div class="metric"><div><strong>${value}</strong><label>${label}</label></div><span class="metric-icon ${color}">${icon}</span></div>`).join('');
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
  data.rows.forEach((row,r)=>{html+=`<tr><td class="row-number">${r+1}</td>`;row.forEach((cell,c)=>{if(covered.has(`${r},${c}`))return;const merge=merges.get(`${r},${c}`);const header=(state.sheet===0&&r>=1&&r<8)||(state.sheet===1&&r>=8&&r<13);const classes=['data-cell',c===0?'label-cell':'',header?'header-cell':'',typeof cell.value==='number'?'numeric':'',cell.changed?'changed':'',cell.error?'error-cell':'',cell.formula?'has-formula':''];const text=cell.value===null?'':cellDisplay(cell);html+=`<td tabindex="0" role="button" data-address="${cell.address}" aria-label="${cell.address}: ${escapeHTML(text||'trống')}" class="${classes.join(' ')}" ${merge?`rowspan="${merge.rows}" colspan="${merge.cols}"`:''} title="${escapeHTML(cell.address+' · '+(cell.editable?'Nhấn để sửa':cell.formula?'Công thức tự tính':'Xem đối chiếu'))}">${escapeHTML(text)}</td>`;});html+='</tr>';});
  $('grid-container').innerHTML=html+'</tbody></table>';
  $('grid-info').textContent=`${data.rows.length} dòng · ${data.columns.length} cột · Đã lưu${data.revision?' · Lần sửa '+data.revision:''}`;
}
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
