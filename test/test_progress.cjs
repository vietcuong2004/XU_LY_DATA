const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const code=fs.readFileSync(require('node:path').join(__dirname,'../ui/app.js'),'utf8');
function harness(){
  const nodes=new Map();let now=1000,tick=null;
  const $=id=>{if(!nodes.has(id))nodes.set(id,{style:{},classList:{toggle(){}},scrollIntoView(){}});return nodes.get(id);};
  const ctx=vm.createContext({$,state:{input:{}},Date:{now:()=>now},
    setInterval:fn=>{tick=fn;return 1;},clearInterval:()=>{tick=null;},clearTimeout(){},
    document:{querySelector:()=>$('track'),querySelectorAll:()=>[]},
    switchTab:tab=>{$('tab').textContent=tab;},history:{replaceState(){}},
    saveFileState(){},setExportState(){},renderMetrics(){},renderFamilies(){},updateUploadStateUI(){},
    loadPreview:async()=>{},api:async url=>url.endsWith('/report')?{families:[{id:0}]}:{id:'job'}
  });
  vm.runInContext(code.slice(code.indexOf('let currentUploadRequest ='),code.indexOf('function cancelCurrentProcess')),ctx);
  vm.runInContext(code.slice(code.indexOf('function showView('),code.indexOf('function outputOptions(')),ctx);
  vm.runInContext(code.slice(code.indexOf('async function openJob('),code.indexOf('function renderMetrics(')),ctx);
  return {ctx,$,advance(ms){now+=ms;if(tick)tick();},hasClock:()=>!!tick};
}
test('clock ticks without network responses and stops after exiting progress',()=>{
  const h=harness();h.ctx.startProgressClock();h.advance(5000);
  assert.equal(h.$('progress-time').textContent,'Đã chạy: 5 giây');
  h.ctx.showView('upload');assert.equal(h.hasClock(),false);
});
test('stream decoder preserves split JSON and emits each record once',()=>{
  const h=harness(),events=[];const decode=h.ctx.progressDecoder(j=>events.push(j));
  decode('{"status":"pro');decode('{"status":"processing"}\n{"status":');
  decode('{"status":"processing"}\n{"status":"ready"}',true);
  assert.deepEqual(events.map(j=>j.status),['processing','ready']);
});
test('100 percent is reserved for ready after source comparison',()=>{
  const h=harness();const job={total:26,completed:26,status:'processing'};
  h.ctx.renderJobProgress(job);assert.equal(h.$('progress-percent').textContent,'80%');
  h.ctx.renderJobProgress({...job,phase:'validate',checked:26});
  assert.equal(h.$('progress-percent').textContent,'99%');
  h.ctx.renderJobProgress({...job,status:'ready'});
  assert.equal(h.$('progress-percent').textContent,'100%');
});
test('opening completed results closes overlay and releases inert controls',async()=>{
  const h=harness();h.ctx.showView('progress');h.ctx.startProgressClock();
  assert.equal(h.$('view-progress').hidden,false);
  await h.ctx.openJob('job','review');
  assert.equal(h.$('view-progress').hidden,true);
  assert.equal(h.$('card-unloaded').inert,false);
  assert.equal(h.$('tab').textContent,'review');assert.equal(h.hasClock(),false);
});
test('ready processing automatically requests the review tab',async()=>{
  const h=harness();let opened;
  h.ctx.openJob=async(id,tab)=>{opened=tab;};h.ctx.loadConfig=async()=>{};
  vm.runInContext(code.slice(code.indexOf('async function pollJob('),code.indexOf('async function loadConfig(')),h.ctx);
  await h.ctx.pollJob('job',{id:'job',status:'ready',total:1,completed:1});
  assert.equal(opened,'review');
});
