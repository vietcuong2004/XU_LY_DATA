const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const path=require('node:path');

function editor(){
  const text=fs.readFileSync(path.join(__dirname,'../ui/app.js'),'utf8');
  const code=text.slice(text.indexOf('const workbookDrafts ='),text.indexOf('function getProblemSuggestion(p)'));
  const elements=new Map();
  const element=id=>{
    if(!elements.has(id))elements.set(id,{events:{},classList:{toggle(){}},addEventListener(name,fn){this.events[name]=fn;}});
    return elements.get(id);
  };
  const requests=[];
  const state={job:{id:'original'},file:0,sheet:0,preview:{revision:3},saving:false};
  const ctx=vm.createContext({state,$:element,requests,
    document:{activeElement:null,querySelector:()=>element('panel'),addEventListener(){}},
    window:{addEventListener(){}},toast(){},renderGrid(){},renderFamilies(){},
    async api(url,body){requests.push({url,body});return {new_job_id:'saved'};},
    async openJob(id){state.job={id};},async loadPreview(){},async loadConfig(){}
  });
  vm.runInContext(code,ctx);
  element('edit-workbook').events.click();
  return {ctx,state,element,requests,run:code=>vm.runInContext(code,ctx)};
}

test('changes across sheets stay local until a single Save request',async()=>{
  const e=editor();
  e.ctx.stageCell({address:'B9',editable:'number',value:1},'12');
  e.state.sheet=1;
  e.ctx.stageCell({address:'G19',editable:'text',value:'old'},'new');
  assert.equal(e.requests.length,0);
  assert.equal(e.run('currentDraft().edits.size'),2);
  await e.ctx.saveWorkbook();
  assert.equal(e.requests.length,1);
  assert.equal(e.requests[0].body.edits.length,2);
  assert.equal(e.requests[0].body.revision,3);
  assert.equal(e.requests[0].body.edits[0].sheet,0);
  assert.equal(e.requests[0].body.edits[1].sheet,1);
  assert.equal(e.state.job.id,'saved');
});

test('cancel and reverting a cell do not write to the server',()=>{
  const e=editor(),cell={address:'B9',editable:'number',value:1};
  e.ctx.stageCell(cell,'12');e.ctx.stageCell(cell,'1');
  assert.equal(e.run('currentDraft().edits.size'),0);
  e.ctx.stageCell(cell,'9');
  e.element('cancel-workbook').events.click();
  assert.equal(e.run('currentDraft()'),undefined);
  assert.equal(e.requests.length,0);
});

test('failed save retains drafts for correction and retry',async()=>{
  const e=editor();
  e.ctx.stageCell({address:'B9',editable:'number',value:1},'invalid');
  e.ctx.api=async()=>{throw new Error('Invalid number');};
  await e.ctx.saveWorkbook();
  assert.equal(e.run('currentDraft().edits.size'),1);
  assert.equal(e.state.job.id,'original');
  assert.equal(e.state.saving,false);
});

test('source formula, literal text, and blank are distinct',()=>{
  const e=editor();e.state.file=-1;e.state.preview.source_editing={available:true};
  e.element('edit-workbook').events.click();
  const cell={address:'A1',editable:'source_cell',value:1};
  e.ctx.stageCell(cell,'=SUM(1,2)');
  assert.equal(e.run("pendingEdit(0,'A1').value_type"),'formula');
  e.ctx.stageCell(cell,"'=SUM(1,2)");
  assert.equal(e.run("pendingEdit(0,'A1').value_type"),'text');
  assert.equal(e.run("pendingEdit(0,'A1').value"),'=SUM(1,2)');
  e.ctx.stageCell(cell,'');
  assert.equal(e.run("pendingEdit(0,'A1').value_type"),'blank');
});
