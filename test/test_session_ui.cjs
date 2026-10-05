const {test}=require('node:test');
const assert=require('node:assert/strict');
const vm=require('node:vm');
const fs=require('node:fs');
const path=require('node:path');
const source=fs.readFileSync(path.join(__dirname,'../ui/app.js'),'utf8');
function setup(error){
  const nodes={};let url,view;
  const ctx=vm.createContext({URLSearchParams,state:{job:{id:'old'}},
    $:id=>nodes[id]??(nodes[id]={}),loadConfig:async()=>{},openJob:async()=>{throw error;},
    location:{pathname:'/',search:'?job='+'a'.repeat(32)+'&cell=A1&other=kept'},
    history:{replaceState:(_a,_b,next)=>url=next},setExportState(){},showView:v=>view=v
  });
  vm.runInContext(source.slice(source.indexOf('async function initializeWorkspace(){'),source.indexOf('\ninitializeWorkspace().catch')),ctx);
  return {ctx,nodes,url:()=>url,view:()=>view};
}
test('expired session links recover to upload and remove only obsolete parameters',async()=>{
  const h=setup(Object.assign(new Error('Phiên đã mất'),{code:'SESSION_NOT_FOUND'}));
  await h.ctx.initializeWorkspace();
  assert.equal(h.url(),'/?other=kept');assert.equal(h.view(),'upload');
  assert.equal(h.nodes['upload-error'].textContent,'Phiên đã mất');
  assert.equal(h.nodes['upload-error'].hidden,false);
  assert.equal(h.ctx.state.job,null);
});
test('network failure does not discard session URL or classify it as lost',async()=>{
  const h=setup(new Error('Offline'));
  await assert.rejects(h.ctx.initializeWorkspace(),/Offline/);
  assert.equal(h.url(),undefined);assert.equal(h.ctx.state.job.id,'old');
});
