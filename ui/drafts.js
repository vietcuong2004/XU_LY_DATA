/* Durable browser drafts. They are not a replacement for server-side sessions. */
(() => {
  let opened;
  let writes=Promise.resolve();
  function database(){
    if(!opened)opened=new Promise((resolve,reject)=>{
      const request=indexedDB.open('ShipmentStudio',1);
      request.onupgradeneeded=()=>request.result.createObjectStore('drafts');
      request.onsuccess=()=>resolve(request.result);
      request.onerror=()=>reject(request.error);
    });
    return opened;
  }
  async function transaction(mode,action){
    const db=await database();
    return new Promise((resolve,reject)=>{
      const tx=db.transaction('drafts',mode),request=action(tx.objectStore('drafts'));
      tx.oncomplete=()=>resolve(request?.result);
      tx.onerror=()=>reject(tx.error);
      tx.onabort=()=>reject(tx.error||new Error('Không lưu được bản nháp.'));
    });
  }
  window.shipmentDrafts={
    read:key=>writes.then(()=>transaction('readonly',store=>store.get(key))),
    save(key,value){
      const snapshot=structuredClone(value);
      const task=writes.catch(()=>{}).then(()=>transaction('readwrite',store=>store.put(snapshot,key)));
      writes=task.catch(()=>{});return task;
    },
    remove(key){
      const task=writes.catch(()=>{}).then(()=>transaction('readwrite',store=>store.delete(key)));
      writes=task.catch(()=>{});return task;
    }
  };
})();
