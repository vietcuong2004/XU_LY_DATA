'use strict';
(function(global){
  function packagePath(base,target){
    const parts=(target.startsWith('/')?target.slice(1):`${base}/${target}`).split('/'),result=[];
    for(const part of parts){
      if(!part||part==='.')continue;
      if(part==='..')result.pop();else result.push(part);
    }
    return result.join('/');
  }

  function parseXml(bytes,label){
    const document=new DOMParser().parseFromString(new TextDecoder().decode(bytes),'application/xml');
    if(document.getElementsByTagName('parsererror').length)throw new Error(`Không đọc được cấu trúc ${label} trong file Excel.`);
    return document;
  }

  function xmlBytes(document){
    return new TextEncoder().encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+new XMLSerializer().serializeToString(document.documentElement));
  }

  async function slimSourceWorkbook(file){
    if(!global.fflate)return file;
    const original=new Uint8Array(await file.arrayBuffer());
    const files=global.fflate.unzipSync(original);
    const workbookPath='xl/workbook.xml',relsPath='xl/_rels/workbook.xml.rels',typesPath='[Content_Types].xml';
    if(!files[workbookPath]||!files[relsPath]||!files[typesPath])return file;
    const workbook=parseXml(files[workbookPath],'workbook.xml');
    const sheets=[...workbook.getElementsByTagNameNS('*','sheet')];
    const sum=sheets.find(sheet=>sheet.getAttribute('name')==='SUM');
    if(!sum)return file;
    const relationNamespace='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const sumRelation=sum.getAttributeNS(relationNamespace,'id')||sum.getAttribute('r:id');
    const relationships=parseXml(files[relsPath],'workbook.xml.rels');
    const relationElements=[...relationships.getElementsByTagNameNS('*','Relationship')];
    const sumEntry=relationElements.find(relation=>relation.getAttribute('Id')===sumRelation);
    if(!sumEntry)return file;
    const sumPath=packagePath('xl',sumEntry.getAttribute('Target'));
    if(!files[sumPath])return file;
    const removed=new Set(['/xl/calcChain.xml']);
    for(const sheet of sheets)if(sheet!==sum)sheet.parentNode.removeChild(sheet);
    // Local print areas/names point to sheets removed from this temporary upload.
    for(const names of [...workbook.getElementsByTagNameNS('*','definedNames')])names.parentNode.removeChild(names);
    for(const relation of relationElements){
      const type=relation.getAttribute('Type')||'';
      if((type.endsWith('/worksheet')&&relation.getAttribute('Id')!==sumRelation)||type.endsWith('/calcChain')){
        const target=packagePath('xl',relation.getAttribute('Target'));
        removed.add('/'+target);
        delete files[target];
        const slash=target.lastIndexOf('/'),parent=target.slice(0,slash),name=target.slice(slash+1);
        delete files[`${parent}/_rels/${name}.rels`];
        relation.parentNode.removeChild(relation);
      }
    }
    delete files['xl/calcChain.xml'];
    const contentTypes=parseXml(files[typesPath],'[Content_Types].xml');
    for(const override of [...contentTypes.getElementsByTagNameNS('*','Override')]){
      if(removed.has(override.getAttribute('PartName')))override.parentNode.removeChild(override);
    }
    files[workbookPath]=xmlBytes(workbook);
    files[relsPath]=xmlBytes(relationships);
    files[typesPath]=xmlBytes(contentTypes);
    const packed=global.fflate.zipSync(files,{level:6});
    if(packed.length>=original.length)return file;
    return new File([packed],file.name,{type:file.type||'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',lastModified:file.lastModified});
  }

  global.sourceOptimizer={slimSourceWorkbook};
})(window);
