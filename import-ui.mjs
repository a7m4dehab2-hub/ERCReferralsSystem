import { autoMapping, findHeader, sheetGrid, importFields, parseRows, digest, classify, isSafe, applySourceRow, sourceProblems, applySafeBatch } from './importer.mjs';
import { normalizeId, clean } from './domain.mjs';
import { changeEvents } from './store.mjs';
import { fetchContactSheet } from './sheet-source.mjs';
let workbook, fileName='', fileHash='', sheetName='', kind='contacts', headerIndex=0, mapping={}, preview=[], message='', filter='Pending', busy=false;
let importEpoch=0;
export function resetImports(){importEpoch++;workbook=null;fileName='';fileHash='';sheetName='';kind='contacts';headerIndex=0;mapping={};preview=[];message='';filter='Pending';busy=false;}
export async function prepareSheetUpdate({render,isActive}){
 const epoch=++importEpoch;busy=true;preview=[];workbook=null;fileName='';mapping={};message='Reading the contact sheet. Your existing patient data will remain unchanged until reviewed.';render();
 try{
  const source=await fetchContactSheet();if(epoch!==importEpoch||!isActive())return;
  workbook=window.XLSX.read(source.csv,{type:'string',raw:true,cellDates:false,cellNF:true});sheetName=workbook.SheetNames[0];kind='contacts';fileName='Google contact sheet.csv';fileHash=await digest(source.csv);
  if(epoch!==importEpoch||!isActive())return;
  const grid=sheetGrid(window.XLSX,workbook.Sheets[sheetName]);headerIndex=findHeader(grid);if(headerIndex<0)throw Error('Patient name and ID headers were not found in the contact sheet. Review the sheet format before updating.');mapping=autoMapping(grid[headerIndex]);
  const rows=await parseRows({grid,headers:grid[headerIndex],mapping,headerIndex,fileName,fileHash,sheet:'بيانات التواصل الصحي',kind});if(epoch!==importEpoch||!isActive())return;
  preview=rows.map(row=>({...row,source:{...row.source,url:source.sourceUrl,fetchedAt:source.fetchedAt}}));message=`Read ${preview.length} submissions from the contact sheet. Review differences below. Identical entries are skipped; conflicting values will not be overwritten.`;
 }catch(error){if(epoch===importEpoch)message='Update data: '+error.message;}finally{if(epoch===importEpoch){busy=false;render();}}
}
const h=(tag,props={},...children)=>{const e=document.createElement(tag);for(const [key,val] of Object.entries(props)){if(key==='on')for(const [event,fn]of Object.entries(val))e.addEventListener(event,fn);else if(key==='text')e.textContent=val;else if(key==='disabled'||key==='checked')e[key]=!!val;else e.setAttribute(key,val);}for(const child of children.flat())if(child!=null)e.append(child instanceof Node?child:document.createTextNode(String(child)));return e;};
const button=(label,fn,disabled=false)=>h('button',{type:'button',text:label,disabled,on:{click:fn}});
const dropdown=(label,options,value,fn)=>{const input=h('select',{'aria-label':label},options.map(o=>h('option',{value:typeof o==='string'?o:o.value,text:typeof o==='string'?o:o.label})));input.value=value;input.addEventListener('change',()=>fn(input.value));return h('label',{class:'small-label'},label,input);};
export function renderImports(page,ctx){
 const {data,admin,save,refresh,showDialog,closeDialog}=ctx;
 const section=h('section',{class:'import-section'});page.append(section);
 const render=()=>{section.replaceChildren();
  section.append(h('h2',{text:'Import and source review'}),h('p',{class:'caption',text:'1. Choose the contact workbook first, then the master workbook’s Follow_UP tab. 2. Preview and save source rows. 3. Apply unambiguous entries or review each held entry below. Saving source rows alone does not create patient profiles or requests.'}),h('p',{class:'caption',text:'For the prepared private ERC JSON package, use Settings → Backups → Restore backup into empty workspace. This file picker accepts Excel and CSV files.'}));
  const controls=h('div',{class:'toolbar'});
  const file=h('input',{type:'file',accept:'.xlsx,.xls,.csv','aria-label':'Choose Excel file',disabled:!admin||busy,on:{change:async()=>{
   const selected=file.files[0];if(!selected)return;if(selected.size>30*1024*1024){message='Choose a workbook smaller than 30 MB.';return render();}
   busy=true;render();try{
    if(!window.XLSX)throw Error('The Excel reader is missing from this installation.');
    const buffer=await selected.arrayBuffer();fileHash=[...new Uint8Array(await crypto.subtle.digest('SHA-256',buffer))].map(n=>n.toString(16).padStart(2,'0')).join('');fileName=selected.name;
    workbook=window.XLSX.read(buffer,{type:'array',cellDates:false,cellNF:true});
    sheetName=workbook.SheetNames.find(n=>n.toLowerCase()==='follow_up')||workbook.SheetNames.find(n=>n.includes('بيانات التواصل'))||workbook.SheetNames[0];
    kind=sheetName.toLowerCase()==='follow_up'?'referrals':'contacts';setupSheet();message='Check the sheet, header row and field mapping before previewing.';
   }catch(error){message=error.message;}finally{busy=false;render();}
  }}});
  controls.append(h('label',{},'Choose Excel file ',file));
  if(workbook){controls.append(h('span',{text:fileName}),dropdown('Import type',[{value:'contacts',label:'Patient contact submissions'},{value:'referrals',label:'Referral requests'}],kind,v=>{kind=v;preview=[];render();}),dropdown('Worksheet',workbook.SheetNames,sheetName,v=>{sheetName=v;setupSheet();render();}));
   controls.append(h('label',{class:'small-label'},'Header row',h('input',{type:'number',min:1,max:50,value:headerIndex+1,'aria-label':'Header row',on:{change:event=>{headerIndex=Math.max(0,Number(event.target.value)-1);mapping=autoMapping(grid()[headerIndex]||[]);preview=[];render();}}})));
  }
  section.append(controls);
  if(busy)section.append(h('p',{role:'status','data-import-progress':'',text:'Working…'}));
  if(message)section.append(h('p',{class:'note-box',role:'status',text:message}));
  if(workbook){
   const headers=grid()[headerIndex]||[], fields=kind==='contacts'?importFields.slice(0,12):importFields;
   const details=h('details',{class:'mapping-panel'},h('summary',{text:'Review field mapping'}),h('p',{class:'caption',text:'Unmapped original columns remain available in Source details. Do not map a received service as an approval unless you have verified that meaning.'}));
   const fieldsDiv=h('div',{class:'mapping-fields'});
   for(const field of fields)fieldsDiv.append(dropdown(field.label,[{value:'',label:'Not in this sheet'},...headers.map((header,i)=>({value:String(i),label:`${i+1}: ${header||'(blank)'}`}))],mapping[field.key]===undefined?'':String(mapping[field.key]),v=>{if(v==='')delete mapping[field.key];else mapping[field.key]=Number(v);preview=[];}));
   details.append(fieldsDiv);section.append(details,button('Preview source rows',async()=>{busy=true;render();try{const links={};const sheet=workbook.Sheets[sheetName];for(const [address,cell]of Object.entries(sheet)){if(address.startsWith('!')||!cell.l?.Target)continue;const pos=window.XLSX.utils.decode_cell(address);links[`${pos.r}:${pos.c}`]=cell.l.Target;}preview=await parseRows({grid:grid(),headers,mapping,headerIndex,fileName,fileHash,sheet:sheetName,kind,links});message=`${preview.length} nonempty source rows. Check the classifications below. No data saved yet.`;}catch(error){message=error.message;}finally{busy=false;render();}},busy||!admin));
  }
  if(preview.length){
   const fingerprints=new Set((data.imports||[]).map(r=>r.fingerprint)),unique=[];
   for(const row of preview)if(!fingerprints.has(row.fingerprint)&&classify(row,data,preview)!=='Identical profile — skipped'){fingerprints.add(row.fingerprint);unique.push(row);}
   section.append(h('p',{class:'caption',text:`${unique.length} new source entries; ${preview.length-unique.length} identical entries will be skipped.`}),button('Save source rows for review',async()=>{
    busy=true;render();try{const next=await refresh();const seen=new Set((next.imports||[]).map(r=>r.fingerprint));let count=0;for(const row of unique)if(!seen.has(row.fingerprint)){next.imports.push(row);seen.add(row.fingerprint);count++;}next.revision++;await save(next);preview=[];message=`Saved ${count} source rows. Patient profiles and requests are created only when you apply them below.`;}catch(error){message=error.message;}finally{busy=false;ctx.rerender();}
   },busy||!admin||!unique.length));
   section.append(sourceTable(preview.slice(0,100),row=>classify(row,data,preview),row=>sourceDialog(row)));
   if(preview.length>100)section.append(h('p',{text:'Showing the first 100 preview rows. All nonempty rows are included when you save.'}));
  }
  const pending=(data.imports||[]).filter(r=>r.status==='pending');
  section.append(h('h2',{text:'Saved source entries'}),h('p',{class:'caption',text:`${pending.length} pending review · ${(data.imports||[]).filter(r=>r.status==='applied').length} applied · ${(data.imports||[]).filter(r=>r.status==='skipped').length} skipped`}),dropdown('Source entry status',['Pending','Applied','Skipped','All'],filter,v=>{filter=v;render();}));
  const safe=pending.filter(row=>isSafe(classify({...row,fingerprint:''},{...data,imports:[]},pending)));
  section.append(button(`Apply ${safe.length} unambiguous entries`,async()=>{
   if(!confirm('Apply unambiguous entries? Imported requests start Under verification. Original statuses and source rows are preserved for review.'))return;
   const allowedKeys=new Set(safe.map(row=>row.key)),epoch=importEpoch;let count=0;
   busy=true;message='';render();try{
    let next=await refresh();
    while(epoch===importEpoch&&section.isConnected){
     const batch=applySafeBatch(next,{limit:50,allowedKeys});if(!batch)break;
     next.revision++;next=await save(next,{reason:'Reviewed import batch: unambiguous entries; requests held for verification'});count+=batch;
     message=`Applied ${count} of ${allowedKeys.size} entries. Each completed batch is saved.`;
     const progress=section.querySelector('[data-import-progress]');if(progress)progress.textContent=message;
    }
    message=`Applied ${count} entries. Conflicting or incomplete entries remain pending.`;
   }catch(error){message=`Applied ${count} entries before stopping. Completed batches stay saved. Refresh records before retrying. ${error.message}`;}finally{busy=false;ctx.rerender();}
  },busy||!admin||!safe.length));
  const rows=(data.imports||[]).filter(r=>filter==='All'||r.status===filter.toLowerCase());
  section.append(sourceTable(rows, row=>row.status==='pending'?classify({...row,fingerprint:''},{...data,imports:[]},pending):row.status,row=>reviewDialog(row)));
  function sourceDialog(row){showDialog('Original source row',rawDetails(row));}
  function reviewDialog(row){
   if(row.status!=='pending'||!admin)return sourceDialog(row);
   let values={...row.values},choices={},requestChoices={};
   const content=h('div',{},h('p',{class:'note-box',text:`${row.source.fileName} / ${row.source.sheet} / row ${row.source.row}. Choose which incoming personal values to keep. The original submission is preserved.`}));
   const idInput=h('input',{'aria-label':'Resolved patient ID',value:values.id||''});
   content.append(h('label',{class:'small-label'},'Resolved patient ID ',idInput));
   const body=h('div');content.append(body);
   let requestTarget='new',progress='Under verification',informed='Not informed',outcome='Not confirmed';
   function updateReview(){
    values.id=normalizeId(idInput.value);choices={};body.replaceChildren();
    const p=data.patients.find(p=>normalizeId(p.id)===values.id);
    const siblings=(data.imports||[]).filter(r=>normalizeId(r.values.id)===values.id);
    body.append(h('p',{text:`${p?'Existing patient: '+p.name:'New patient profile'} · ${siblings.length} saved submission(s) for this ID`}),sourceTable(siblings,r=>r.status,r=>sourceDialog(r)));
    const table=h('table',{class:'history-table'},h('thead',{},h('tr',{},['Field','Current profile','Incoming value','Keep'].map(text=>h('th',{text}))))),tbody=h('tbody');
    for(const field of importFields.slice(0,12).filter(f=>f.key!=='id')){
     const sourceInput=h('input',{'aria-label':'Incoming '+field.label,value:values[field.key]||'',type:['birth','entryDate'].includes(field.key)?'date':'text',on:{input:event=>{values[field.key]=event.target.value;choices[field.key]=true;checkbox.checked=true;}}});
     choices[field.key]=!p?.[field.key]&&!!values[field.key];
     const checkbox=h('input',{type:'checkbox','aria-label':'Use incoming '+field.label,checked:choices[field.key],on:{change:event=>{choices[field.key]=event.target.checked;}}});
     tbody.append(h('tr',{},h('td',{text:field.label}),h('td',{text:p?.[field.key]||'—'}),h('td',{},sourceInput),h('td',{},checkbox)));
    }
    table.append(tbody);body.append(h('div',{class:'dialog-scroll'},table));
    if(row.kind==='referrals'){
     const requests=data.requests.filter(r=>r.patientKey===p?.key);requestTarget=requests.length?'':'new';
     const requestPanel=h('div');
     function showRequestFields(){
      const existing=requests.find(r=>r.key===requestTarget);requestChoices={};requestPanel.replaceChildren();
      progress=existing?.progress||'Under verification';informed=existing?.informed||'Not informed';outcome=existing?.outcome||'Not confirmed';
      const table=h('table',{class:'history-table'},h('thead',{},h('tr',{},['Request field','Current request','Incoming value','Keep'].map(text=>h('th',{text}))))),tbody=h('tbody');
      for(const field of importFields.slice(12)){
       requestChoices[field.key]=!existing?.[field.key]&&!!values[field.key];
       const choice=h('input',{type:'checkbox','aria-label':'Use incoming '+field.label,checked:requestChoices[field.key],on:{change:event=>requestChoices[field.key]=event.target.checked}});
       const input=h('input',{'aria-label':'Incoming '+field.label,type:['appointment','reportDate','sourceDate'].includes(field.key)?'date':'text',value:values[field.key]||'',on:{input:event=>{values[field.key]=event.target.value;requestChoices[field.key]=true;choice.checked=true;}}});
       tbody.append(h('tr',{},h('td',{text:field.label}),h('td',{text:existing?.[field.key]||'—'}),h('td',{},input),h('td',{},choice)));
      }
      table.append(tbody);requestPanel.append(h('div',{class:'dialog-scroll'},table),dropdown('Verified request progress',['Under verification','Ready to send','Awaiting approval','Approved','Completed','Closed'],progress,v=>progress=v),dropdown('Patient informed?',['Not informed','Awaiting confirmation','Informed','No answer','Unavailable number','Wrong number','Callback requested'],informed,v=>informed=v),dropdown('Verified service outcome',['Not confirmed','Full service','Partial service','No service'],outcome,v=>outcome=v),h('label',{class:'small-label'},'Closure reason',h('input',{'aria-label':'Closure reason',value:existing?.closureReason||'',on:{input:event=>values.closureReason=event.target.value}})));
     }
     body.append(dropdown('Apply to request',[{value:'',label:'Choose request'},...requests.map(r=>({value:r.key,label:`${r.reference} · ${r.requested||r.approved||r.received||'No procedure recorded'} · ${r.hospital||'No hospital'}`})),{value:'new',label:'New independent request'}],requestTarget,v=>{requestTarget=v;showRequestFields();}),requestPanel);showRequestFields();
    }
   }
   updateReview();idInput.addEventListener('change',updateReview);
   const reason=h('input',{'aria-label':'Import review reason',placeholder:'Describe corrections or verification'}),error=h('p',{class:'field-error',role:'alert'});
   content.append(h('label',{class:'small-label'},'Review note ',reason),error);
   showDialog('Review source entry',content,[button('Apply reviewed entry',async()=>{
    try{const next=await refresh();if(next.revision!==data.revision)throw Error('Records changed. Close this review, refresh records and review the latest values.');const before=structuredClone(next);values.id=normalizeId(idInput.value);const useFields=Object.keys(choices).filter(k=>choices[k]),useRequestFields=Object.keys(requestChoices).filter(k=>requestChoices[k]);applySourceRow(next,row,{values,useFields,useRequestFields,requestKey:requestTarget,progress,informed,outcome,reason:reason.value});addLocalAudit(before,next);next.revision++;await save(next,{reason:reason.value||'Source entry reviewed manually'});closeDialog();message='Reviewed entry applied.';ctx.rerender();}catch(failure){error.textContent=failure.message;}
   },busy),button('Keep for investigation',()=>closeDialog()),button('Skip this entry',async()=>{
    if(!clean(reason.value)){error.textContent='Write a reason for skipping this entry.';return;}
    try{const next=await refresh();next.imports=next.imports.map(r=>r.key===row.key?{...r,status:'skipped',resolution:reason.value,resolvedAt:new Date().toISOString()}:r);next.revision++;await save(next,{reason:reason.value});closeDialog();ctx.rerender();}catch(failure){error.textContent=failure.message;}
   })]);
  }
  function addLocalAudit(before,next){for(const [collection,type]of [['patients','patient'],['requests','request']])for(const row of next[collection]){const old=before[collection].find(r=>r.key===row.key);if(JSON.stringify(old)!==JSON.stringify(row))next.audit.push(...changeEvents(old,row,type,ctx.actor));}}
 };
 function grid(){return sheetGrid(window.XLSX,workbook.Sheets[sheetName],{date1904:!!workbook.Workbook?.WBProps?.date1904});}
 function setupSheet(){headerIndex=Math.max(0,findHeader(grid()));mapping=autoMapping(grid()[headerIndex]||[]);preview=[];}
 function rawDetails(row){return h('div',{},h('p',{text:`${row.source.fileName} / ${row.source.sheet} / row ${row.source.row}`}),h('table',{class:'history-table'},h('tbody',{},row.source.cells.map(cell=>h('tr',{},h('th',{text:`${cell.column}: ${cell.header}`}),h('td',{text:cell.value+(cell.link?'\n'+cell.link:'')}))))));}
 function sourceTable(rows,label,review){return h('div',{class:'dialog-scroll'},h('table',{class:'history-table source-table'},h('thead',{},h('tr',{},['Source row','Patient ID','Patient name','Classification','Warnings / fields to review','Actions'].map(text=>h('th',{text})))),h('tbody',{},rows.map(row=>h('tr',{},h('td',{text:`${row.source.sheet}: ${row.source.row}`}),h('td',{text:row.values.id||'Missing ID'}),h('td',{text:row.values.name||'Missing name'}),h('td',{text:label(row)}),h('td',{text:[...new Set([...(row.issues||[]),...sourceProblems(row,data.patients.find(p=>normalizeId(p.id)===normalizeId(row.values.id)))])].join('; ')||'—'}),h('td',{},button(row.status==='pending'?'Review':'Source details',()=>review(row))))))));}
 render();
}
