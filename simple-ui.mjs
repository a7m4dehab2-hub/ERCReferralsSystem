import { unifiedRows, workFilters, actionSummary } from './workspace.mjs';
import { narratives, isOpen, dateLabel, validId, missingDetails } from './domain.mjs';
export function simpleWorkspace(page,ctx){
 const {node:n,action:a,t,data,state}=ctx;
 if(state.editor)return simpleEditor(page,ctx);
 const controls=n('section',{class:'simple-controls'}),search=n('input',{type:'search','aria-label':t('Search name, ID or phone'),placeholder:t('Search name, ID or phone'),value:state.search});
 let timer;search.addEventListener('input',()=>{state.search=search.value;clearTimeout(timer);timer=setTimeout(()=>drawRecords(),100);});
 const sort=n('select',{'aria-label':t('Sort patients')},['Priority','Next action due','Appointment date','Patient name','Age: youngest','Age: oldest','Arrival date: newest','Arrival date: oldest'].map(value=>n('option',{value,text:t(value)})));sort.value=state.simpleSort||'Priority';sort.addEventListener('change',()=>{state.simpleSort=sort.value;drawRecords();});
 controls.append(n('div',{class:'simple-search-row'},search,n('label',{class:'small-label'},t('Sort by'),sort),a('More filters',()=>{state.moreFilters=!state.moreFilters;ctx.render();}),a(state.tab==='patients'?'Add patient':'Add request',()=>ctx.beginEditor(state.tab==='patients'?'new-patient':'new-request'),{primary:true})));
 if(!['patients','completed','closed'].includes(state.tab)){
  const group=n('div',{class:'work-filter-buttons','aria-label':t('Work filters')});
  for(const [key,label]of workFilters){const active=(state.simpleFilters||[]).includes(key)||(!state.simpleFilters?.length&&(state.tab===key||(key==='urgent'&&state.priority==='Urgent')));group.append(a(label,()=>{const selected=new Set(state.simpleFilters?.length?state.simpleFilters:workFilters.some(([k])=>k===state.tab)?[state.tab]:[]);if(selected.has(key))selected.delete(key);else selected.add(key);ctx.setFilters([...selected]);},{className:active?'filter-active':'',help:'Filter patients by this outstanding work. Multiple selected filters match any category.'}));}
  group.append(a('All active',()=>ctx.setFilters([]),{className:!state.simpleFilters?.length&&['all','overview'].includes(state.tab)?'filter-active':''}));controls.append(group);
 }
 if(state.tab!=='patients')controls.append(n('div',{class:'archive-links'},['completed','closed'].includes(state.tab)?a('All active',()=>ctx.setFilters([])):null,a('Completed requests',()=>ctx.navigate('completed')),a('Closed requests',()=>ctx.navigate('closed'))));
 if((state.tab==='investigations'||state.simpleFilters?.includes('investigations'))&&(data.imports||[]).some(row=>row.status==='pending'))controls.append(n('div',{class:'note-box'},t('Source entries awaiting review')+': '+data.imports.filter(row=>row.status==='pending').length,a('Review source entries',()=>ctx.navigate('imports'))));
 if(state.moreFilters)ctx.renderFilters(controls);
 page.append(controls);
 const records=n('section',{class:'patient-records','aria-label':t('Unified patient records')});page.append(records);
 function drawRecords(){
  records.replaceChildren();const rows=unifiedRows(ctx.getRows(),data,{filters:state.simpleFilters||[],sort:state.simpleSort||'Priority',now:ctx.now});
  records.append(n('p',{class:'record-count',text:rows.length+' '+t('patients')}));
  if(!rows.length)records.append(n('div',{class:'empty-state'},n('h2',{text:t(data.patients.length?'No matching records':'No patients yet')}),n('p',{text:t(data.patients.length?'Try changing your search or filters.':'Add a patient or import your existing records to start.')}),!data.patients.length?a('Add patient',()=>{ctx.navigate('patients');ctx.beginEditor('new-patient');},{primary:true}):null));
  for(const group of rows){const {p,requests,first,active}=group,card=n('article',{class:'patient-record','data-patient':p.key});
   const badges=n('div',{class:'badges'});if(requests.some(r=>r.urgent))badges.append(n('span',{class:'badge urgent',text:t('Urgent')}));if(!validId(p.id))badges.append(n('span',{class:'badge warning',text:t('ID needs verification')}));if(missingDetails(p,data.settings.necessary).length)badges.append(n('span',{class:'badge warning',text:t('Missing details')}));
   const title=n('div',{class:'record-heading'},n('div',{},n('button',{class:'patient-title',text:p.name,on:{click:()=>patientPanel(ctx,p)}}),n('p',{class:'caption',text:p.id+' · '+(p.centre||t('Not recorded'))})),badges);
   const phones=n('div',{class:'record-phones'});for(const number of [p.phone,p.otherPhone].filter(Boolean))phones.append(n('span',{class:'phone-line'},n('bdi',{text:number}),a('WhatsApp',()=>ctx.whatsapp(p,number),{className:'whatsapp-button',help:'Choose a request and review an editable message before opening WhatsApp.'})));if(!phones.childNodes.length)phones.append(n('span',{class:'field-error',text:t('Contact number missing')}));
   const summary=n('div',{class:'case-narrative'},n('strong',{text:t('What needs attention')}),n('p',{text:actionSummary(p,requests,ctx.language)}));
   const latest=(data.contactOutcomes||[]).filter(c=>c.patientKey===p.key&&requests.some(r=>r.key===c.requestKey)).sort((a,b)=>b.contactAt.localeCompare(a.contactAt))[0];
   if(latest)summary.append(n('p',{class:'latest-outcome',text:t('Latest contact')+': '+t(latest.result)+' · '+latest.staff+(latest.notes?'\n'+latest.notes:'')}));
   const footer=n('div',{class:'record-footer'},n('span',{class:'caption',text:requests.length+' '+t('matching requests')+' · '+active.length+' '+t('active requests')+(first?.owner?' · '+first.owner:'')}),a('Open patient',()=>patientPanel(ctx,p)),a('Record contact outcome',()=>ctx.contact(p,requests.length===1?first:null),{primary:true,disabled:!active.length,help:'Save the result and original notes against the correct request.'}));
   card.append(title,phones,summary,footer);records.append(card);
  }
 }
 drawRecords();
}
export function patientPanel(ctx,p){const {node:n,action:a,t,data}=ctx,requests=data.requests.filter(r=>r.patientKey===p.key),content=n('div',{class:'patient-profile'});
 const info=n('dl',{class:'profile-grid'});for(const field of ctx.patientFields.filter(f=>f.key!=='age'))info.append(n('div',{},n('dt',{text:t(field.label)}),n('dd',{text:p[field.key]||t('Not recorded')})));
 content.append(info,n('h3',{text:t('Requests')}));
 for(const r of requests)content.append(n('section',{class:'profile-request'},n('h3',{text:r.reference+' · '+(r.requested||r.approved||r.received||'—')}),n('span',{class:'badge',text:t(r.progress)}),n('p',{text:[r.hospital,dateLabel(r.appointment),r.next].filter(Boolean).join(' · ')}),n('div',{class:'record-footer'},a('Open request',()=>requestPanel(ctx,p,r)),a('Edit request',()=>{ctx.closeDialog();ctx.edit(p,r);}),a('Record contact outcome',()=>{ctx.closeDialog();ctx.contact(p,r);})))) // end request
;
 const summary=narratives(p,requests);content.append(n('h3',{text:t('English summary')}),n('p',{class:'full-summary',text:summary.en,dir:'ltr'}),n('h3',{text:t('Arabic summary')}),n('p',{class:'full-summary',text:summary.ar,dir:'rtl'}));
 const contacts=(data.contactOutcomes||[]).filter(c=>c.patientKey===p.key);content.append(n('h3',{text:t('Contact outcomes')}));for(const c of contacts)content.append(n('div',{class:'profile-contact'},n('strong',{text:t(c.result)+' · '+c.staff}),n('p',{class:'caption',text:new Date(c.contactAt).toLocaleString('en-GB',{timeZone:'Africa/Cairo'})}),n('p',{class:'full-summary',text:c.notes})));
 ctx.panel(p.name,content,[a('Edit patient',()=>{ctx.closeDialog();ctx.edit(p,null);}),a('Add request',()=>{ctx.closeDialog();ctx.addForPatient(p);})]);
}
export function requestPanel(ctx,p,r){const {node:n,action:a,t}=ctx,content=n('div',{class:'patient-profile'},n('p',{class:'case-narrative',text:actionSummary(p,[r],ctx.language)}));const grid=n('dl',{class:'profile-grid'});for(const field of ctx.requestFields)grid.append(n('div',{},n('dt',{text:t(field.label)}),n('dd',{text:field.type==='date'?dateLabel(r[field.key])||'—':typeof r[field.key]==='boolean'?t(r[field.key]?'Yes':'No'):r[field.key]||'—'})));content.append(grid);ctx.panel(p.name+' · '+r.reference,content,[a('Edit request',()=>{ctx.closeDialog();ctx.edit(p,r);},{primary:true}),a('Record contact outcome',()=>{ctx.closeDialog();ctx.contact(p,r);}),a('View history',()=>{ctx.closeDialog();ctx.history(p,r);})]);}
function simpleEditor(page,ctx){const {node:n,action:a,t,state}=ctx,editor=state.editor,form=n('section',{class:'simple-editor'},n('h2',{text:t(editor.r?'Request details':'Patient details')}));
 if(editor.lookupNote)form.append(n('p',{class:'note-box',text:t(editor.lookupNote)}));
 const fields=n('div',{class:'simple-form-grid'});
 for(const field of [...ctx.patientFields,...(editor.r?ctx.requestFields:[])]){
  const target=field.scope==='p'?editor.p:editor.r;if(field.type==='computed')continue;
  let input;const options=field.key==='owner'?['',...ctx.staffNames()]:field.type==='hospital'?['',...ctx.data.settings.hospitals.map(h=>h.name),target[field.key]||'']:field.type==='agency'?['',...ctx.data.settings.agencies.map(h=>h.name),target[field.key]||'']:field.options;
  if(options){input=n('select',{},[...new Set(options)].map(value=>n('option',{value,text:t(value||'Not recorded')})));input.value=target[field.key]||'';}
  else if(field.type==='long')input=n('textarea',{rows:3,value:target[field.key]||''});
  else input=n('input',{type:field.type==='checkbox'?'checkbox':field.type==='date'?'date':field.type==='number'?'number':'text',value:target[field.key]??'',checked:field.type==='checkbox'?!!target[field.key]:undefined});
  input.dataset.input=field.id;input.setAttribute('aria-label',t(field.label));input.addEventListener('input',()=>{target[field.key]=field.type==='checkbox'?input.checked:input.value;editor.dirty=true;});if(field.key==='id'&&editor.mode==='new-request')input.addEventListener('blur',ctx.lookup);
  if(field.key==='source')input.addEventListener('change',()=>{if(editor.mode==='new-request'&&target.source==='Direct ministry approval')editor.r.progress='Approved';});
  const label=n('label',{class:field.type==='long'?'form-wide':''},t(field.label),input);if(editor.errors[field.id]){input.setAttribute('aria-invalid','true');label.append(n('span',{class:'field-error',text:t(editor.errors[field.id])}));}fields.append(label);
 }
 form.append(fields,n('label',{class:'small-label'},t('Reason for correction / reopening'),n('input',{'aria-label':t('Reason for correction / reopening'),value:editor.reason||'',on:{input:e=>{editor.reason=e.target.value;editor.dirty=true;}}})),n('div',{class:'record-footer'},a('Save',ctx.saveEditor,{primary:true,id:'save-record'}),a('Cancel',ctx.cancelEditor,{id:'cancel-edit'})));page.append(form);
}
