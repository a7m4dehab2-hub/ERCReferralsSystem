import { isOpen, priorityRank, queueMatches, narratives, clean, normalizeId, dateLabel, ageAt } from './domain.mjs';
export const workFilters=[['ready','Ready to send'],['awaiting','Awaiting approval'],['appointment','Needs appointment'],['notification','Needs notification'],['today','Follow-ups today'],['overdue','Overdue follow-ups'],['investigations','Investigations'],['urgent','Urgent requests']];
export function workMatch(r,p,key,settings,now){return key==='today'?isOpen(r)&&r.due===now:key==='overdue'?isOpen(r)&&!!r.due&&r.due<now:key==='urgent'?isOpen(r)&&r.urgent:queueMatches(r,p,key,settings,now);}
export function unifiedRows(rows,data,{filters=[],sort='Priority',now}){
 const groups=new Map();
 for(const {p,r}of rows){if(r&&filters.length&&!filters.some(f=>workMatch(r,p,f,data.settings,now)))continue;const group=groups.get(p.key)||{p,requests:[],active:data.requests.filter(r=>r.patientKey===p.key&&isOpen(r))};if(r)group.requests.push(r);groups.set(p.key,group);}
 const first=group=>group.requests.slice().sort((a,b)=>priorityRank(a,now)-priorityRank(b,now)||(a.due||'9999').localeCompare(b.due||'9999'))[0];
 const age=group=>ageAt(group.p.birth,now)??(group.p.reportedAge!==''&&group.p.reportedAge!=null?Number(group.p.reportedAge):null);
 const result=[...groups.values()].map(group=>({...group,first:first(group)}));
 const due=group=>group.requests.map(r=>r.due).filter(Boolean).sort()[0]||'9999',appt=group=>group.requests.map(r=>r.appointment).filter(Boolean).sort()[0]||'9999';
 return result.sort((a,b)=>sort==='Patient name'?a.p.name.localeCompare(b.p.name):sort==='Next action due'?due(a).localeCompare(due(b)):sort==='Appointment date'?appt(a).localeCompare(appt(b)):sort==='Arrival date: newest'?(b.p.entryDate||'').localeCompare(a.p.entryDate||''):sort==='Arrival date: oldest'?(a.p.entryDate||'9999').localeCompare(b.p.entryDate||'9999'):sort==='Age: youngest'?(age(a)??999)-(age(b)??999):sort==='Age: oldest'?(age(b)??-1)-(age(a)??-1):(a.first?priorityRank(a.first,now):5)-(b.first?priorityRank(b.first,now):5)||due(a).localeCompare(due(b))||a.p.name.localeCompare(b.p.name));
}
export function actionSummary(p,requests,language='en'){
 const ar=language==='ar';const actions=[];
 for(const r of requests){if(r.progress==='Completed'){actions.push(r.reference+': '+(ar?'تم تأكيد الخدمة كاملة':'Full service confirmed'));continue;}if(r.progress==='Closed'){actions.push(r.reference+': '+(ar?'مغلق: ':'Closed: ')+(r.closureReason||''));continue;}const ref=r.reference+': ';let needs=[];
  if(r.approved&&(!r.hospital||!r.appointment))needs.push(ar?'استكمال المستشفى والموعد الناقصين':'Obtain missing hospital/appointment details');
  else if(r.approved&&r.hospital&&r.appointment&&r.informed!=='Informed')needs.push(ar?'إبلاغ المريض وتأكيد استلامه':'Notify the patient and confirm acknowledgement');
  if(r.progress==='Awaiting approval')needs.push(ar?'متابعة رد جهة الإحالة':'Follow up the agency response');
  if(r.progress==='Ready to send')needs.push(ar?'جاهز للإرسال اليدوي':'Ready for manual submission');
  if(r.progress==='Under verification')needs.push(ar?'مراجعة بيانات الطلب':'Verify the request details');
  if(r.outcome==='Partial service')needs.push(ar?'متابعة الجزء المتبقي من الخدمة':'Follow up the outstanding service');
  if(r.next)needs.push(r.next+(r.due?' ('+dateLabel(r.due)+')':''));
  if(!needs.length)needs.push(ar?'راجع الطلب وحدد الإجراء التالي':'Review the request and record the next action');
  actions.push(ref+needs.join(' · '));
 }
 return actions.join('\n')||narratives(p,requests)[ar?'ar':'en'];
}
export const contactResults=['Answered','No answer','Unavailable number','Wrong number','Callback requested','Message sent — awaiting acknowledgement'];
export function validateContact(c,data){
 const r=data.requests.find(r=>r.key===c.requestKey);if(!r||r.patientKey!==c.patientKey)throw Error('Choose the request this contact concerns.');
 if(!isOpen(r)&&(c.informed||c.nextAction||c.nextDue))throw Error('Closed requests allow historical contact notes only.');
 if(!contactResults.includes(c.result))throw Error('Choose a contact result.');
 if(!['Phone','WhatsApp','In person','Other'].includes(c.channel))throw Error('Choose a contact channel.');
 if(!['Notification','Service follow-up','Data verification','Other'].includes(c.purpose))throw Error('Choose a contact purpose.');
 if(!c.contactAt||Number.isNaN(Date.parse(c.contactAt))||new Date(c.contactAt).getTime()>Date.now()+60000)throw Error('Enter a valid contact date/time, not in the future.');
 if(c.informed&&c.result!=='Answered')throw Error('Confirm acknowledgement only after an answered contact.');
 if(c.nextDue&&!/^\d{4}-\d{2}-\d{2}$/.test(c.nextDue))throw Error('Choose a valid next-action date.');
}
export function appendContact(data,c,actor){validateContact(c,data);if(data.contactOutcomes?.some(row=>row.key===c.key))throw Error('This contact has already been recorded.');const row={schemaVersion:1,...c,recordedAt:new Date().toISOString(),staff:actor};(data.contactOutcomes||=[]).push(row);const r=data.requests.find(r=>r.key===row.requestKey);if(row.informed)r.informed='Informed';if(row.nextAction)r.next=row.nextAction;if(row.nextDue)r.due=row.nextDue;if(row.informed||row.nextAction||row.nextDue){r.version++;r.updatedAt=row.recordedAt;}return row;}
export function csvContacts(rows){const fields=['key','patientKey','requestKey','contactAt','recordedAt','staff','staffId','requestReference','channel','purpose','result','notes','language','informed','nextAction','nextDue',...(rows.some(r=>Object.hasOwn(r,'patientName'))?['patientName','patientID','phone','contactNumber']:[])];const cell=value=>{let text=String(value??'');if(/^[\s]*[=+@-]/.test(text))text="'"+text;return '"'+text.replace(/"/g,'""')+'"';};return '\ufeff'+[fields.map(cell).join(','),...rows.map(row=>fields.map(f=>cell(row[f])).join(','))].join('\r\n');}
export const messageTemplates={
 en:{appointment:'Hello {patient_name}, your appointment for {approved_procedure} is at {hospital} on {appointment_date}. Please confirm you received these details.',followup:'Hello {patient_name}, we are following up on your request ({request_reference}). Please tell us what service you received and whether anything is still needed.',general:'Hello {patient_name}, this is the Egyptian Red Crescent following up on your request ({request_reference}). Please contact us to complete the required details.'},
 ar:{appointment:'مرحباً {patient_name}، موعد {approved_procedure} في {hospital} بتاريخ {appointment_date}. يرجى تأكيد استلام تفاصيل الموعد.',followup:'مرحباً {patient_name}، نتابع طلبكم ({request_reference}). يرجى إبلاغنا بالخدمة التي تلقيتموها وما إذا كانت هناك احتياجات متبقية.',general:'مرحباً {patient_name}، معكم الهلال الأحمر المصري لمتابعة طلبكم ({request_reference}). يرجى التواصل لاستكمال البيانات المطلوبة.'}
};
export function fillMessage(template,p,r){const values={patient_name:p.name,approved_procedure:r.approved,hospital:r.hospital,appointment_date:dateLabel(r.appointment),request_reference:r.reference};return template.replace(/\{(\w+)\}/g,(match,key)=>values[key]||match);}
export function whatsappNumber(value){let n=normalizeId(value).replace(/[()+.\-]/g,'');if(n.startsWith('00'))n=n.slice(2);if(/^01\d{9}$/.test(n))n='20'+n.slice(1);if(!/^[1-9]\d{7,14}$/.test(n))throw Error('Enter a WhatsApp number with country code, or an Egyptian mobile number.');return n;}
