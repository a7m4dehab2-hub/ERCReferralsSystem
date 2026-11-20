import { clean, normalizeId, searchText, validId, todayCairo, patientErrors, requestErrors } from './domain.mjs';
export const importFields = [
 ['name','Patient name',['الاسم الرباعي','الاسم - Name','Patient Name - اسم المريض','Patient name']],
 ['id','Patient ID',['رقم الهوية الفلسطينية','رقم الهوية- ID','Patient ID - رقم الهوية','Patient ID']],
 ['phone','WhatsApp number',['رقم واتساب مصري','رقم هاتف الحالة','Phone Number - رقم الهاتف','Phone']],
 ['otherPhone','Other phone',['رقم تواصل مصري آخر (اختياري)','Other phone']],
 ['centre','Accommodation centre',['مركز الإقامة','مركز الإقامة/ Accommodation centre','Accommodation centre']],
 ['building','Building',['رقم العمارة','العمارة','Building']], ['apartment','Apartment',['رقم الشقة','الشقة','Apartment']],
 ['gender','Gender',['الجنس','Gender - الجنس','Gender']], ['birth','Birth date',['تاريخ الميلاد','Birth date']],
 ['reportedAge','Reported age',['العمر','Age - العمر','Age']], ['beneficiary','Beneficiary type',['نوع المستفيد','Beneficiary type']],
 ['entryDate','Egypt entry date',['تاريخ الدخول من المعبر إلى مصر','Egypt entry date','Arrival date']],
 ['requested','Requested procedure',['Procedure requested','Requested procedure','التوصيات/ الإجراءات الطبية المطلوبة/ Medical recommendations']],
 ['approved','Approved procedure',['Procedure approved','Approved procedure','الإجراء المعتمد']],
 ['received','Service received',['الخدمة الصحية المقدمة','Service received']],
 ['hospital','Hospital',['Hospital','Hospital - المستشفى','المستشفى']],
 ['appointment','Appointment date',['Appointment Date - موعد الخدمة/ التنسيق','Appointment date']],
 ['specialty','Specialty',['Specialty','Specialty - التخصص']],
 ['diagnosis','Diagnosis',['التشخيص من الهلال الاحمر','Diagnosis - التشخيص','Diagnosis']],
 ['doctor','Referring doctor',['Doctor Name - اسم الطبيب / اسم المحيل','Referring doctor']],
 ['reportDate','Report date',['تاريخ التقرير - date of the report','Report date']],
 ['reportUrl','Report link',['ملف التقرير المعتمد - report file','Report link']],
 ['agency','External referral agency',['اسم الجهة الخارجية','External referral agency']],
 ['legacyStatus','Original service status',['Service Status - حالة الخدمة الكلية','Service Status','Status']],
 ['sourceDate','Original referral/email date',['تاريخ الايميل','Referral Date - تاريخ الإحالة']],
 ['notes','Written outcome',['Notes','Notes1','Written outcome']]
].map(([key,label,aliases]) => ({ key,label,aliases }));
const patientKeys = importFields.slice(0,12).map(f=>f.key);
const dateKeys = new Set(['birth','entryDate','appointment','reportDate','sourceDate']);
export function sheetGrid(XLSX,sheet,{date1904=false}={}){
 const grid=[];
 // Read typed cells directly: a date format applied to a whole Excel column
 // must not turn its text header into Invalid Date.
 for(const [address,cell]of Object.entries(sheet))if(!address.startsWith('!')&&cell.v!=null&&cell.v!==''){
  const position=XLSX.utils.decode_cell(address);
  if(position.r>=50000)throw Error('This sheet has more than 50,000 populated rows. Import it in smaller files.');
  grid[position.r] ||= [];
  grid[position.r][position.c]=cell.t==='e'?(cell.w||'#ERROR'):cell.v;
  // Respect explicit zero padding; never guess missing ID digits.
  if(cell.t==='n'&&/^0{2,}$/.test(cell.z||''))grid[position.r][position.c]=XLSX.SSF.format(cell.z,cell.v);
  else if(cell.t==='n'&&XLSX.SSF.is_date(cell.z||'')){
   const date=XLSX.SSF.parse_date_code(cell.v,{date1904});
   if(date)grid[position.r][position.c]=new Date(date.y,date.m-1,date.d,date.H,date.M,date.S);
  }
 }
 return grid;
}
export function autoMapping(headers) {
 const map = {};
 for (const field of importFields) { const index = headers.findIndex(h=>field.aliases.some(a=>searchText(a)===searchText(h))); if(index>=0)map[field.key]=index; }
 return map;
}
export function findHeader(grid) {
 return grid.slice(0,50).findIndex(row=>{const mapping=autoMapping(row||[]);return mapping.id !== undefined && mapping.name !== undefined;});
}
export function importDate(value) {
 // Excel calendar dates have no timezone. SheetJS creates local Date objects;
 // UTC conversion would move Cairo midnight to the previous calendar day.
 if(value instanceof Date && !Number.isNaN(+value))return `${value.getFullYear()}-${String(value.getMonth()+1).padStart(2,'0')}-${String(value.getDate()).padStart(2,'0')}`;
 const text=normalizeId(value);
 let match=/^(\d{4})-(\d{1,2})-(\d{1,2})(?:T.*)?$/.exec(text), result;
 if(match) result=`${match[1]}-${match[2].padStart(2,'0')}-${match[3].padStart(2,'0')}`;
 else { match=/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})(?:\s.*)?$/.exec(text); if(match)result=`${match[3]}-${match[2].padStart(2,'0')}-${match[1].padStart(2,'0')}`; }
 return result && !Number.isNaN(Date.parse(result)) && new Date(result+'T12:00:00Z').toISOString().slice(0,10)===result ? result : '';
}
export const stable = value => value && typeof value === 'object' && !Array.isArray(value) ? Object.fromEntries(Object.keys(value).sort().map(k=>[k,stable(value[k])])) : Array.isArray(value) ? value.map(stable) : value;
export async function digest(value) { return [...new Uint8Array(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(JSON.stringify(stable(value)))))].map(n=>n.toString(16).padStart(2,'0')).join(''); }
export async function parseRows({ grid, headers, mapping, headerIndex, fileName, fileHash, sheet, kind, links = {} }) {
 if(mapping.id===undefined || mapping.name===undefined)throw Error('Map Patient ID and Patient name before previewing.');
 const rows=[];
 for(let index=headerIndex+1;index<grid.length;index++){
  const raw=grid[index] || [], values={}, issues=[];
  for(const [key,col] of Object.entries(mapping)) {
   const input=raw[col] ?? '';
   if(dateKeys.has(key)){ values[key]=importDate(input); if(input!==''&&!values[key])issues.push(importFields.find(f=>f.key===key).label+': unrecognized date'); }
   else values[key]=key==='id'?normalizeId(input):clean(input);
   if(key==='reportUrl' && links[`${index}:${col}`])values[key]=links[`${index}:${col}`];
  }
  if(![values.name,values.id,values.requested,values.approved,values.received,values.hospital].some(Boolean))continue;
  if(!values.id)issues.push('Missing ID');else if(!validId(values.id))issues.push('ID needs verification');
  if(!values.name)issues.push('Missing patient name');
  values.originalId=clean(raw[mapping.id]);
  if(kind==='contacts'){const stamp=headers.findIndex(header=>/(وقت.*(الإرسال|الارسال)|timestamp|submitted at)/i.test(String(header||'')));if(stamp>=0)values.ageRecorded=importDate(raw[stamp]);}
  const cells=raw.map((value,col)=>({column:col+1,header:String(headers[col]||''),value:value instanceof Date?importDate(value)+(value.getHours()||value.getMinutes()||value.getSeconds()?` ${String(value.getHours()).padStart(2,'0')}:${String(value.getMinutes()).padStart(2,'0')}:${String(value.getSeconds()).padStart(2,'0')}`:''):String(value??''),...(links[`${index}:${col}`]?{link:links[`${index}:${col}`]}:{})})).filter(c=>c.value||c.link);
  // Timestamp-only form re-submissions are identical business entries.
  const businessCells=cells.filter(c=>!/(وقت.*(الإرسال|الارسال)|timestamp|submitted at)/i.test(c.header)).map(c=>{const key=Object.keys(mapping).find(k=>mapping[k]===c.column-1);return {header:searchText(c.header),value:dateKeys.has(key)?importDate(c.value):['id','phone','otherPhone'].includes(key)?normalizeId(c.value):clean(c.value),link:c.link||''};});
  const fingerprint=await digest({kind,cells:businessCells});
  rows.push({key:crypto.randomUUID(),fingerprint,kind,values,issues,status:'pending',source:{fileName,fileHash,sheet,row:index+1,cells},createdAt:new Date().toISOString()});
 }
 return rows;
}
export function classify(row,data,peers=[]){
 if((data.imports||[]).some(r=>r.key!==row.key&&r.fingerprint===row.fingerprint))return 'Identical entry — skipped';
 if(!row.values.id)return 'Missing ID — investigation';
 const existing=data.patients.find(p=>normalizeId(p.id)===row.values.id);
 if(!existing && !row.values.name)return 'Missing name — investigation';
 if(peers.some(r=>r.key!==row.key&&r.values.id===row.values.id&&r.fingerprint!==row.fingerprint))return 'Multiple submissions — review';
 if(existing && patientKeys.some(k=>clean(existing[k])&&clean(row.values[k])&&clean(existing[k])!==clean(row.values[k])))return 'Conflicting details — review';
 if(row.kind==='referrals'&&existing&&data.requests.some(r=>r.patientKey===existing.key))return 'Existing requests — choose request';
 if(row.issues.some(issue=>!issue.startsWith('ID needs')))return 'Source values need review';
 if(row.kind==='contacts'&&existing&&patientKeys.every(k=>!clean(row.values[k])||clean(existing[k])===clean(row.values[k])))return 'Identical profile — skipped';
 return row.kind==='contacts'?(existing?'Fill missing details':'New patient'):'New request — verify details';
}
export const isSafe = classification => ['New patient','Fill missing details','New request — verify details'].includes(classification);
export function applySourceRow(data,row,{ values=row.values, useFields=patientKeys.filter(k=>clean(values[k])), useRequestFields=Object.keys(values), requestKey='', progress='Under verification', informed='Not informed', outcome='Not confirmed', reason='' }={}) {
 if(row.status==='applied' || data.imports?.find(r=>r.key===row.key)?.status==='applied')throw Error('This source row has already been applied.');
 if(!normalizeId(values.id))throw Error('A supplied ID is required. Keep this row in investigation until it is found.');
 let p=data.patients.find(p=>normalizeId(p.id)===normalizeId(values.id)), previous=p?structuredClone(p):null;
 if(!p)p={key:crypto.randomUUID(),version:0,name:'',id:normalizeId(values.id),originalId:row.values.originalId||values.id,phone:'',otherPhone:'',centre:'',building:'',apartment:'',birth:'',reportedAge:'',entryDate:''};
 for(const key of patientKeys) if(useFields.includes(key))p[key]=key==='id'?normalizeId(values[key]):clean(values[key]);
 p.id=normalizeId(values.id);
 const errors=patientErrors(p,data.patients,p.key); if(Object.keys(errors).length)throw Error('Complete the patient name, ID and valid personal dates before applying.');
 if(p.reportedAge&&(!p.ageRecorded||(useFields.includes('reportedAge')&&p.reportedAge!==previous?.reportedAge)))p.ageRecorded=values.ageRecorded||'';
 p.version++;p.updatedAt=new Date().toISOString();p.createdAt ||= p.updatedAt;
 if(previous)data.patients=data.patients.map(value=>value.key===p.key?p:value);else data.patients.push(p);
 let request=null;
 if(row.kind==='referrals'){
  const existingRequests=data.requests.filter(r=>r.patientKey===p.key);
  if(existingRequests.length&&!requestKey)throw Error('Choose an existing request or explicitly choose New independent request.');
  request=requestKey&&requestKey!=='new'?structuredClone(existingRequests.find(r=>r.key===requestKey)):null;
  if(requestKey&&requestKey!=='new'&&!request)throw Error('The selected request does not belong to this patient.');
  const old=request?structuredClone(request):null;
  request ||= {key:crypto.randomUUID(),version:0,patientKey:p.key,reference:'REQ-'+String(data.requests.length+1).padStart(6,'0'),source:'Legacy migration',contacts:[],owner:'',urgent:false,verified:false,createdDate:todayCairo()};
  for(const key of ['requested','approved','received','hospital','appointment','specialty','diagnosis','doctor','reportDate','reportUrl','agency','legacyStatus','notes'])if(useRequestFields.includes(key)&&Object.hasOwn(values,key))request[key]=values[key];
  Object.assign(request,{progress,informed,outcome,sourceRow:`${row.source.fileName} / ${row.source.sheet} / ${row.source.row}`,sourceDate:useRequestFields.includes('sourceDate')?values.sourceDate||'':request.sourceDate||'',due:request.due||'',next:request.next||'',closureReason:values.closureReason||request.closureReason||'',version:request.version+1,updatedAt:new Date().toISOString()});
  if(old && ['Completed','Closed'].includes(old.progress) && !['Completed','Closed'].includes(progress) && (!reason||!values.next||!values.due))throw Error('Reopening an existing request requires a reason, next action and date. Edit it from Requests first.');
  const problems=requestErrors(request);if(Object.keys(problems).length)throw Error('Check the procedure, report verification, closure reason or full-service confirmation.');
  if(progress==='Completed')request.completedDate=values.completedDate||request.completedDate||'';
  if(progress==='Closed')request.closedDate=values.closedDate||request.closedDate||'';
  if(old)data.requests=data.requests.map(value=>value.key===request.key?request:value);else data.requests.push(request);
  for(const [key,field] of [['hospitals','hospital'],['agencies','agency']])if(request[field]&&!data.settings[key].some(e=>e.name===request[field]))data.settings[key].push({key:crypto.randomUUID(),name:request[field],arabic:''});
 }
 const saved={...row,values,status:'applied',patientKey:p.key,requestKey:request?.key||'',resolvedAt:new Date().toISOString()};
 const existingIndex=(data.imports||=[]).findIndex(r=>r.key===row.key);
 if(existingIndex>=0)data.imports[existingIndex]=saved;else data.imports.push(saved);
 return {patient:p,previous,request};
}
