export const normalizeId = value => String(value ?? '').replace(/[٠-٩۰-۹]/g, c => '٠١٢٣٤٥٦٧٨٩'.includes(c) ? String('٠١٢٣٤٥٦٧٨٩'.indexOf(c)) : String('۰۱۲۳۴۵۶۷۸۹'.indexOf(c))).replace(/\s+/g, '');
export const clean = value => String(value ?? '').trim().replace(/\s+/g, ' ');
// Search folding never changes stored names or identity matching.
export const searchText = value => normalizeId(clean(value).toLowerCase().normalize('NFKC').replace(/[\u064b-\u065f\u0670\u06d6-\u06edـ]/g, '').replace(/[أإآٱ]/g, 'ا').replace(/[ىئ]/g, 'ي').replace(/ؤ/g, 'و').replace(/ة/g, 'ه'));
export function matchesDemographics(patient, filters, now) {
  const age = ageAt(patient.birth, now) ?? (patient.reportedAge !== '' && patient.reportedAge != null ? Number(patient.reportedAge) : null);
  if (filters.ageMin !== '' && (age === null || age < Number(filters.ageMin))) return false;
  if (filters.ageMax !== '' && (age === null || age > Number(filters.ageMax))) return false;
  if (filters.arrivalFrom && (!patient.entryDate || patient.entryDate < filters.arrivalFrom)) return false;
  if (filters.arrivalTo && (!patient.entryDate || patient.entryDate > filters.arrivalTo)) return false;
  return true;
}
export function validId(value) {
  const id = normalizeId(value);
  if (!/^[0-9]{9}$/.test(id)) return false;
  return [...id].reduce((sum, digit, i) => { const n = Number(digit) * (i % 2 ? 2 : 1); return sum + (n > 9 ? n - 9 : n); }, 0) % 10 === 0;
}
export function makeId(base) {
  const eight = String(base).padStart(8, '0').slice(-8);
  const sum = [...eight].reduce((s, d, i) => { const n = Number(d) * (i % 2 ? 2 : 1); return s + (n > 9 ? n - 9 : n); }, 0);
  return eight + String((10 - sum % 10) % 10);
}
export function todayCairo() {
  const parts = new Intl.DateTimeFormat('en-CA', { timeZone: 'Africa/Cairo', year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(new Date());
  return ['year', 'month', 'day'].map(type => parts.find(p => p.type === type).value).join('-');
}
export function offsetDate(date, days) { const d = new Date(date + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + days); return d.toISOString().slice(0, 10); }
export function dateLabel(value) { return /^\d{4}-\d{2}-\d{2}$/.test(value || '') ? value.split('-').reverse().join('/') : String(value || ''); }
export function ageAt(birth, now) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(birth || '') || birth > now) return null;
  let years = Number(now.slice(0, 4)) - Number(birth.slice(0, 4));
  if (now.slice(5) < birth.slice(5)) years--;
  return years;
}
export function patientErrors(patient, patients, excluding = null) {
  const errors = {};
  if (!clean(patient.name)) errors.name = 'nameRequired';
  if (!normalizeId(patient.id)) errors.id = 'idRequired';
  else if (patients.some(p => p.key !== excluding && normalizeId(p.id) === normalizeId(patient.id))) errors.id = 'idExists';
  if (patient.birth && (Number.isNaN(Date.parse(patient.birth)) || patient.birth > todayCairo())) errors.birth = 'birthInvalid';
  if (patient.reportedAge !== '' && patient.reportedAge !== undefined && (!Number.isFinite(Number(patient.reportedAge)) || Number(patient.reportedAge) < 0)) errors.reportedAge = 'ageInvalid';
  return errors;
}
export function missingDetails(patient, necessary) { return necessary.filter(key => !clean(patient[key])); }
export function isOpen(request) { return !['Completed', 'Closed'].includes(request.progress); }
export function queueMatches(request, patient, queue, settings, now) {
  if (queue === 'all') return isOpen(request);
  if (queue === 'ready') return request.progress === 'Ready to send';
  if (queue === 'awaiting') return request.progress === 'Awaiting approval';
  if (queue === 'appointment') return isOpen(request) && !!request.approved && (!request.hospital || !request.appointment);
  if (queue === 'notification') return isOpen(request) && !!request.approved && !!request.hospital && !!request.appointment && request.informed !== 'Informed';
  if (queue === 'followups') return isOpen(request) && !!request.due;
  if (queue === 'investigations') return isOpen(request) && (request.progress === 'Under verification' || missingDetails(patient, settings.necessary).length > 0 || !validId(patient.id));
  if (queue === 'completed') return request.progress === 'Completed';
  if (queue === 'closed') return request.progress === 'Closed';
  if (queue === 'overview') return isOpen(request) && (request.urgent || (!!request.approved && (!request.hospital || !request.appointment)) || (!!request.due && request.due <= now));
  return true;
}
export function priorityRank(request, now) { return request.urgent ? 0 : request.approved && (!request.hospital || !request.appointment) ? 1 : request.due && request.due < now ? 2 : request.due === now ? 3 : 4; }
export function narratives(patient, requests) {
  const open = requests.filter(isOpen);
  const en = [`${patient.name} has ${requests.length} request${requests.length === 1 ? '' : 's'}; ${open.length} open.`];
  const ar = [`لدى ${patient.name} عدد ${requests.length} من الطلبات، منها ${open.length} طلب مفتوح.`];
  for (const r of requests) {
    en.push(`${r.reference}: ${r.requested || 'Requested procedure not recorded'}. ${r.approved ? 'Approved: ' + r.approved + '.' : 'Approval not recorded.'} ${r.appointment ? 'Appointment: ' + dateLabel(r.appointment) + '.' : 'Appointment not recorded.'} ${r.informed === 'Informed' ? 'Patient informed.' : 'Patient notification not confirmed.'} ${r.outcome === 'Full service' && r.progress === 'Completed' ? 'Full service confirmed.' : r.outcome === 'Partial service' ? 'Partial service; outstanding work remains.' : 'Full service not confirmed.'}${r.next ? ' Next: ' + r.next + (r.due ? ' (' + dateLabel(r.due) + ')' : '') + '.' : ''}`);
    ar.push(`${r.reference}: ${r.requested || 'الإجراء المطلوب غير مسجل'}. ${r.approved ? 'الموافقة المسجلة: ' + r.approved + '.' : 'الموافقة غير مسجلة.'} ${r.appointment ? 'الموعد: ' + dateLabel(r.appointment) + '.' : 'الموعد غير مسجل.'} ${r.informed === 'Informed' ? 'تم إبلاغ المريض.' : 'إبلاغ المريض غير مؤكد.'} ${r.outcome === 'Full service' && r.progress === 'Completed' ? 'تم تأكيد تلقي الخدمة كاملة.' : r.outcome === 'Partial service' ? 'تم تلقي جزء من الخدمة وما زال هناك عمل مطلوب.' : 'تلقي الخدمة كاملة غير مؤكد.'}${r.next ? ' الإجراء التالي: ' + r.next + (r.due ? ' (' + dateLabel(r.due) + ')' : '') + '.' : ''}`);
  }
  return { en: en.join('\n\n'), ar: ar.join('\n\n') };
}
export function requestErrors(request) {
  const errors = {};
  if (request.source === 'Clinic referral' && !clean(request.requested)) errors.requested = 'procedureRequired';
  if (request.source === 'Direct ministry approval' && !clean(request.approved)) errors.approved = 'approvalRequired';
  if (request.progress === 'Ready to send') for (const key of ['doctor', 'reportDate', 'reportUrl']) if (!clean(request[key])) errors[key] = 'verificationRequired';
  if (request.progress === 'Ready to send' && !request.verified) errors.verified = 'verificationRequired';
  if (request.progress === 'Completed' && request.outcome !== 'Full service') errors.outcome = 'fullServiceRequired';
  if (request.progress === 'Closed' && !clean(request.closureReason)) errors.closureReason = 'closureRequired';
  for (const key of ['reportUrl', 'documentUrl']) if (request[key]) { try { if (new URL(request[key]).protocol !== 'https:') errors[key] = 'httpsRequired'; } catch { errors[key] = 'httpsRequired'; } }
  return errors;
}

