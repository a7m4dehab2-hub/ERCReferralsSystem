import { normalizeId, validId, searchText, matchesDemographics, clean, todayCairo, dateLabel, ageAt, missingDetails, queueMatches, priorityRank, narratives, patientErrors, requestErrors, isOpen } from './domain.mjs';
import { readData, persist, readPreferences, savePreferences, changeEvents, refreshData } from './store.mjs';
import { cloudMode, member, initialiseCloud, signIn, signOut, emptyData, restoreCloud } from './cloud.mjs';
import { renderImports, resetImports, prepareSheetUpdate } from './import-ui.mjs';
import { simpleWorkspace, patientPanel } from './simple-ui.mjs';
import { contactForm, whatsappForm, outcomesView, resetContactUI } from './contact-ui.mjs';
import { workMatch } from './workspace.mjs';
import { arabic } from './i18n.mjs';

const root = document.getElementById('app');
const tooltip = document.getElementById('tooltip');
const dialog = document.getElementById('detail-dialog');
const dialogTitle = document.getElementById('dialog-title');
const dialogContent = document.getElementById('dialog-content');
const now = todayCairo();
let data = readData();
let preferences = readPreferences();
const state = { tab: 'overview', user: '', search: '', ageMin: '', ageMax: '', arrivalFrom: '', arrivalTo: '', priority: 'All', progress: 'All', owner: 'All staff', notification: 'All', followup: 'Today + overdue', scope: 'Team work', selected: null, editor: null, columns: false, expanded: new Set(), settingsTab: 'Accounts', reportFrom: now.slice(0, 7) + '-01', reportTo: now, conflict: null };
let toastTimer;
let tooltipTimer;
const patientFields = [
  ['name', 'Patient name'], ['id', 'Patient ID'], ['phone', 'WhatsApp number'], ['otherPhone', 'Other phone'], ['centre', 'Accommodation centre'], ['building', 'Building'], ['apartment', 'Apartment'], ['gender', 'Gender', 'select', ['', 'Male', 'Female']], ['birth', 'Birth date', 'date'], ['age', 'Age', 'computed'], ['reportedAge', 'Reported age', 'number'], ['ageRecorded', 'Age recorded on', 'date'], ['beneficiary', 'Beneficiary type'], ['entryDate', 'Egypt entry date', 'date']
].map(([key, label, type = 'text', options]) => ({ id: 'p:' + key, key, label, type, options, scope: 'p' }));
const requestFields = [
  ['reference', 'Request number', 'computed'], ['source', 'Request source', 'select', ['Clinic referral', 'Direct ministry approval', 'Legacy migration']], ['requested', 'Requested procedure'], ['approved', 'Approved procedure'], ['received', 'Service received'], ['progress', 'Request progress', 'select', ['Under verification', 'Ready to send', 'Awaiting approval', 'Approved', 'Completed', 'Closed']], ['agency', 'External referral agency', 'agency'], ['hospital', 'Hospital', 'hospital'], ['appointment', 'Appointment date', 'date'], ['informed', 'Patient informed?', 'select', ['Not informed', 'Awaiting confirmation', 'Informed', 'No answer', 'Unavailable number', 'Wrong number', 'Callback requested']], ['next', 'Next action'], ['due', 'Due date', 'date'], ['owner', 'Responsible staff', 'select', ['']], ['urgent', 'Urgent', 'checkbox'], ['specialty', 'Specialty'], ['diagnosis', 'Diagnosis', 'long'], ['doctor', 'Referring doctor'], ['reportDate', 'Report date', 'date'], ['reportUrl', 'Report link'], ['documentUrl', 'Document link'], ['responseDate', 'Response date', 'date'], ['verified', 'Report verified', 'checkbox'], ['outcome', 'Service outcome', 'select', ['Not confirmed', 'Full service', 'Partial service', 'No service']], ['closureReason', 'Closure reason', 'long'], ['notes', 'Written outcome', 'long'], ['legacyStatus', 'Original service status'], ['sourceRow', 'Source row', 'computed']
].map(([key, label, type = 'text', options]) => ({ id: 'r:' + key, key, label, type, options, scope: 'r' }));
const derivedFields = [
  { id: 'd:requests', key: 'requests', label: 'Requests', type: 'requests' },
  { id: 'd:contacts', key: 'contacts', label: 'Contact history', type: 'long' },
  { id: 'd:history', key: 'history', label: 'Change history', type: 'long' },
  { id: 'd:en', key: 'en', label: 'English summary', type: 'long' },
  { id: 'd:ar', key: 'ar', label: 'Arabic summary', type: 'long' }
];
const navItems = [['overview', 'Daily overview'], ['patients', 'Patients'], ['all', 'All requests'], ['imports', 'Import & review'], ['agencies', 'External referral agencies'], ['outcomes', 'Contact outcomes'], ['reports', 'Reports'], ['settings', 'Settings']];
const queueItems = [['overview', 'Priority work'], ['ready', 'Ready to send'], ['awaiting', 'Awaiting approval'], ['appointment', 'Needs appointment'], ['notification', 'Needs notification'], ['followups', 'Follow-ups'], ['investigations', 'Investigations'], ['completed', 'Completed requests'], ['closed', 'Closed requests']];
const helpText = {
  outcomes: 'Collect and export contact outcomes and original notes, linked to each patient and request.',
  imports: 'Import contact submissions and legacy referral rows, compare differences and review records held for investigation.',
  agencies: 'Maintain the external referral agency directory and see how many requests are associated with each agency.',
  overview: 'See urgent, incomplete approvals, overdue and today’s work. A request can appear in more than one queue.',
  patients: 'Find one patient profile. Patient details are shared by all their independent requests.',
  all: 'View open requests with their separate approval, contact, service and next-action details.',
  ready: 'Verified requests ready to export and submit manually. Exporting is not marking them sent.',
  awaiting: 'Submitted requests waiting for a recorded external agency response. Omitted requests stay here.',
  appointment: 'Approved requests missing a hospital or appointment date. Obtain the missing details promptly.',
  notification: 'Appointments still needing patient notification or acknowledgement. Opening WhatsApp does not mark informed.',
  followups: 'Due, overdue and upcoming actions. Successful contact alone does not prove full service delivery.',
  investigations: 'Missing necessary personal details, invalid-ID warnings or report-verification work.',
  completed: 'Requests explicitly completed after full service was confirmed.',
  closed: 'Requests closed without completion, with a reason. This is separate from successful delivery.',
  reports: 'Count events by their own dates, and distinguish the current open workload.',
  settings: 'Manage workspace settings. Administrators maintain necessary details, hospitals and referral agencies.',
  columns: 'Choose visible fields for this user and table. Hiding a column does not delete data.',
  direction: 'Set this table’s direction independently of the interface language. Your choice is remembered.',
  help: 'Show or hide these information icons and explanations. Essential labels and warnings always remain.',
  addPatient: 'Create a profile with name and a supplied ID. A missing ID is not allowed; an invalid supplied ID warns.',
  addRequest: 'Enter an ID first. Existing patient details are pulled into the row; a new patient requires a name.',
  edit: 'Edit the selected record and explicitly Save or Cancel. Shared patient edits apply to their requests.',
  save: 'Validate and save your changes, with old/new values in the edit history. Connected workspaces save to Supabase.'
};

const isAdmin = () => member?.role === 'admin';
const staffNames = () => (data.staff || []).map(s => s.name);
function userPrefs() { return preferences[cloudMode && member ? member.id : state.user] ||= { language: 'en', help: true, tables: {} }; }
function simpleMode(){return (userPrefs().view || 'simple') === 'simple';}
function tablePrefs() { return userPrefs().tables[state.tab] ||= { hidden: [], direction: userPrefs().language === 'ar' ? 'rtl' : 'ltr' }; }
function t(text) { if(cloudMode && text === 'Saved locally') text = 'Saved'; if (text === 'Direct ministry approval') text = 'Direct agency approval'; return userPrefs().language === 'ar' ? arabic[text] || text : text; }
function node(tag, attrs = {}, ...children) {
  const item = document.createElement(tag);
  for (const [key, value] of Object.entries(attrs)) {
    if (value === null || value === undefined) continue;
    if (key === 'class') item.className = value;
    else if (key === 'text') item.textContent = value;
    else if (key === 'on') for (const [event, handler] of Object.entries(value)) item.addEventListener(event, handler);
    else if (key === 'checked' || key === 'disabled' || key === 'hidden') item[key] = !!value;
    else if (key === 'value') item.value = value;
    else item.setAttribute(key, String(value));
  }
  for (const child of children.flat()) if (child !== undefined && child !== null) item.append(child instanceof Node ? child : document.createTextNode(String(child)));
  return item;
}
function info(text, label) {
  if (!userPrefs().help) return null;
  const button = node('button', { type: 'button', class: 'help-icon', 'aria-label': t('Help tips') + ': ' + t(label), 'data-help': text }, node('span', { 'aria-hidden': 'true', text: 'i' }));
  return button;
}
function action(label, handler, { primary = false, disabled = false, help, className = '', id } = {}) {
  const button = node('button', { type: 'button', class: `${primary ? 'primary' : ''} ${className}`.trim(), disabled, id, text: t(label), on: { click: event => Promise.resolve().then(() => handler(event)).catch(error => toast(error.message, true)) } });
  return node('span', { class: 'tool-group' }, button, help ? info(help, label) : null);
}
function select(label, values, current, handler, attrs = {}) {
  const input = node('select', { 'aria-label': t(label), ...attrs });
  if (cloudMode && label === 'Responsible staff') values = values[0] === 'All staff' ? ['All staff', ...staffNames(), ''] : ['', ...staffNames()];
  values.forEach(value => input.append(node('option', { value, text: t(value || 'Shared queue') })));
  input.value = current;
  input.addEventListener('change', () => handler(input.value));
  return node('label', { class: 'small-label' }, t(label), input);
}
function remember() { try { preferences.activeUser = state.user; savePreferences(preferences); } catch { toast('Not saved — local storage unavailable. Keep this page open and retry.', true); } }
function toast(message, error = false) {
  const item = document.getElementById('toast'); item.className = error ? 'error' : ''; item.textContent = t(message); item.hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => item.hidden = true, 5000);
}
function showDialog(title, content, actions = []) {
  dialog.classList.remove('profile-panel');
  dialogTitle.textContent = t(title); dialogContent.replaceChildren(content, actions.length ? node('div', { class: 'dialog-actions' }, actions) : ''); dialog.showModal();
}
document.getElementById('close-dialog').addEventListener('click', () => dialog.close());
function navigate(tab, search = '') {
  if (state.editor?.dirty && !confirm(t('Discard unsaved changes?'))) return false;
  state.simpleFilters = [];
  state.tab = tab; state.search = search; state.editor = null; state.selected = null; state.columns = false; state.conflict = null; state.priority = 'All'; state.progress = 'All'; state.notification = 'All'; render(); return true;
}
function allFields() {
  if (state.tab === 'patients') return [...patientFields, ...derivedFields.filter(f => f.key !== 'contacts')];
  const p = id => patientFields.find(f => f.key === id), r = id => requestFields.find(f => f.key === id);
  const first = [p('name'), p('id'), p('phone'), r('reference'), r('requested'), r('approved'), r('received'), r('progress'), r('hospital'), r('appointment'), r('informed'), r('next'), r('due'), r('owner'), r('urgent')];
  return [...first, ...patientFields.filter(f => !first.includes(f)), ...requestFields.filter(f => !first.includes(f)), ...derivedFields.filter(f => f.key !== 'requests')];
}
function getRows() {
  const q = searchText(state.search);
  let rows = state.tab === 'patients' ? data.patients.map(p => ({ p, r: null, key: p.key })) : data.requests.map(r => ({ p: data.patients.find(p => p.key === r.patientKey), r, key: r.key })).filter(row => row.p && queueMatches(row.r, row.p, simpleMode() && state.tab === 'overview' ? 'all' : state.tab, data.settings, now));
  rows = rows.filter(({ p, r }) => {
    if (q && ![p.name, p.id, p.phone, p.otherPhone, r?.reference].some(value => searchText(value).includes(q))) return false;
    if (!matchesDemographics(p, state, now)) return false;
    if(state.hospital && state.hospital!=='All' && !(r?r.hospital===state.hospital:data.requests.some(req=>req.patientKey===p.key&&req.hospital===state.hospital)))return false;
    if (!r) return state.priority === 'ID warning' ? !validId(p.id) : state.priority === 'Missing details' ? missingDetails(p, data.settings.necessary).length > 0 : true;
    if(state.simpleFilters?.length && !state.simpleFilters.some(key=>workMatch(r,p,key,data.settings,now)))return false;
    if (state.scope === 'My work' && r.owner !== state.user) return false;
    if (state.owner !== 'All staff' && r.owner !== state.owner) return false;
    if (state.priority === 'Urgent' && !r.urgent) return false;
    if (state.priority === 'Overdue' && !(r.due && r.due < now)) return false;
    if (state.priority === 'Today' && r.due !== now) return false;
    if (state.progress !== 'All' && r.progress !== state.progress) return false;
    if (state.notification !== 'All' && r.informed !== state.notification) return false;
    if (state.tab === 'followups' && state.followup === 'Today + overdue' && r.due > now) return false;
    if (state.tab === 'followups' && state.followup === 'Upcoming' && r.due <= now) return false;
    return true;
  });
  return rows.sort((a, b) => a.r && b.r ? priorityRank(a.r, now) - priorityRank(b.r, now) || (a.r.due || '9999').localeCompare(b.r.due || '9999') || a.r.reference.localeCompare(b.r.reference) : a.p.name.localeCompare(b.p.name));
}
function patientRequests(p) { return data.requests.filter(r => r.patientKey === p.key); }
function historyFor(p, r = null) { return data.audit.filter(event => event.recordKey === p.key || (r ? event.recordKey === r.key : patientRequests(p).some(req => req.key === event.recordKey))); }
function valueFor(field, row) {
  const { p, r } = row;
  if (field.id === 'p:age') { const computed = ageAt(p.birth, now); return computed !== null ? String(computed) : p.reportedAge ? `${p.reportedAge} (${t('Reported age')}, ${dateLabel(p.ageRecorded) || t('Not recorded')})` : ''; }
  if (field.scope === 'p') return p[field.key] ?? '';
  if (field.scope === 'r') return r?.[field.key] ?? '';
  if(field.key==='contacts' && data.contactOutcomes?.some(c=>c.requestKey===r?.key))return data.contactOutcomes.filter(c=>c.requestKey===r.key).map(c=>`${new Date(c.contactAt).toLocaleString('en-GB',{timeZone:'Africa/Cairo'})} · ${c.staff} · ${t(c.result)}\n${c.notes}`).join('\n\n')+(r?.contacts?.length?'\n\n'+r.contacts.map(c=>`${c.date} · ${c.staff}: ${c.notes}`).join('\n'):'');
  if (field.key === 'contacts') return (r?.contacts || []).map(c => `${dateLabel(c.date)} · ${c.staff}: ${t(c.result)}\n${c.notes}`).join('\n\n') || t('No contact attempts recorded');
  if (field.key === 'history') return historyFor(p, r).map(e => `${new Date(e.at).toLocaleString('en-GB', { timeZone: 'Africa/Cairo' })} · ${t(e.actor)}\n${e.field}: ${String(e.before)} → ${String(e.after)}`).join('\n\n') || t('No changes yet');
  if (field.key === 'en' || field.key === 'ar') return narratives(p, patientRequests(p))[field.key];
  return '';
}
function render() {
  hideHelp();
  root.classList.toggle('simple-mode',simpleMode());
  if(!cloudMode)return renderSetup();
  if(!member)return renderLogin();
  document.documentElement.lang = userPrefs().language;
  document.documentElement.dir = userPrefs().language === 'ar' ? 'rtl' : 'ltr';
  const header = node('header', { class: 'app-header' }, node('img', { class: 'erc-logo', src: './assets/erc-logo.webp', alt: t('Egyptian Red Crescent'), width: '96', height: '56' }), node('span', { class: 'app-brand', text: t('ERC External Referral System') }));
  const headerControls = node('div', { class: 'header-controls' }, select('Language', ['English', 'العربية'], userPrefs().language === 'ar' ? 'العربية' : 'English', value => { if (state.editor?.dirty && !confirm(t('Discard unsaved changes?'))) return render(); state.editor = null; userPrefs().language = value === 'العربية' ? 'ar' : 'en'; remember(); render(); }), node('span', { text: state.user }));
  headerControls.prepend(node('label', { class: 'help-toggle' }, node('input', { type: 'checkbox', checked: userPrefs().help, 'aria-label': t('Help tips'), on: { change: event => { userPrefs().help = event.target.checked; remember(); render(); } } }), t('Help tips')), info(helpText.help, 'Help tips'));
  headerControls.append(action('Sign out', async () => { if(state.editor?.dirty && !confirm(t('Discard unsaved changes?')))return; await signOut(); resetImports(); resetContactUI(); data=emptyData(); state.editor=null; state.selected=null; render(); }));
  const viewSwitch=node('div',{class:'view-switch','aria-label':t('View')});
  for(const [value,label]of [['simple','Simple'],['spreadsheet','Spreadsheet']]) viewSwitch.append(node('button',{text:t(label),'aria-pressed':(simpleMode()?'simple':'spreadsheet')===value,on:{click:()=>{if(state.editor?.dirty&&!confirm(t('Discard unsaved changes?')))return;if(state.editor?.dirty)state.editor=null;userPrefs().view=value;remember();render();}}}));
  headerControls.prepend(viewSwitch);
  if(isAdmin())headerControls.append(action('Update data',async()=>{if(!navigate('imports'))return;await prepareSheetUpdate({render,isActive:()=>!!member&&isAdmin()});},{help:'Read the contact sheet, skip identical submissions and review changes before saving.'}));
  header.append(headerControls);
  const banner = node('div', { class: 'live-banner' }, node('span', { text: 'Connected workspace · '+state.user }), action('Refresh records', async()=>{try{data=await refreshData();render();toast('Records refreshed');}catch(error){toast(error.message,true);}}), action('Download backup', downloadBackup));
  const nav = node('nav', { class: 'main-nav', 'aria-label': 'Main navigation' });
  for (const [key, label] of navItems) nav.append(node('span', { class: `nav-group ${state.tab === key ? 'active' : ''}` }, node('button', { type: 'button', text: t(label), 'aria-current': state.tab === key ? 'page' : null, on: { click: () => navigate(key) } }), info(helpText[key], label)));
  const page = node('main', { class: 'workspace' });
  const title = [...navItems, ...queueItems].find(([key]) => key === state.tab)?.[1] || 'Daily overview';
  const heading = node('div', { class: 'page-heading' }, node('div', {}, node('h1', { text: t(title) }), node('p', { class: 'caption', text: dateLabel(now) + ' · ' + t('Live workspace') })));
  if (!['patients', 'settings', 'reports', 'agencies', 'imports', 'outcomes'].includes(state.tab)) heading.append(node('div', { class: 'scope' }, ['Team work', 'My work'].map(label => node('button', { type: 'button', text: t(label), 'aria-pressed': state.scope === label, on: { click: () => { state.scope = label; render(); } } }))));
  page.append(heading);
  if(state.tab==='outcomes') outcomesView(page,workspaceContext());
  else if (state.tab === 'settings') renderSettings(page);
  else if (state.tab === 'reports') renderReports(page);
  else if (state.tab === 'agencies') renderDirectory(page, 'agencies');
  else if (state.tab === 'imports') renderImports(page, { data, admin:isAdmin(), actor:state.user, save:async(next,options)=>{data=await persist(next,options);return data;}, refresh:refreshData, showDialog, closeDialog:()=>dialog.close(), rerender:render });
  else if(simpleMode()) { if(state.tab==='overview') heading.querySelector('h1').textContent=t('Active patients'); simpleWorkspace(page,workspaceContext()); }
  else {
    if (state.tab !== 'patients') page.append(renderCounts());
    if (state.tab === 'investigations' && (data.imports||[]).some(row=>row.status==='pending')) page.append(node('div',{class:'note-box'}, `${data.imports.filter(row=>row.status==='pending').length} source entries are awaiting review. `, action('Review source entries',()=>navigate('imports'))));
    const sheet = node('section', { class: 'sheet', 'aria-label': t(title) });
    if (state.tab !== 'patients') {
      const tabs = node('nav', { class: 'queue-tabs', 'aria-label': 'Work queues' });
      queueItems.forEach(([key, label]) => tabs.append(node('span', { class: `nav-group ${state.tab === key ? 'active' : ''}` }, node('button', { type: 'button', text: t(label), on: { click: () => navigate(key) } }), info(helpText[key], label))));
      sheet.append(tabs);
    }
    renderToolbar(sheet);
    renderFilters(sheet);
    if (state.columns) renderColumns(sheet);
    if (state.editor) renderEditorBanner(sheet);
    renderTable(sheet);
    page.append(sheet);
  }
  root.replaceChildren(header, banner, nav, page, node('footer', { class: 'app-footer', text: t('Live workspace') + ' · ' + t('Change history is independent of backup retention.') }));
  requestAnimationFrame(pinColumns);
}
function renderCounts() {
  const counts = [
    ['ready', 'Ready to send', 'requests'], ['awaiting', 'Awaiting approval', 'requests'], ['appointment', 'Needs appointment', 'requests'], ['notification', 'Needs notification', 'appointments'], ['today', 'Follow-ups today', 'tasks'], ['overdue', 'Overdue follow-ups', 'tasks'], ['investigations', 'Investigations', 'tasks'], ['urgent', 'Urgent requests', 'requests']
  ];
  return node('div', { class: 'counts' }, counts.map(([key, label, unit]) => {
    const total = data.requests.filter(r => {
      const p = data.patients.find(p => p.key === r.patientKey);
      if (!p || (state.scope === 'My work' && r.owner !== state.user)) return false;
      if (key === 'today') return isOpen(r) && r.due === now;
      if (key === 'overdue') return isOpen(r) && r.due && r.due < now;
      if (key === 'urgent') return isOpen(r) && r.urgent;
      return queueMatches(r, p, key, data.settings, now);
    }).length;
    const selected = key === state.tab || (state.tab === 'followups' && ((key === 'today' && state.priority === 'Today') || (key === 'overdue' && state.priority === 'Overdue'))) || (key === 'urgent' && state.tab === 'all' && state.priority === 'Urgent');
    return node('div', { class: 'metric' + (selected ? ' selected' : '') }, node('button', { type: 'button', class: 'metric-open', 'aria-pressed': selected, 'aria-label': t(label) + ': ' + total, on: { click: () => { if (!navigate(['today', 'overdue'].includes(key) ? 'followups' : key === 'urgent' ? 'all' : key)) return; if (key === 'urgent') state.priority = 'Urgent'; if (key === 'overdue') state.priority = 'Overdue'; if (key === 'today') state.priority = 'Today'; render(); } } }, node('span', { class: 'metric-label', text: t(label) }), node('span', { class: 'metric-number', text: total }), node('span', { class: 'metric-unit', text: t(unit) })), info(helpText[key] || helpText.followups, label));
  }));
}
function renderToolbar(sheet) {
  const toolbar = node('div', { class: 'toolbar' });
  if (state.editor) {
    toolbar.append(action('Save', saveEditor, { primary: true, help: helpText.save, id: 'save-record' }), action('Cancel', cancelEditor, { id: 'cancel-edit' }));
  } else {
    toolbar.append(action(state.tab === 'patients' ? 'Add patient' : 'Add request', () => beginEditor(state.tab === 'patients' ? 'new-patient' : 'new-request'), { primary: true, help: state.tab === 'patients' ? helpText.addPatient : helpText.addRequest, id: 'add-record' }), action('Edit selected', () => beginEditor(state.tab === 'patients' ? 'edit-patient' : 'edit-request'), { disabled: !state.selected, help: helpText.edit, id: 'edit-record' }));
    if (state.selected) toolbar.append(action('View history', showSelectedHistory));
  }
  const search = node('input', { type: 'search', 'aria-label': t(state.tab === 'patients' ? 'Search patient, ID or phone' : 'Search patient, ID or request'), placeholder: t(state.tab === 'patients' ? 'Search patient, ID or phone' : 'Search patient, ID or request'), value: state.search });
  search.addEventListener('input', () => { state.search = search.value; rerenderRows(); });
  toolbar.append(node('label', { class: 'search-box' }, node('span', { 'aria-hidden': 'true', text: '⌕' }), search), action('Show / hide columns', () => { state.columns = !state.columns; render(); }, { help: helpText.columns, id: 'column-toggle' }));
  if(state.selected){const req=data.requests.find(r=>r.key===state.selected);const p=data.patients.find(p=>p.key===(req?.patientKey||state.selected));if(p)toolbar.append(action('Record contact outcome',()=>contactForm(workspaceContext(),p,req),{help:'Preserve a structured contact result and the original notes.'}));}
  if(state.simpleFilters?.length)toolbar.append(action('Clear work filters',()=>{state.simpleFilters=[];render();}));
  toolbar.append(node('span', { class: 'spacer' }), select('Direction', ['Left to right', 'Right to left'], tablePrefs().direction === 'rtl' ? 'Right to left' : 'Left to right', value => { tablePrefs().direction = value === 'Right to left' ? 'rtl' : 'ltr'; remember(); render(); }), info(helpText.direction, 'Direction'));
  sheet.append(toolbar);
}
function renderFilters(sheet) {
  const bar = node('div', { class: 'filter-row' });
  if (state.tab === 'patients') bar.append(select('All patients', ['All', 'ID warning', 'Missing details'], state.priority, value => { state.priority = value; state.selected = null; render(); }));
  else {
    bar.append(select('Priority', ['All', 'Urgent', 'Overdue', 'Today'], state.priority, value => { state.priority = value; render(); }), select('Progress', ['All', 'Under verification', 'Ready to send', 'Awaiting approval', 'Approved', 'Completed', 'Closed'], state.progress, value => { state.progress = value; render(); }), select('Responsible staff', ['All staff',...staffNames(),''], state.owner, value => { state.owner = value; render(); }), select('Notification', ['All', 'Not informed', 'Awaiting confirmation', 'Informed', 'No answer'], state.notification, value => { state.notification = value; render(); }));
    if (state.tab === 'followups') bar.append(select('Follow-ups', ['Today + overdue', 'Upcoming', 'All'], state.followup, value => { state.followup = value; render(); }));
  }
  for (const [key, label, type] of [['ageMin', 'Age from', 'number'], ['ageMax', 'Age to', 'number'], ['arrivalFrom', 'Arrival from', 'date'], ['arrivalTo', 'Arrival to', 'date']]) bar.append(node('label', { class: 'small-label' }, t(label), node('input', { type, min: type === 'number' ? '0' : null, step: type === 'number' ? '1' : null, 'aria-label': t(label), value: state[key], on: { change: event => { state[key] = event.target.value; state.selected = null; render(); } } })));
  if ((state.ageMin !== '' && state.ageMax !== '' && Number(state.ageMin) > Number(state.ageMax)) || (state.arrivalFrom && state.arrivalTo && state.arrivalFrom > state.arrivalTo)) bar.append(node('span', { class: 'field-error', role: 'alert', text: t('The start of a range must not exceed its end.') }));
  bar.append(select('Hospital',['All',...new Set(data.requests.map(r=>r.hospital).filter(Boolean))],state.hospital||'All',value=>{state.hospital=value;render();}));
  bar.append(action('Clear filters', () => { state.simpleFilters=[];state.hospital='All';state.search = ''; state.ageMin = ''; state.ageMax = ''; state.arrivalFrom = ''; state.arrivalTo = ''; state.followup = 'All'; state.priority = 'All'; state.progress = 'All'; state.owner = 'All staff'; state.notification = 'All'; render(); }, { className: 'quiet' }), node('span', { class: 'selection-note', id: 'selection-note', text: state.selected ? '1 ' + t('selected') : t('Select one row to edit') }));
  sheet.append(bar);
}
function renderColumns(sheet) {
  const panel = node('section', { class: 'column-panel', 'aria-label': t('Visible columns') });
  panel.append(node('div', { class: 'panel-heading' }, t('Visible columns'), action('Show all', () => { tablePrefs().hidden = []; remember(); render(); })));
  panel.append(node('div', { class: 'column-grid' }, allFields().map(field => {
    const input = node('input', { type: 'checkbox', checked: !tablePrefs().hidden.includes(field.id), 'data-column': field.id, on: { change: event => { tablePrefs().hidden = event.target.checked ? tablePrefs().hidden.filter(id => id !== field.id) : [...new Set([...tablePrefs().hidden, field.id])]; remember(); render(); } } });
    return node('label', {}, input, t(field.label));
  })));
  sheet.append(panel);
}
function renderEditorBanner(sheet) {
  const editor = state.editor;
  const label = editor.mode === 'new-patient' ? 'New patient' : editor.mode === 'new-request' ? 'New request' : editor.mode === 'edit-patient' ? 'Editing patient' : 'Editing request';
  const banner = node('div', { class: 'edit-banner' + (state.conflict ? ' error' : '') }, node('strong', { text: t(label) }), node('span', { text: t(editor.mode === 'new-request' ? 'New request starts with patient ID lookup.' : 'Fill the horizontal fields, then save.') }));
  if (editor.mode.startsWith('edit')) banner.append(node('label', { class: 'small-label' }, t('Reason for correction / reopening'), node('input', { class: 'editor-control', 'aria-label': t('Reason for correction / reopening'), value: editor.reason || '', on: { input: event => { editor.reason = event.target.value; editor.dirty = true; } } })));
  if (editor.lookupNote) banner.append(node('span', { text: t(editor.lookupNote) }));
  if (state.conflict) banner.append(node('span', { text: t('This record changed in another tab. Review the differences before saving.') }), action('Review conflicts', showConflict));
  sheet.append(banner);
}
function visibleFields() {
  const fields = allFields();
  if (state.editor) return fields.filter(f => !tablePrefs().hidden.includes(f.id) || ['p:name', 'p:id', 'r:source', 'r:requested', 'r:approved'].includes(f.id) || state.editor.errors[f.id]);
  return fields.filter(field => !tablePrefs().hidden.includes(field.id));
}
function renderTable(sheet) {
  const fields = visibleFields(), rows = getRows();
  const wrapper = node('div', { class: 'table-scroll', dir: tablePrefs().direction });
  const table = node('table', { class: 'data-table', dir: tablePrefs().direction, 'aria-label': t(state.tab === 'patients' ? 'Patients' : 'All requests') });
  const heading = node('tr', {}, node('th', { class: 'selection', scope: 'col', text: t('Select') }));
  fields.forEach(field => heading.append(node('th', { scope: 'col', 'data-field': field.id, text: t(field.label) })));
  const body = node('tbody');
  if (state.editor) body.append(renderEditRow(fields));
  for (const row of rows) if (row.key !== state.editor?.recordKey) body.append(renderRow(row, fields));
  table.append(node('thead', {}, heading), body); wrapper.append(table); sheet.append(wrapper);
  if (!rows.length && !state.editor) sheet.append(node('div', { class: 'empty-state' }, node('h2', { text: t('No matching records') }), node('p', { text: t('Try changing your search or filters.') })));
  sheet.append(node('div', { class: 'table-footer' }, node('span', { text: rows.length + ' ' + t('rows') }), node('span', { text: fields.length + ' ' + t('of') + ' ' + allFields().length + ' ' + t('columns') }), node('span', { class: 'footer-right', text: t('Patient name') + ' / ' + t('Patient ID') + ' · ' + t('Direction') + ': ' + tablePrefs().direction.toUpperCase() })));
}
function rerenderRows() {
  const sheet = root.querySelector('.sheet');
  if (!sheet) return;
  sheet.querySelector('.table-scroll')?.remove(); sheet.querySelector('.empty-state')?.remove(); sheet.querySelector('.table-footer')?.remove(); renderTable(sheet); requestAnimationFrame(pinColumns);
}
function renderRow(row, fields) {
  const tr = node('tr', { class: row.key === state.selected ? 'selected' : '', 'data-record': row.key });
  const radio = node('input', { type: 'checkbox', checked: row.key === state.selected, 'aria-label': t('Select') + ' ' + row.p.name, on: { change: () => { state.selected = radio.checked ? row.key : null; render(); } } });
  tr.append(node('td', { class: 'selection' }, radio));
  for (const field of fields) {
    const td = node('td', { 'data-field': field.id, class: field.id === 'p:name' ? 'patient-name' : '' });
    if (field.id === 'p:id') {
      td.append(node('bdi', { text: row.p.id }));
      if (!validId(row.p.id)) td.append(node('span', { class: 'row-meta' }, node('span', { class: 'badge warning', text: t('ID needs verification') })));
    } else if (field.id === 'p:name') {
      td.append(node('span', { text: row.p.name }));
      const missing = missingDetails(row.p, data.settings.necessary);
      if (missing.length) td.append(node('span', { class: 'row-meta' }, node('span', { class: 'badge warning', text: t('Missing details') + ': ' + missing.map(key => t(patientFields.find(f => f.key === key)?.label || key)).join(', ') })));
      if (row.r) td.append(node('span', { class: 'row-meta', text: row.r.reference }));
    } else if(['phone','otherPhone'].includes(field.key) && row.p[field.key]) {td.append(node('bdi',{text:row.p[field.key]}),action('WhatsApp',()=>whatsappForm(workspaceContext(),row.p,row.p[field.key],row.r),{className:'whatsapp-button'}));
    } else if (field.key === 'urgent') td.append(node('span', { class: 'badge' + (row.r?.urgent ? ' urgent' : ''), text: row.r?.urgent ? t('Urgent') : '—' }));
    else if (field.key === 'requests') {
      const requests = patientRequests(row.p);
      if (!requests.length) td.textContent = '—';
      requests.forEach(r => td.append(node('span', { class: 'request-ref' }, node('button', { type: 'button', class: 'link', text: r.reference, on: { click: () => { navigate(r.progress === 'Completed' ? 'completed' : r.progress === 'Closed' ? 'closed' : 'all', r.reference); state.selected = r.key; render(); } } }))));
      td.append(node('span', { class: 'row-meta', text: requests.length + ' ' + t('requests') }));
    } else if (field.type === 'long') renderLongCell(td, field, row, String(valueFor(field, row) || t('Not recorded')));
    else if (field.key === 'progress' || field.key === 'informed' || field.key === 'outcome') td.append(node('span', { class: 'badge' + (row.r?.progress === 'Completed' && field.key === 'progress' ? ' success' : ''), text: t(String(valueFor(field, row) || 'Not recorded')) }));
    else {
      const value = valueFor(field, row);
      td.textContent = field.type === 'date' ? dateLabel(value) || t('Not recorded') : typeof value === 'boolean' ? (value ? '✓' : '—') : t(String(value || 'Not recorded'));
      if (field.key === 'due' && value && value < now) td.classList.add('date-overdue');
    }
    tr.append(td);
  }
  return tr;
}
function renderLongCell(td, field, row, text) {
  td.classList.add('text-cell');
  const key = row.key + ':' + field.id;
  const preview = node('div', { class: 'cell-preview' + (state.expanded.has(key) ? ' expanded' : ''), text, dir: field.key === 'ar' ? 'rtl' : field.key === 'en' ? 'ltr' : null });
  const actions = node('div', { class: 'cell-actions' });
  actions.append(node('button', { type: 'button', class: 'link', text: t(state.expanded.has(key) ? 'Show less' : 'Show full text'), 'aria-expanded': state.expanded.has(key), on: { click: event => { if (state.expanded.has(key)) state.expanded.delete(key); else state.expanded.add(key); preview.classList.toggle('expanded', state.expanded.has(key)); event.currentTarget.textContent = t(state.expanded.has(key) ? 'Show less' : 'Show full text'); event.currentTarget.setAttribute('aria-expanded', state.expanded.has(key)); } } }));
  if (['en', 'ar'].includes(field.key)) actions.append(node('button', { type: 'button', class: 'link', text: t('Copy'), on: { click: async () => { try { await navigator.clipboard.writeText(text); toast('Copied'); } catch { showDialog(field.label, node('div', { class: 'dialog-content-text', text })); } } } }));
  if (field.key === 'history') actions.append(node('button', { type: 'button', class: 'link', text: t('View history'), on: { click: () => showHistory(row.p, row.r) } }));
  td.append(preview, actions);
}
function pinColumns() {
  const table = root.querySelector('.data-table'); if (!table) return;
  table.querySelectorAll('td,th').forEach(cell => { cell.style.removeProperty('inset-inline-start'); if (cell.tagName === 'TD') cell.style.removeProperty('position'); cell.style.removeProperty('z-index'); });
  let offset = 0;
  const keys = window.innerWidth <= 620 ? ['p:name'] : ['p:name', 'p:id'];
  keys.forEach(key => {
    const head = table.querySelector(`th[data-field="${key}"]`); if (!head) return;
    table.querySelectorAll(`[data-field="${key}"]`).forEach(cell => { cell.style.position = 'sticky'; cell.style.insetInlineStart = offset + 'px'; cell.style.zIndex = cell.tagName === 'TH' ? '8' : '3'; }); offset += head.getBoundingClientRect().width;
  });
}
function emptyPatient() { return { key: crypto.randomUUID(), version: 0, name: '', id: '', originalId: '', phone: '', otherPhone: '', centre: '', building: '', apartment: '', gender: '', birth: '', reportedAge: '', ageRecorded: '', beneficiary: '', entryDate: '' }; }
function emptyRequest(patientKey) { return { key: crypto.randomUUID(), version: 0, reference: 'New request', patientKey, source: 'Clinic referral', requested: '', approved: '', received: '', progress: 'Under verification', hospital: '', appointment: '', informed: 'Not informed', next: '', due: '', owner: '', urgent: false, specialty: '', diagnosis: '', doctor: '', reportDate: '', reportUrl: '', documentUrl: '', responseDate: '', verified: false, outcome: 'Not confirmed', closureReason: '', contacts: [], createdDate: now, sentDate: '' }; }
function beginEditor(mode) {
  if (state.editor?.dirty && !confirm(t('Discard unsaved changes?'))) return;
  let p, r;
  if (mode === 'edit-patient') p = data.patients.find(p => p.key === state.selected);
  if (mode === 'edit-request') { r = data.requests.find(r => r.key === state.selected); p = data.patients.find(p => p.key === r?.patientKey); }
  if (mode.startsWith('edit') && !p) return;
  p ||= emptyPatient();
  if (mode === 'new-request') r = emptyRequest(p.key);
  state.editor = { mode, recordKey: mode.startsWith('edit') ? state.selected : null, p: structuredClone(p), r: r ? structuredClone(r) : null, baseP: mode.startsWith('edit') ? structuredClone(p) : null, baseR: mode === 'edit-request' ? structuredClone(r) : null, errors: {}, dirty: false, lookupNote: '', reason: '' };
  state.conflict = null;
  render();
  setTimeout(() => root.querySelector(`[data-input="${mode === 'new-request' ? 'p:id' : 'p:name'}"]`)?.focus(), 0);
}
function cancelEditor() { if (state.editor?.dirty && !confirm(t('Discard unsaved changes?'))) return; state.editor = null; state.conflict = null; render(); }
function lookupPatient() {
  const editor = state.editor; if (editor?.mode !== 'new-request') return;
  const raw = editor.p.id, id = normalizeId(raw); if (!id) return;
  const match = data.patients.find(p => normalizeId(p.id) === id);
  if (match) { editor.p = structuredClone(match); editor.baseP = structuredClone(match); editor.r.patientKey = match.key; editor.lookupNote = 'Existing patient found; personal details loaded.'; }
  else { if (editor.baseP) { editor.p = emptyPatient(); editor.p.id = raw; editor.r.patientKey = editor.p.key; } editor.baseP = null; editor.lookupNote = 'New ID; complete the patient name and available personal details.'; }
  render();
}
function renderEditRow(fields) {
  const editor = state.editor;
  const tr = node('tr', { class: 'editing', 'data-editor': 'true' }, node('td', { class: 'selection', text: '✎' }));
  for (const field of fields) {
    const td = node('td', { 'data-field': field.id, class: field.id === 'p:name' ? 'patient-name' : '' });
    const target = field.scope === 'p' ? editor.p : field.scope === 'r' ? editor.r : null;
    if (!target || field.type === 'computed' || field.scope === undefined) { td.classList.add('readonly'); td.textContent = field.id === 'p:age' ? String(ageAt(editor.p.birth, now) ?? '') : target?.[field.key] || '—'; tr.append(td); continue; }
    let input;
    if (['hospital', 'agency'].includes(field.type)) {
      const entries = data.settings[field.type === 'hospital' ? 'hospitals' : 'agencies'] || [];
      const options = [...new Set(['', ...entries.map(entry => entry.name), target[field.key] || ''])];
      input = node('select', { class: 'editor-control' }, options.map(value => { const entry = entries.find(e => e.name === value); return node('option', { value, text: value ? (userPrefs().language === 'ar' && entry?.arabic ? entry.arabic : value) : '—' }); })); input.value = target[field.key] || '';
    }
    else if (field.type === 'select') { input = node('select', { class: 'editor-control' }, (field.key === 'owner' && cloudMode ? ['', ...staffNames()] : field.options).map(value => node('option', { value, text: value ? t(value) : '—' }))); input.value = target[field.key] || ''; }
    else if (field.type === 'long') input = node('textarea', { class: 'editor-control long-input', rows: '2', value: target[field.key] || '' });
    else input = node('input', { class: 'editor-control', type: field.type === 'checkbox' ? 'checkbox' : field.type === 'date' ? 'date' : field.type === 'number' ? 'number' : 'text', value: field.type === 'checkbox' ? undefined : target[field.key] || '', checked: field.type === 'checkbox' ? !!target[field.key] : undefined, inputmode: field.key === 'id' ? 'numeric' : null });
    input.setAttribute('aria-label', t(field.label)); input.dataset.input = field.id;
    if (editor.errors[field.id]) { input.setAttribute('aria-invalid', 'true'); input.setAttribute('aria-describedby', 'error-' + field.id.replace(':', '-')); }
    input.addEventListener('input', () => { target[field.key] = field.type === 'checkbox' ? input.checked : input.value; editor.dirty = true; });
    if (field.key === 'id') input.addEventListener('blur', lookupPatient);
    if (field.key === 'source') input.addEventListener('change', () => { if (editor.mode === 'new-request' && editor.r.source === 'Direct ministry approval') editor.r.progress = 'Approved'; render(); });
    td.append(input);
    if (editor.errors[field.id]) td.append(node('span', { class: 'field-error', id: 'error-' + field.id.replace(':', '-'), text: t(editor.errors[field.id]) }));
    if (field.key === 'id' && editor.p.id && !validId(editor.p.id)) td.append(node('span', { class: 'field-error', text: t('ID needs verification') }));
    tr.append(td);
  }
  return tr;
}
let saving = false;
async function saveEditor() {
 if(saving)return;
 saving=true;
 try { await performSaveEditor(); }
 catch(error) {
  if(cloudMode && error.conflict) { try { data=await refreshData(); const editor=state.editor; if(editor){state.conflict={p:data.patients.find(p=>p.key===editor.p.key),r:editor.r?data.requests.find(r=>r.key===editor.r.key):null};render();showConflict();} }catch{} }
  toast(error.message || 'Save failed. Your draft is still open.',true);
 } finally {saving=false;}
}
async function performSaveEditor() {
  const editor = state.editor; if (!editor) return;
  const latest = await refreshData();
  // Refresh identity matching on Save as well as blur, so keyboard submission cannot create a duplicate profile.
  if (editor.mode === 'new-request') {
    const match = latest.patients.find(p => normalizeId(p.id) === normalizeId(editor.p.id));
    if (match && !editor.baseP) { editor.p = structuredClone(match); editor.baseP = structuredClone(match); editor.r.patientKey = match.key; }
  }
  editor.errors = {};
  for (const [key, message] of Object.entries(patientErrors(editor.p, latest.patients, editor.baseP?.key))) editor.errors['p:' + key] = message;
  if (editor.r) for (const [key, message] of Object.entries(requestErrors(editor.r))) editor.errors['r:' + key] = message;
  if (editor.baseP && normalizeId(editor.baseP.id) !== normalizeId(editor.p.id) && !clean(editor.reason)) editor.errors['p:id'] = 'identityReasonRequired';
  if (editor.baseR && ['Completed', 'Closed'].includes(editor.baseR.progress) && isOpen(editor.r)) {
    if (!clean(editor.reason)) editor.errors['r:progress'] = 'reopenReasonRequired';
    if (!clean(editor.r.next)) editor.errors['r:next'] = 'nextActionRequired';
    if (!editor.r.due) editor.errors['r:due'] = 'dueRequired';
  }
  if (Object.keys(editor.errors).length) { render(); setTimeout(() => root.querySelector('[aria-invalid=true]')?.focus(), 0); return; }
  const currentP = editor.baseP ? latest.patients.find(p => p.key === editor.baseP.key) : null;
  const currentR = editor.baseR ? latest.requests.find(r => r.key === editor.baseR.key) : null;
  if ((editor.baseP && currentP?.version !== editor.baseP.version) || (editor.baseR && currentR?.version !== editor.baseR.version)) { state.conflict = { p: currentP, r: currentR }; render(); showConflict(); return; }
  const timestamp = new Date().toISOString();
  const p = { ...editor.p, name: clean(editor.p.name), id: normalizeId(editor.p.id), originalId: currentP && editor.p.id === currentP.id ? currentP.originalId || currentP.id : editor.p.id, version: (currentP?.version || 0) + 1, createdAt: currentP?.createdAt || timestamp, updatedAt: timestamp };
  if (p.reportedAge && !p.ageRecorded) p.ageRecorded = now;
  if (editor.baseP) latest.patients = latest.patients.map(row => row.key === p.key ? p : row); else latest.patients.push(p);
  latest.audit.push(...changeEvents(currentP, p, 'patient', state.user, timestamp));
  if (editor.r) {
    const r = { ...editor.r, patientKey: p.key, version: (currentR?.version || 0) + 1, updatedAt: timestamp };
    if (!editor.baseR) { let number = latest.requests.reduce((m, row) => Math.max(m, Number(row.reference.replace(/\D/g, '')) || 0), 0) + 1; r.reference = 'REQ-' + String(number).padStart(3, '0'); }
    if (r.progress === 'Completed' && currentR?.progress !== 'Completed') { r.completedDate = now; r.due = ''; r.next = ''; }
    if (r.progress === 'Closed' && currentR?.progress !== 'Closed') { r.closedDate = now; r.due = ''; r.next = ''; }
    if (editor.baseR) latest.requests = latest.requests.map(row => row.key === r.key ? r : row); else latest.requests.push(r);
    latest.audit.push(...changeEvents(currentR, r, 'request', state.user, timestamp)); state.selected = r.key;
  } else state.selected = p.key;
  if (clean(editor.reason)) latest.audit.push({ key: crypto.randomUUID(), recordKey: editor.r?.key || p.key, recordType: editor.r ? 'request' : 'patient', field: 'changeReason', before: '', after: clean(editor.reason), actor: state.user, at: timestamp });
  latest.revision++;
  data = await persist(latest, {reason:editor.reason}); state.editor = null; state.conflict = null; render(); toast('Saved locally');
}
function showConflict() {
  const editor = state.editor, conflict = state.conflict; if (!editor || !conflict) return;
  const table = node('table', { class: 'history-table' }, node('thead', {}, node('tr', {}, ['Field', 'Current saved value', 'Your draft'].map(label => node('th', { text: t(label) })))));
  const body = node('tbody');
  for (const [scope, current, draft] of [['p', conflict.p, editor.p], ['r', conflict.r, editor.r]]) if (draft) for (const field of [...patientFields, ...requestFields].filter(f => f.scope === scope)) if (String(current?.[field.key] ?? '') !== String(draft[field.key] ?? '')) body.append(node('tr', {}, node('td', { text: t(field.label) }), node('td', { text: String(current?.[field.key] ?? '') }), node('td', { text: String(draft[field.key] ?? '') })));
  table.append(body);
  showDialog('Review conflicts', node('div', { class: 'dialog-scroll' }, table), [action('Keep my draft for review', () => { editor.baseP = structuredClone(conflict.p); if (editor.r) editor.baseR = conflict.r ? structuredClone(conflict.r) : null; state.conflict = null; dialog.close(); render(); }), action('Reload latest', () => { if (!confirm(t('Discard your draft and load the latest saved version?'))) return; if (conflict.p) { editor.p = structuredClone(conflict.p); editor.baseP = structuredClone(conflict.p); } if (conflict.r) { editor.r = structuredClone(conflict.r); editor.baseR = structuredClone(conflict.r); } editor.dirty = false; state.conflict = null; dialog.close(); data = readData(); render(); })]);
}
function showSelectedHistory() {
  const r = data.requests.find(r => r.key === state.selected);
  const p = data.patients.find(p => p.key === (r ? r.patientKey : state.selected)); if (p) showHistory(p, r);
}
function showHistory(p, r) {
  const events = historyFor(p, r).slice().reverse();
  const table = node('table', { class: 'history-table' }, node('thead', {}, node('tr', {}, ['When', 'Who', 'Record', 'Field', 'Before', 'After'].map(label => node('th', { text: t(label) })))));
  table.append(node('tbody', {}, events.map(e => node('tr', {}, [new Date(e.at).toLocaleString('en-GB', { timeZone: 'Africa/Cairo' }), t(e.actor), e.recordType, t([...patientFields, ...requestFields].find(f => f.key === e.field)?.label || e.field), String(e.before), String(e.after)].map(value => node('td', { text: value }))))));
  showDialog('Change history', events.length ? node('div', { class: 'dialog-scroll' }, table) : node('p', { text: t('No changes yet') }));
}
function renderSettings(page) {
  const sections = ['Accounts', 'Follow-ups & retries', 'Message templates', 'Custom columns', 'Hospitals', 'Backups'];
  const layout = node('div', { class: 'settings-layout' }), nav = node('nav', { class: 'settings-nav' });
  sections.forEach(label => nav.append(node('button', { type: 'button', text: t(label), class: state.settingsTab === label ? 'active' : '', on: { click: () => { state.settingsTab = label; render(); } } })));
  const panel = node('section', { class: 'settings-section' }, node('h2', { text: t(state.settingsTab) }), node('p', { text: t(cloudMode ? 'Workspace settings are saved in Supabase. External file-storage and automated backups are not connected.' : 'Production setup remains pending. No external accounts or services are connected.') }));
  if (state.settingsTab === 'Accounts') {
    panel.append(node('div', { class: 'note-box', text: t('Signed in as') + ': ' + t(state.user) + '. ' + t('Live workspace') }), node('h2', { text: t('Necessary patient details') }), node('p', { text: t('These settings control patient completeness warnings. Name and nonblank ID are always required for creating a profile.') }));
    panel.append(node('div', { class: 'settings-fields' }, patientFields.filter(f => !['name', 'id', 'age'].includes(f.key)).map(field => node('label', {}, node('input', { type: 'checkbox', checked: data.settings.necessary.includes(field.key), disabled: !isAdmin(), on: { change: async event => { const latest = await refreshData(); const before = structuredClone(latest.settings.necessary); latest.settings.necessary = event.target.checked ? [...new Set([...before, field.key])] : before.filter(key => key !== field.key); latest.audit.push({ key: crypto.randomUUID(), recordKey: 'settings', recordType: 'settings', field: 'necessary', before, after: latest.settings.necessary, actor: state.user, at: new Date().toISOString() }); latest.revision++; try { data = await persist(latest); toast('Saved locally'); } catch { toast('Save failed', true); render(); } } } }), t(field.label)))));
  } else if (state.settingsTab === 'Hospitals') { renderDirectory(panel, 'hospitals');
  } else if (state.settingsTab === 'Backups') {
    if(cloudMode) panel.append(action('Download backup',downloadBackup),node('p',{text:'Manual backups include patient records, requests, source rows and edit history. Store them in your restricted SharePoint or Google Drive folder.'}),action('Restore backup into empty workspace',showRestore,{disabled:!isAdmin()}));
    panel.append(node('div', { class: 'settings-fields' }, select('Provider', ['Microsoft SharePoint', 'Google Drive'], userPrefs().backupProvider || 'Microsoft SharePoint', value => { userPrefs().backupProvider = value; remember(); })), node('div', { class: 'note-box' }, t('Not connected'), ' · Supabase Free'), node('p', { text: t('30 days, editable in production settings') }), node('p', { text: t('Change history is independent of backup retention.') }));
  } else panel.append(node('div', { class: 'note-box', text: t('Not configured') + ' · ' + t('Live workspace') }));
  layout.append(nav, panel); page.append(layout);
}
function renderDirectory(page, kind) {
  const label = kind === 'hospitals' ? 'Hospitals' : 'External referral agencies';
  page.append(node('p', { class: 'caption', text: t('Administrators maintain this directory. Existing request values are preserved.') }));
  const table = node('table', { class: 'history-table' }, node('thead', {}, node('tr', {}, ['English name', 'Arabic name', 'Requests', 'Actions'].map(text => node('th', { text: t(text) })))));
  const body = node('tbody');
  for (const entry of data.settings[kind] || []) body.append(node('tr', {}, node('td', { text: entry.name }), node('td', { text: entry.arabic || '—', dir: 'rtl' }), node('td', { text: data.requests.filter(r => r[kind === 'hospitals' ? 'hospital' : 'agency'] === entry.name).length }), node('td', {}, action('Edit', () => editDirectory(kind, entry), { disabled: !isAdmin() }))));
  table.append(body); page.append(table, action('Add entry', () => editDirectory(kind), { primary: true, disabled: !isAdmin() }));
}
function editDirectory(kind, existing = null) {
  const english = node('input', { value: existing?.name || '', 'aria-label': t('English name'), required: true });
  const arabicName = node('input', { value: existing?.arabic || '', 'aria-label': t('Arabic name'), dir: 'rtl' });
  const error = node('p', { class: 'field-error', role: 'alert' });
  showDialog(kind === 'hospitals' ? 'Hospitals' : 'External referral agencies', node('div', { class: 'settings-fields' }, node('label', {}, t('English name'), english), node('label', {}, t('Arabic name'), arabicName), error), [action('Save', async () => {
    if (!isAdmin()) return;
    const latest = await refreshData(), entries = latest.settings[kind], name = clean(english.value);
    if (!name || entries.some(e => e.key !== existing?.key && searchText(e.name) === searchText(name))) { error.textContent = t('Enter a unique name.'); return; }
    const entry = { key: existing?.key || crypto.randomUUID(), name, arabic: clean(arabicName.value) };
    latest.settings[kind] = existing ? entries.map(e => e.key === existing.key ? entry : e) : [...entries, entry];
    // A directory rename retains all existing associations and is recorded in history.
    if (existing && existing.name !== name) for (const r of latest.requests) { const field = kind === 'hospitals' ? 'hospital' : 'agency'; if (r[field] === existing.name) { const before = structuredClone(r); r[field] = name; r.version++; latest.audit.push(...changeEvents(before, r, 'request', state.user)); } }
    latest.audit.push({ key: crypto.randomUUID(), recordKey: entry.key, recordType: kind, field: 'directory', before: existing || '', after: entry, actor: state.user, at: new Date().toISOString() }); latest.revision++;
    try { data = await persist(latest); dialog.close(); render(); toast('Saved locally'); } catch { error.textContent = t('Save failed'); }
  }, { primary: true })]);
}
function renderReports(page) {
  const controls = node('div', { class: 'toolbar' });
  for (const [key, label] of [['reportFrom', 'From'], ['reportTo', 'To']]) controls.append(node('label', { class: 'small-label' }, t(label), node('input', { type: 'date', value: state[key], on: { change: event => { state[key] = event.target.value; render(); } } })));
  controls.append(action('This month', () => { state.reportFrom = now.slice(0, 7) + '-01'; state.reportTo = now; render(); })); page.append(controls);
  const inPeriod = value => value && value.slice(0, 10) >= state.reportFrom && value.slice(0, 10) <= state.reportTo;
  const items = [
    ['New patient profiles', data.patients.filter(p => inPeriod(p.createdAt)).length],
    ['New requests', data.requests.filter(r => inPeriod(r.createdDate)).length],
    ['Approvals recorded', data.requests.filter(r => r.approved && inPeriod(r.responseDate)).length],
    ['Appointments in period', data.requests.filter(r => inPeriod(r.appointment)).length],
    ['Contact attempts', data.requests.flatMap(r => r.contacts).filter(c => inPeriod(c.date)).length],
    ['Completed requests', data.requests.filter(r => r.progress === 'Completed' && inPeriod(r.completedDate)).length],
    ['Closed requests', data.requests.filter(r => r.progress === 'Closed' && inPeriod(r.closedDate)).length],
    ['Current open requests', data.requests.filter(isOpen).length]
  ];
  page.append(node('div', { class: 'note-box', text: t('Live workspace') + ' · ' + t('Current open requests') + ': ' + t('Today') }), node('div', { class: 'report-grid' }, items.map(([label, value]) => node('div', { class: 'report-value' }, node('span', { text: t(label) }), node('strong', { text: value })))));
}
function hideHelp() { clearTimeout(tooltipTimer); tooltip.hidden = true; }
function displayHelp(target) {
  if (!userPrefs().help) return;
  tooltip.textContent = t(target.dataset.help);
  tooltip.hidden = false;
  const box = target.getBoundingClientRect(), tip = tooltip.getBoundingClientRect();
  tooltip.style.left = Math.max(8, Math.min(innerWidth - tip.width - 8, box.left + box.width / 2 - tip.width / 2)) + 'px';
  tooltip.style.top = (box.bottom + tip.height + 12 < innerHeight ? box.bottom + 7 : Math.max(8, box.top - tip.height - 7)) + 'px';
  target.setAttribute('aria-describedby', 'tooltip');
}
root.addEventListener('pointerover', event => { const target = event.target.closest('[data-help]'); if (target) { clearTimeout(tooltipTimer); tooltipTimer = setTimeout(() => displayHelp(target), 220); } });
root.addEventListener('pointerout', event => { if (event.target.closest('[data-help]') && !event.relatedTarget?.closest?.('[data-help]')) hideHelp(); });
root.addEventListener('focusin', event => { const target = event.target.closest('[data-help]'); if (target) displayHelp(target); });
root.addEventListener('focusout', hideHelp);
root.addEventListener('click', event => { const target = event.target.closest('[data-help]'); if (target) displayHelp(target); else hideHelp(); });
document.addEventListener('keydown', event => { if (event.key === 'Escape') hideHelp(); if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's' && state.editor) { event.preventDefault(); saveEditor(); } });
window.addEventListener('resize', () => { hideHelp(); pinColumns(); });
window.addEventListener('beforeunload', event => { if (state.editor?.dirty) { event.preventDefault(); event.returnValue = ''; } });
function workspaceContext(){return {node,action,t,data,state,now,navigate,actor:state.user,language:userPrefs().language,patientFields,requestFields,staffNames,getRows,render,renderFilters,beginEditor,saveEditor,cancelEditor,lookup:lookupPatient,history:showHistory,showDialog,closeDialog:()=>dialog.close(),toast,refresh:refreshData,save:async(next,options)=>{data=await persist(next,options);return data;},panel:(title,content,actions)=>{showDialog(title,content,actions);dialog.classList.add('profile-panel');},contact:(p,r)=>contactForm(workspaceContext(),p,r),whatsapp:(p,number,r)=>whatsappForm(workspaceContext(),p,number,r),edit:(p,r)=>{state.tab=r?(r.progress==='Completed'?'completed':r.progress==='Closed'?'closed':'all'):'patients';state.selected=r?.key||p.key;beginEditor(r?'edit-request':'edit-patient');},addForPatient:p=>{state.tab='all';beginEditor('new-request');state.editor.p=structuredClone(p);state.editor.baseP=structuredClone(p);state.editor.r.patientKey=p.key;render();},setFilters:filters=>{state.tab='all';state.simpleFilters=filters;state.priority='All';state.progress='All';state.notification='All';state.followup='All';render();}};}
async function start() {
 try { await initialiseCloud(); data=readData(); render(); }
 catch(error){renderSetup(error.message);}
}
function renderSetup(message='') {
 document.documentElement.lang='en';document.documentElement.dir='ltr';root.classList.remove('simple-mode');
 root.replaceChildren(node('main',{class:'login-page'},node('section',{class:'login-form'},node('img',{class:'erc-logo',src:'./assets/erc-logo.webp',alt:'Egyptian Red Crescent'}),node('h1',{text:'Workspace setup required'}),node('p',{text:'The application starts empty. Your administrator must configure its database before staff can sign in.'}),node('ol',{},node('li',{text:'Install the supplied database script in your Supabase project.'}),node('li',{text:'Set the project URL and publishable key in config.json.'}),node('li',{text:'Create your staff account, then reload this page.'})),message?node('p',{class:'field-error',role:'alert',text:message}):null,node('a',{href:'./setup.html',text:'Setup instructions',class:'setup-link'}),action('Check setup',()=>location.reload()))));
}
function renderLogin(message='') {
 document.documentElement.lang='en';document.documentElement.dir='ltr';root.classList.remove('simple-mode');
 const email=node('input',{type:'email',required:true,autocomplete:'username','aria-label':'Email'}),password=node('input',{type:'password',required:true,autocomplete:'current-password','aria-label':'Password'}),error=node('p',{class:'field-error',role:'alert',text:message});
 const submit=node('button',{type:'submit',class:'primary',text:'Sign in'});
 const form=node('form',{class:'login-form',on:{submit:async event=>{event.preventDefault();submit.disabled=true;error.textContent='Signing in…';try{data=await signIn(email.value,password.value);state.user=member.name;state.tab='overview';state.editor=null;render();}catch(failure){error.textContent=failure.message;}finally{password.value='';submit.disabled=false;}}}},node('img',{class:'erc-logo',src:'./assets/erc-logo.webp',alt:'Egyptian Red Crescent'}),node('h1',{text:'ERC External Referral System'}),node('p',{text:'Staff sign-in'}),node('label',{},'Email',email),node('label',{},'Password',password),submit,error,node('a',{href:'./setup.html',text:'Administrator setup instructions',class:'setup-link'}));
 root.replaceChildren(node('main',{class:'login-page'},form));
}
async function downloadBackup() {
 if(cloudMode) data=await refreshData();
 const blob=new Blob([JSON.stringify(data,null,2)],{type:'application/json'}),url=URL.createObjectURL(blob),link=node('a',{href:url,download:'erc-backup-'+todayCairo()+'.json'});link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);toast('Backup downloaded');
}
function showRestore() {
 const file=node('input',{type:'file',accept:'.json','aria-label':'ERC backup file'}),error=node('p',{class:'field-error',role:'alert'});
 showDialog('Restore backup into empty workspace',node('div',{class:'connection-form'},node('p',{text:'Restore is allowed only in a fresh, empty database. It preserves the saved records and historical edits. It cannot replace an existing working database. Recreate staff accounts separately.'}),file,error),[action('Restore selected backup',async()=>{
  try{if(!file.files[0])throw Error('Choose your ERC JSON backup.');if(!confirm('Restore this backup into the empty workspace?'))return;const backup=JSON.parse(await file.files[0].text());data=await restoreCloud(backup);dialog.close();render();toast('Backup restored');}catch(failure){error.textContent=failure.message;}
 },{primary:true,disabled:!isAdmin()})]);
}
window.addEventListener('online',()=>{if(cloudMode&&member)toast('Connection restored. Refresh records before continuing.');});
window.addEventListener('offline',()=>{if(cloudMode)toast('Connection lost. Wait for reconnection; your open draft remains on this page.',true);});
start();
