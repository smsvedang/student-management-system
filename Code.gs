/**
 * STUDENT MANAGEMENT SYSTEM — Google Apps Script backend
 * Spreadsheet: 1ZnzqvJqoWTOzJ6mD25yl_YO_UzUcmP9mU6xUOa5oDlM
 *
 * Setup:
 * 1. Paste into the spreadsheet-bound Apps Script project.
 * 2. In Project Settings > Script Properties add ADMIN_PASSWORD with a NEW strong password.
 * 3. Run setup() once and authorize.
 * 4. Deploy as Web App: Execute as Me; access Anyone.
 *
 * Public reads are intentionally limited to published form definitions.
 * All admin operations require an expiring server-side session token.
 */
const CFG = {
  SPREADSHEET_ID: '1ZnzqvJqoWTOzJ6mD25yl_YO_UzUcmP9mU6xUOa5oDlM',
  HEADER_SCAN_ROWS: 10,
  SESSION_SECONDS: 21600,
  STUDENTS_TAB: 'Students',
  FORMS_TAB: 'Form_Config',
  RESPONSES_TAB: 'Form_Responses',
  LOGS_TAB: 'Admin_Logs'
};

function setup() {
  const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  ensureSystemTab_(ss, CFG.FORMS_TAB, ['Form ID','Form Title','Description','Fields JSON','Published','Created At','Updated At']);
  ensureSystemTab_(ss, CFG.RESPONSES_TAB, ['Timestamp','Form ID','Form Title','Student Row','Enrollment Number','Submission JSON']);
  ensureSystemTab_(ss, CFG.LOGS_TAB, ['Timestamp','Action','Details','Actor']);

  // If the spreadsheet has no student-data tab yet, create an empty one.
  // Never replace, rename, or clear an existing non-system sheet.
  const systemNames = [CFG.FORMS_TAB, CFG.RESPONSES_TAB, CFG.LOGS_TAB];
  const hasStudentTab = ss.getSheets().some(s => systemNames.indexOf(s.getName()) < 0);
  if (!hasStudentTab) {
    const students = ss.insertSheet(CFG.STUDENTS_TAB);
    students.getRange(1,1,1,8).setValues([[
      'Name of Student','Mobile Number','Enrollment Number','Email ID of Student',
      'Name of Father','Mobile Number of Father','Email ID of Father','Name of Mentor'
    ]]);
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', CFG.SPREADSHEET_ID);
  if (!props.getProperty('ADMIN_PASSWORD')) {
    throw new Error('Set ADMIN_PASSWORD in Script Properties, then run setup() again.');
  }
  return 'Setup complete. Main student sheet was not modified.';
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || 'ping');
  let result;
  try {
    if (action === 'ping') result = {ok:true, message:'Student Management API ready'};
    else if (action === 'postResult') {
      const key = 'sms_result_' + String(p.nonce || '');
      const cached = p.nonce ? CacheService.getScriptCache().get(key) : null;
      result = cached ? JSON.parse(cached) : {ok:false, pending:true};
      if (cached) CacheService.getScriptCache().remove(key);
    }
    else if (action === 'publicForm') result = publicForm_(p.formId);
    else if (action === 'lookupStudent') result = lookupStudent_(p.enrollment);
    else if (action === 'admin') {
      requireSession_(p.token);
      result = adminAction_(p);
    } else throw new Error('Unknown action');
  } catch (err) {
    result = {ok:false, error:err.message || String(err)};
  }
  return output_(result, p.callback);
}

function doPost(e) {
  let result;
  try {
    const p = parsePost_(e);
    const action = String(p.action || '');
    if (action === 'login') result = login_(p.password);
    else if (action === 'submitForm') result = submitForm_(p.formId, p.data || {});
    else {
      requireSession_(p.token);
      result = adminAction_(p);
    }
  } catch (err) {
    result = {ok:false, error:err.message || String(err)};
  }
  const p = parsePost_(e);
  const nonce = String((e && e.parameter && e.parameter.nonce) || p.nonce || '');
  if (nonce) CacheService.getScriptCache().put('sms_result_' + nonce, JSON.stringify(result), 120);
  return HtmlService.createHtmlOutput('<!doctype html><html><body>Request received.</body></html>').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function parsePost_(e) {
  const raw = e && e.postData && e.postData.contents;
  if (!raw) return {};
  if ((e.postData.type || '').indexOf('application/json') >= 0) return JSON.parse(raw);
  const p = Object.assign({}, e.parameter || {});
  if (p.payload) return JSON.parse(p.payload);
  if (p.data && typeof p.data === 'string') { try { p.data = JSON.parse(p.data); } catch (_) {} }
  return p;
}

function output_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback && /^[a-zA-Z_$][0-9a-zA-Z_$\.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback+'('+json+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function login_(password) {
  const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!expected) throw new Error('Admin password is not configured.');
  if (String(password || '') !== expected) throw new Error('Incorrect password.');
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g,'');
  CacheService.getScriptCache().put('sms_session_'+token, 'admin', CFG.SESSION_SECONDS);
  log_('LOGIN','Admin login successful','admin');
  return {ok:true, token:token, expiresIn:CFG.SESSION_SECONDS};
}

function requireSession_(token) {
  if (!token || CacheService.getScriptCache().get('sms_session_'+token) !== 'admin') {
    throw new Error('Session expired. Please log in again.');
  }
}

function adminAction_(p) {
  const action = String(p.task || p.action || '');
  switch (action) {
    case 'dashboard': return dashboard_();
    case 'listStudents': return listStudents_(p.q || '');
    case 'saveStudent': return saveStudent_(p.record || {});
    case 'deleteStudent': return deleteStudent_(Number(p.row));
    case 'listForms': return listForms_();
    case 'formMeta': return formMeta_();
    case 'saveForm': return saveForm_(p.form || {});
    case 'deleteForm': return deleteForm_(p.formId);
    case 'responses': return getResponses_();
    case 'export': return exportCsv_();
    case 'logout': CacheService.getScriptCache().remove('sms_session_'+p.token); return {ok:true};
    default: throw new Error('Unknown admin action: '+action);
  }
}

function getMainSheet_() {
  const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  const skip = [CFG.FORMS_TAB,CFG.RESPONSES_TAB,CFG.LOGS_TAB];
  const sheet = ss.getSheets().find(s => skip.indexOf(s.getName()) < 0);
  if (!sheet) throw new Error('No student-data tab found. Run setup() once in the spreadsheet-bound Apps Script project.');
  return sheet;
}

function headerInfo_(sheet) {
  const rows = Math.min(CFG.HEADER_SCAN_ROWS, Math.max(sheet.getLastRow(),1));
  const cols = Math.max(sheet.getLastColumn(),1);
  const values = sheet.getRange(1,1,rows,cols).getDisplayValues();
  let best = {row:1,score:0,headers:[]};
  values.forEach((r,i) => {
    const nonempty = r.filter(v => String(v).trim()).length;
    const words = r.filter(v => /student|enrollment|mobile|email|father|mother|mentor|name|program|institute/i.test(v)).length;
    const score = words*3 + nonempty;
    if (score > best.score) best = {row:i+1,score,headers:r};
  });
  const headers = best.headers.map((h,i)=>({name:String(h||'').trim(),col:i+1})).filter(x=>x.name);
  if (!headers.length) throw new Error('Could not detect header row in main sheet.');
  return {row:best.row,headers:headers};
}

function ensureSystemTab_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  return sh;
}

function dashboard_() {
  const sh = getMainSheet_(), hi = headerInfo_(sh);
  const dataRows = Math.max(0, sh.getLastRow()-hi.row);
  const headers = hi.headers.map(x=>x.name);
  const values = dataRows ? sh.getRange(hi.row+1,1,dataRows,sh.getLastColumn()).getDisplayValues() : [];
  const filled = values.filter(r=>r.some(v=>String(v).trim())).length;
  let missing=0;
  const nameCol = findCol_(hi,'student name|name of student|name');
  const enrollCol = findCol_(hi,'enrollment number|enrollment no');
  values.forEach(r=>{if ((nameCol && !r[nameCol-1]) || (enrollCol && !r[enrollCol-1])) missing++;});
  return {ok:true,total:filled,missing:missing,headerRow:hi.row,columns:headers.length,forms:listForms_().forms.length};
}

function listStudents_(query) {
  const sh=getMainSheet_(), hi=headerInfo_(sh), last=sh.getLastRow();
  const values=last>hi.row?sh.getRange(hi.row+1,1,last-hi.row,sh.getLastColumn()).getDisplayValues():[];
  const q=String(query||'').toLowerCase().trim();
  const records=[];
  values.forEach((r,i)=>{
    if (!r.some(v=>String(v).trim())) return;
    const obj={_row:hi.row+1+i};
    hi.headers.forEach(h=>obj[h.name]=r[h.col-1]||'');
    if (!q || Object.keys(obj).some(k=>String(obj[k]).toLowerCase().includes(q))) records.push(obj);
  });
  return {ok:true,headers:hi.headers.map(h=>h.name),records:records};
}

function saveStudent_(record) {
  const sh=getMainSheet_(), hi=headerInfo_(sh);
  const row=Number(record._row||0);
  const existing = row>hi.row && row<=sh.getLastRow();
  const target = existing ? row : Math.max(sh.getLastRow()+1,hi.row+1);
  const width=sh.getLastColumn();
  const old=existing?sh.getRange(target,1,1,width).getValues()[0]:new Array(width).fill('');
  const normalized={};
  Object.keys(record).forEach(k=>normalized[norm_(k)]=record[k]);
  hi.headers.forEach(h=>{
    const key=norm_(h.name);
    if (Object.prototype.hasOwnProperty.call(normalized,key)) old[h.col-1]=normalized[key];
  });
  const enroll=findCol_(hi,'enrollment number|enrollment no');
  if (enroll) {
    const val=String(old[enroll-1]||'').trim();
    if (val) {
      const data=sh.getRange(hi.row+1,enroll,Math.max(0,sh.getLastRow()-hi.row),1).getDisplayValues().flat();
      if (data.some((v,i)=>v.trim()===val && hi.row+1+i!==target)) throw new Error('Enrollment number already exists.');
    }
  }
  if (!existing && target>sh.getLastRow()+1) sh.insertRowsAfter(sh.getLastRow(),target-sh.getLastRow()-1);
  sh.getRange(target,1,1,width).setValues([old]);
  log_(existing?'UPDATE_STUDENT':'ADD_STUDENT','Row '+target,'admin');
  return {ok:true,row:target};
}

function deleteStudent_(row) {
  const sh=getMainSheet_(), hi=headerInfo_(sh);
  if (!row || row<=hi.row || row>sh.getLastRow()) throw new Error('Invalid student row.');
  sh.deleteRow(row); log_('DELETE_STUDENT','Row '+row,'admin'); return {ok:true};
}

function formMeta_() {
  const hi = headerInfo_(getMainSheet_());
  return {ok:true, columns:hi.headers.map(h=>h.name), identityColumn:findCol_(hi,'enrollment number|enrollment no') ? hi.headers.find(h=>h.col===findCol_(hi,'enrollment number|enrollment no')).name : ''};
}

function lookupStudent_(enrollment) {
  const key = String(enrollment || '').trim().toLowerCase();
  if (!key) throw new Error('Enter your enrollment number first.');
  const sh = getMainSheet_(), hi = headerInfo_(sh);
  const col = findCol_(hi,'enrollment number|enrollment no');
  if (!col) throw new Error('Enrollment Number column is missing in the student sheet.');
  const last = sh.getLastRow();
  const rows = last > hi.row ? sh.getRange(hi.row+1,1,last-hi.row,sh.getLastColumn()).getDisplayValues() : [];
  const idx = rows.findIndex(r => String(r[col-1] || '').trim().toLowerCase() === key);
  if (idx < 0) throw new Error('Enrollment number not found. Please check it or contact your administrator.');
  const record = {};
  hi.headers.forEach(h => record[h.name] = rows[idx][h.col-1] || '');
  return {ok:true,student:record};
}

function listForms_() {
  const sh=SpreadsheetApp.openById(CFG.SPREADSHEET_ID).getSheetByName(CFG.FORMS_TAB);
  const rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,7).getValues():[];
  return {ok:true,forms:rows.filter(r=>r[0]).map(r=>({id:r[0],title:r[1],description:r[2],fields:JSON.parse(r[3]||'[]'),published:String(r[4])==='true'}))};
}

function publicForm_(id) {
  if (!id) throw new Error('Missing form ID.');
  const sh=SpreadsheetApp.openById(CFG.SPREADSHEET_ID).getSheetByName(CFG.FORMS_TAB);
  const rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,7).getValues():[];
  const r=rows.find(x=>String(x[0])===String(id) && String(x[4])==='true');
  if (!r) throw new Error('This form is unavailable.');
  return {ok:true,form:{id:r[0],title:r[1],description:r[2],fields:JSON.parse(r[3]||'[]')}};
}

function saveForm_(form) {
  const title=String(form.title||'').trim();
  if (!title) throw new Error('Form title is required.');
  const fields=Array.isArray(form.fields)?form.fields:[];
  if (!fields.length) throw new Error('Add at least one field.');
  const sh=SpreadsheetApp.openById(CFG.SPREADSHEET_ID).getSheetByName(CFG.FORMS_TAB);
  const id=form.id||Utilities.getUuid().slice(0,8);
  const main=getMainSheet_(), mainHi=headerInfo_(main);
  fields.forEach(field=>{ if(!mainHi.headers.some(h=>norm_(h.name)===norm_(field.label))) { const added=ensureColumn_(main,mainHi,field.label); mainHi.headers.push(added); } });
  const rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,7).getValues():[];
  const idx=rows.findIndex(r=>String(r[0])===String(id));
  const now=new Date();
  const row=[id,title,String(form.description||''),JSON.stringify(fields),form.published?'true':'false',idx>=0?rows[idx][5]:now,now];
  if(idx>=0) sh.getRange(idx+2,1,1,7).setValues([row]); else sh.appendRow(row);
  log_('SAVE_FORM',title,'admin');
  return {ok:true,id:id};
}

function deleteForm_(id) {
  const sh=SpreadsheetApp.openById(CFG.SPREADSHEET_ID).getSheetByName(CFG.FORMS_TAB);
  const rows=sh.getLastRow()>1?sh.getRange(2,1,sh.getLastRow()-1,7).getValues():[];
  const idx=rows.findIndex(r=>String(r[0])===String(id));
  if(idx<0) throw new Error('Form not found.');
  sh.deleteRow(idx+2); log_('DELETE_FORM',String(id),'admin'); return {ok:true};
}

function submitForm_(id,data) {
  const pub=publicForm_(id).form;
  const sh=getMainSheet_(), hi=headerInfo_(sh);
  const values={};
  pub.fields.forEach(f=>{
    const v=data[f.key];
    if(f.required && (v===undefined || String(v).trim()==='')) throw new Error(f.label+' is required.');
    values[norm_(f.label)]=v||'';
  });
  const width=sh.getLastColumn();
  const row=new Array(width).fill('');
  let target=sh.getLastRow()+1;
  const enroll=findCol_(hi,'enrollment number|enrollment no');
  const submittedEnroll=values[norm_('Enrollment Number')]||'';
  if(enroll && submittedEnroll) {
    const existing=sh.getRange(hi.row+1,enroll,Math.max(0,sh.getLastRow()-hi.row),1).getDisplayValues().flat();
    if(existing.some(v=>v.trim()===String(submittedEnroll).trim())) throw new Error('This enrollment number is already registered.');
  }
  pub.fields.forEach(f=>{
    let col=hi.headers.find(h=>norm_(h.name)===norm_(f.label));
    if(!col) col=ensureColumn_(sh,hi,f.label);
    row[col.col-1]=values[norm_(f.label)]||'';
  });
  sh.getRange(target,1,1,sh.getLastColumn()).setValues([row.concat(new Array(Math.max(0,sh.getLastColumn()-row.length)).fill(''))]);
  SpreadsheetApp.openById(CFG.SPREADSHEET_ID).getSheetByName(CFG.RESPONSES_TAB).appendRow([new Date(),id,pub.title,target,submittedEnroll,JSON.stringify(data)]);
  log_('PUBLIC_SUBMISSION',pub.title,'public');
  return {ok:true,message:'Your information has been submitted.'};
}

function ensureColumn_(sh,hi,label) {
  const last=sh.getLastColumn();
  // Add a blank spacer column, then the new data column.
  if(last>0) sh.insertColumnsAfter(last,2);
  else sh.insertColumns(1,2);
  const dataCol=last+2;
  sh.getRange(hi.row,dataCol).setValue(label);
  return {name:label,col:dataCol};
}

function getResponses_() {
  const sh=SpreadsheetApp.openById(CFG.SPREADSHEET_ID).getSheetByName(CFG.RESPONSES_TAB);
  const vals=sh.getDataRange().getDisplayValues();
  return {ok:true,headers:vals[0]||[],rows:vals.slice(1).reverse().slice(0,500)};
}

function exportCsv_() {
  const result=listStudents_('');
  const headers=result.headers;
  const rows=result.records.map(r=>headers.map(h=>r[h]||''));
  const csv=[headers,...rows].map(r=>r.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(',')).join('\r\n');
  return {ok:true,csv:csv,filename:'students.csv'};
}

function findCol_(hi, aliases) {
  const list=aliases.split('|').map(norm_);
  const hit=hi.headers.find(h=>list.includes(norm_(h.name)));
  return hit?hit.col:0;
}
function norm_(s){return String(s||'').toLowerCase().replace(/[^a-z0-9]/g,'');}
function log_(action,details,actor) {
  const sh=SpreadsheetApp.openById(CFG.SPREADSHEET_ID).getSheetByName(CFG.LOGS_TAB);
  if(sh) sh.appendRow([new Date(),action,details,actor||'admin']);
}
