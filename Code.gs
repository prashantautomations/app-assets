/* =========================
   CONFIG
========================= */
const CFG = {
  VDB: 'V-Database',
  SHEET_RIGHTS: 'Sheetrights',
  DEFAULT_SHEET: 'PT',

  // Main sheet columns (1-based)
  COL_PO_STATUS: 1,
  COL_DATE_ONBOARDING: 2,
  COL_COMPANY: 3,
  COL_ACCOUNT: 4,
  COL_CLIENT_SUBJ: 5,
  COL_VENDOR_SUBJ: 6,
  COL_SOURCE: 7,
  COL_TARGET: 8,
  COL_EMAIL_USED: 9,
  COL_TRANSLATOR: 10,
  COL_TRANSLATOR_EMAIL: 11,
  COL_SERVICE: 12,
  COL_START_DATE: 13,
  COL_DELIVERY_DATE: 14,
  COL_WORDS: 15,
  COL_HOURS: 16,
  COL_TRANSLATOR_RATE: 17,
  COL_TRANSLATOR_VALUE: 18,
  COL_VENDOR_CURRENCY: 19,
  COL_CLIENT_CURRENCY: 20,
  COL_CLIENT_PO_NUMBER: 21,
  COL_CLIENT_PO_VALUE: 22,
  COL_PROJECT_MANAGER: 23,
  COL_ATTACHMENT: 24,
  COL_CLIENT_SERVICE_TYPE: 25,
  COL_ACCEPTED_THROUGH: 26,
  COL_REMARKS: 27,
  COL_QUALITY: 28,
  COL_ANUJ_REMARKS: 29,
  COL_DATE_UPDATION: 30,
  COL_BASE_TRANSLATOR_VALUE: 31,

  // V-Database columns
  COL_V_NAME: 3,
  COL_V_EMAIL: 4,
  COL_V_SRC: 8,
  COL_V_TGT: 9,
  COL_V_CURRENCY: 12,
  COL_V_RATE: 13,

  SRC_RANGE: 'H2:H',
  TGT_RANGE: 'I2:I',
  SERVICE_RANGE: 'M1:X1',

  // Header names
  HDR_VENDOR_SUBJECT: 'Vendor Subject',
  HDR_VENDOR_SUBJECT_ALT: 'Vendor Subject line',
  HDR_DATE_ONBOARDING: 'Date of Onboarding',
  HDR_DATE_UPDATION: 'Date of updation',
  HDR_PO_STATUS: 'PO Status',
  HDR_ATTACHMENT: 'Attachment (If Any)',
  HDR_CLIENT_SERVICE_TYPE: 'Client Service Type',
  HDR_ACCEPTED_THROUGH: 'Project Accepted through',
  HDR_SR_SHEET: 'Sheet',
  HDR_SR_EDITOR: 'Editor',
  HDR_SR_UNI_EDITOR: 'UNI Editor',
  HDR_BASE_TRANSLATOR_VALUE: 'Base Translator Invoice Value',
  HDR_PROJECT_MANAGER: 'Project Manager name'
};

const MAIN_SPREADSHEET_ID = '1iZYNM2KsCbMPGRJShNwSdcqrIOSQhRkiLtktJHQJPLc';
const VDB_SPREADSHEET_ID  = '1iZYNM2KsCbMPGRJShNwSdcqrIOSQhRkiLtktJHQJPLc';

function _getMainSS_() { return SpreadsheetApp.openById(MAIN_SPREADSHEET_ID); }
function _getVdbSS_()  { return SpreadsheetApp.openById(VDB_SPREADSHEET_ID); }
function _getVdbSheet_() {
  const sh = _getVdbSS_().getSheetByName(CFG.VDB);
  if (!sh) throw new Error(`Sheet not found in VDB spreadsheet: ${CFG.VDB}`);
  return sh;
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Team Tracker Entry')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================
   HELPERS
========================= */
function _getSheet_(sheetName) {
  const name = (sheetName || '').toString().trim();
  if (!name) throw new Error('No sheet selected.');
  const sh = _getMainSS_().getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  return sh;
}

function _norm_(s) { return (s ?? '').toString().trim().toLowerCase(); }

function _getHeaders_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => (h ?? '').toString().trim());
  let end = headers.length;
  while (end > 0 && !headers[end - 1]) end--;
  return headers.slice(0, end);
}

function _headerMap_(headers) {
  const map = {};
  headers.forEach((h, i) => { const key = _norm_(h); if (key) map[key] = i + 1; });
  return map;
}

function _findHeaderIndex_(headers, headerNameCandidates) {
  const hm = _headerMap_(headers);
  for (const cand of headerNameCandidates) {
    const idx = hm[_norm_(cand)];
    if (idx) return idx;
  }
  return null;
}

function _getEditorEmail_() {
  return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'user@unknown';
}

function _formatLogTimestamp_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function _isBlank_(v) { return v === '' || v === null || typeof v === 'undefined'; }

function _coerceDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const s = (v ?? '').toString().trim();
  if (!s) return null;
  const m = s.match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m) {
    const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4]||'0'), Number(m[5]||'0'), Number(m[6]||'0'));
    if (!isNaN(d.getTime())) return d;
  }
  const d2 = new Date(s);
  if (!isNaN(d2.getTime())) return d2;
  return null;
}

function _compare_(a, op, b) {
  const da = _coerceDate_(a), db = _coerceDate_(b);
  if (da && db) {
    const x = da.getTime(), y = db.getTime();
    switch(op){ case '=': return x===y; case '!=': return x!==y; case '<': return x<y; case '<=': return x<=y; case '>': return x>y; case '>=': return x>=y; }
  }
  const na = (typeof a==='number')?a:Number((a??'').toString().trim());
  const nb = Number((b??'').toString().trim());
  const aIsNum = typeof a==='number'||(!isNaN(na)&&(a??'').toString().trim()!=='');
  const bIsNum = !isNaN(nb)&&(b??'').toString().trim()!=='';
  if (aIsNum && bIsNum && ['=','!=','<','<=','>','>='].includes(op)) {
    switch(op){ case '=': return na===nb; case '!=': return na!==nb; case '<': return na<nb; case '<=': return na<=nb; case '>': return na>nb; case '>=': return na>=nb; }
  }
  const sa=(a??'').toString(), sb=(b??'').toString();
  switch(op){
    case '=': return _norm_(sa)===_norm_(sb);
    case '!=': return _norm_(sa)!==_norm_(sb);
    case 'contains': return _norm_(sa).includes(_norm_(sb));
    case 'not_contains': return !_norm_(sa).includes(_norm_(sb));
    case 'starts_with': return _norm_(sa).startsWith(_norm_(sb));
    case 'ends_with': return _norm_(sa).endsWith(_norm_(sb));
    case 'is_blank': return _isBlank_(a);
    case 'is_not_blank': return !_isBlank_(a);
    case '<': return sa<sb; case '<=': return sa<=sb; case '>': return sa>sb; case '>=': return sa>=sb;
    default: return false;
  }
}

function _evalFiltersLeftToRight_(rowObj, filters) {
  let acc = true, first = true;
  for (const f of (filters||[])) {
    const cell = rowObj.values[f.colIndex - 1];
    const ok = _compare_(cell, f.op, f.value);
    if (first) { acc = ok; first = false; continue; }
    if (f.join === 'OR') acc = acc || ok; else acc = acc && ok;
  }
  return acc;
}

function _toNum_(v) {
  if (typeof v === 'number') return v;
  const s = (v||'').toString().replace(/,/g,'').replace('%','').trim();
  if (!s) return 0;
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function _getPenaltyRateFromQuality_(qualityValue) {
  const q = _toNum_(qualityValue);
  if (!qualityValue || qualityValue.toString().trim() === '') return null;
  if (q >= 85) return 0;
  if (q >= 75) return 0.10;
  if (q >= 50) return 0.25;
  if (q >= 0)  return 0.50;
  return 0;
}

function _recalculateFinalTranslatorInvoiceValue_(sh, row) {
  const quality = sh.getRange(row, CFG.COL_QUALITY).getValue();
  if (quality === '' || quality === null || typeof quality === 'undefined') return;
  const penaltyRate = _getPenaltyRateFromQuality_(quality);
  if (penaltyRate === null) return;
  const baseCell = sh.getRange(row, CFG.COL_BASE_TRANSLATOR_VALUE);
  let baseValue = _toNum_(baseCell.getValue());
  if (!baseValue) {
    baseValue = _toNum_(sh.getRange(row, CFG.COL_TRANSLATOR_VALUE).getValue());
    if (!baseValue) return;
    baseCell.setValue(baseValue);
  }
  sh.getRange(row, CFG.COL_TRANSLATOR_VALUE).setValue(Number((baseValue * (1 - penaltyRate)).toFixed(3)));
}

function _applyPOStatusRowColor_(sh, row) {
  const sheetName = sh.getName();
  if (!_isInSheetRights_(sheetName)) return;
  const status = (sh.getRange(row, CFG.COL_PO_STATUS).getDisplayValue()||'').toString().trim().toLowerCase();
  const C = { CYAN:'#00FFFF', AMBER:'#f5b310', RED:'#FF0000', WHITE:'#FFFFFF', GREEN:'#00B050', LG1:'#b6d7a8' };
  let cA = C.WHITE, cBN = C.WHITE;
  switch(status){
    case 'po issued but not updated':                  cA=C.CYAN;  cBN=C.AMBER; break;
    case 'po pending':                                 cA=C.RED;   cBN=C.WHITE; break;
    case 'po issued':                                  cA=C.GREEN; cBN=C.AMBER; break;
    case 'project cancelled for vendor':               cA=C.RED;   cBN=C.AMBER; break;
    case 'project cancelled from client':              cA=C.WHITE; cBN=C.RED;   break;
    case 'line item moved':
    case 'line item moved but client po is pending':   cA=C.LG1;   cBN=C.LG1;   break;
    case 'po not for us':                              cA=C.WHITE; cBN=C.AMBER; break;
  }
  sh.getRange(row, 1).setBackground(cA);
  sh.getRange(row, 2, 1, 13).setBackground(cBN);
}

function _serviceDisplayName_(header) {
  const raw = (header||'').toString().trim();
  if (!raw) return '';
  return raw.replace(/\s*\([^)]+\)\s*$/, '').trim();
}

function _normalizeServiceKey_(s) {
  return (s||'').toString().trim().toLowerCase()
    .replace(/\s*\([^)]+\)\s*/g,'').replace(/per\s+hour/g,'').replace(/[^a-z0-9]/g,'');
}

function _getServiceRateColumn_(vdb, serviceName) {
  const wanted = _normalizeServiceKey_(serviceName);
  if (!wanted) return CFG.COL_V_RATE;
  const headers = vdb.getRange(1, 1, 1, vdb.getLastColumn()).getDisplayValues()[0];
  for (let i = 0; i < headers.length; i++) {
    const candidate = _normalizeServiceKey_(headers[i]);
    if (candidate && candidate === wanted) return i + 1;
  }
  return CFG.COL_V_RATE;
}

/* =========================
   DRAFT EMAIL BUILDER
========================= */
function _buildDraftEmail_(p, vendorSubject) {
  const projectNo  = vendorSubject || p.vendorSubject || '';
  const source     = p.source || '';
  const target     = p.target || '';
  const deadline   = p.deliveryDate || '';
  const svc        = (p.clientServiceType || p.service || '').toString().trim();
  const svcNorm    = svc.toLowerCase().replace(/\s+/g,' ').trim();

  function row(label, value) {
    return `<tr><td style="padding:4px 12px 4px 0;font-weight:600;white-space:nowrap;vertical-align:top;">${label}</td><td style="padding:4px 0;">${value}</td></tr>`;
  }

  const NOTE = `<b>Note:</b> Quality expectations are 92% and above; if the quality is less than 85%, then there will be a straight penalty of 10%, and if less than 75%, then a 25% and if less than 50%, then 50% penalty will apply to the complete project cost.`;
  const FOOTER = `<p>Please let me know if you have any questions.</p><p>Kindly confirm the safe receipt of this email.</p>`;

  function buildBody(serviceLabel, toolDefault) {
    return `<p>Hello,</p>
<p>I hope you are doing well.</p>
<p>We have received a new <b>${serviceLabel}</b> project.</p>
<p>Please find the details of the project below:</p>
<table style="border-collapse:collapse;font-size:14px;">
  ${row('Project No.:', projectNo)}
  ${row('Service:', serviceLabel)}
  ${row('Source:', source)}
  ${row('Target:', target)}
  ${row('Deadline:', deadline)}
  ${row('Tool:', toolDefault)}
  ${row('TM Name:', '&nbsp;')}
  ${row('Token:', '&nbsp;')}
</table>
<br>
<p>${NOTE}</p>
${FOOTER}`;
  }

  let serviceLabel, tool, subject;

  if (svcNorm.includes('mtpe') || svcNorm.includes('machine translation')) {
    serviceLabel = 'MTPE'; tool = 'Translation Workspace';
  } else if (svcNorm.includes('editing') || svcNorm.includes('post editing')) {
    serviceLabel = 'Editing/Post Editing'; tool = 'Translation Workspace';
  } else if (svcNorm.includes('review')) {
    serviceLabel = 'Review + Editing'; tool = 'Translation Workspace';
  } else if (svcNorm.includes('lso')) {
    serviceLabel = 'LSO Check'; tool = '';
  } else if (svcNorm.includes('lqs') || svcNorm.includes('lqa') || svcNorm.includes('lqi')) {
    serviceLabel = 'LQS/LQA/LQI'; tool = '';
  } else if (svcNorm.includes('proofreading')) {
    serviceLabel = 'Proofreading'; tool = 'Translation Workspace';
  } else if (svcNorm.includes('adaptation')) {
    serviceLabel = 'Adaptation'; tool = '';
  } else if (svcNorm.includes('voice')) {
    serviceLabel = 'Voice Recording'; tool = '';
  } else if (svcNorm.includes('subtitl')) {
    serviceLabel = 'Subtitling'; tool = '';
  } else if (svcNorm.includes('transcreation')) {
    serviceLabel = 'Transcreation'; tool = '';
  } else if (svcNorm.includes('transcription')) {
    serviceLabel = 'Transcription'; tool = '';
  } else if (svcNorm.includes('test')) {
    serviceLabel = 'Test Translation'; tool = '';
  } else if (svcNorm.includes('implementation')) {
    serviceLabel = 'Implementation'; tool = '';
  } else {
    serviceLabel = 'Translation'; tool = '';
  }

  subject = vendorSubject || projectNo || `New ${serviceLabel} Project - ${source} to ${target}`;
  return { subject, htmlBody: buildBody(serviceLabel, tool) };
}

function _createGmailDraft_(p, vendorSubject) {
  const toEmail = (p.translatorEmail || '').toString().trim();
  if (!toEmail || toEmail.toLowerCase() === 'self') return null;
  const { subject, htmlBody } = _buildDraftEmail_(p, vendorSubject);
  try {
    GmailApp.createDraft(toEmail, subject, '', { htmlBody });
    return { draftCreated: true, draftTo: toEmail, draftSubject: subject };
  } catch(e) {
    Logger.log('Draft creation failed: ' + e.message);
    return { draftCreated: false, draftError: e.message };
  }
}

/* =========================
   INIT DATA
========================= */
function getInitData() {
  const vdb = _getVdbSheet_();
  const srcVals = vdb.getRange(CFG.SRC_RANGE).getValues().flat().map(v=>(v||'').toString().trim()).filter(Boolean);
  const tgtVals = vdb.getRange(CFG.TGT_RANGE).getValues().flat().map(v=>(v||'').toString().trim()).filter(Boolean);
  const src = [...new Set(srcVals)].sort();
  const tgt = [...new Set(tgtVals)].sort();
  const services = vdb.getRange(CFG.SERVICE_RANGE).getDisplayValues()[0].map(v=>_serviceDisplayName_(v)).filter(Boolean);
  return { src, tgt, services };
}

function getTargetsForSource(source) {
  const vdb = _getVdbSheet_();
  const last = vdb.getLastRow();
  if (last < 2) return [];
  const data = vdb.getRange(2, 1, last - 1, vdb.getLastColumn()).getValues();
  return [...new Set(
    data
      .filter(r => String(r[CFG.COL_V_SRC-1]).trim().toLowerCase() === source.toLowerCase())
      .map(r => (r[CFG.COL_V_TGT-1]||'').toString().trim())
      .filter(Boolean)
  )].sort();
}

function _appendSelfTranslatorOption_(list) {
  const out = Array.isArray(list) ? list.slice() : [];
  const hasSelf = out.some(t => _norm_(t?.name||'')==='self' && _norm_(t?.email||'')==='self');
  if (!hasSelf) out.push({ name:'Self', email:'Self', rate:'0', currency:'EUR' });
  return out;
}

function getTranslators(src, tgt, service) {
  const vdb = _getVdbSheet_();
  const last = vdb.getLastRow();
  if (last < 2) return _appendSelfTranslatorOption_([]);
  const data = vdb.getRange(2, 1, last-1, vdb.getLastColumn()).getDisplayValues();
  const rateCol = _getServiceRateColumn_(vdb, service);
  const out = [], seen = new Set();
  data.forEach(r => {
    const rSrc = (r[CFG.COL_V_SRC-1]||'').toString().trim().toLowerCase();
    const rTgt = (r[CFG.COL_V_TGT-1]||'').toString().trim().toLowerCase();
    if (rSrc === src.toLowerCase() && rTgt === tgt.toLowerCase()) {
      const name     = (r[CFG.COL_V_NAME-1]||'').toString().trim();
      const email    = (r[CFG.COL_V_EMAIL-1]||'').toString().trim();
      const rate     = (r[rateCol-1]||'').toString().trim();
      const currency = (r[CFG.COL_V_CURRENCY-1]||'').toString().trim();
      const key = `${name}__${email}__${rate}__${currency}`;
      if (name && email && !seen.has(key)) { seen.add(key); out.push({name,email,rate,currency}); }
    }
  });
  out.sort((a,b)=>a.name.localeCompare(b.name));
  return _appendSelfTranslatorOption_(out);
}

/* =========================
   SHEET NAMES
========================= */
function getAllSheetNames() {
  try {
    const sh = _getMainSS_().getSheetByName(CFG.SHEET_RIGHTS);
    if (!sh) throw new Error(`Sheet Rights tab "${CFG.SHEET_RIGHTS}" not found.`);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return [];
    return [...new Set(
      sh.getRange(2, 1, lastRow-1, 1).getDisplayValues().flat()
        .map(v=>(v||'').toString().trim()).filter(Boolean)
    )];
  } catch(e) { Logger.log('getAllSheetNames error: '+e.message); return []; }
}

/* =========================
   SHEET RIGHTS HELPERS
========================= */
function _getSheetRightsSheet_() {
  const sh = _getMainSS_().getSheetByName(CFG.SHEET_RIGHTS);
  if (!sh) throw new Error(`Sheet not found: ${CFG.SHEET_RIGHTS}`);
  return sh;
}

function _normalizeEmail_(s) { return (s||'').toString().trim().toLowerCase(); }

function _splitEmails_(value) {
  return (value||'').toString().split(/[,\n;]/).map(v=>_normalizeEmail_(v)).filter(Boolean);
}

function _getSheetRightsData_() {
  const sh = _getSheetRightsSheet_();
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const headers = _getHeaders_(sh), hm = _headerMap_(headers);
  const colSheet = hm[_norm_(CFG.HDR_SR_SHEET)];
  const colEditor = hm[_norm_(CFG.HDR_SR_EDITOR)];
  const colUniEditor = hm[_norm_(CFG.HDR_SR_UNI_EDITOR)];
  if (!colSheet||!colEditor||!colUniEditor) throw new Error(
    `Sheetrights must have headers: "${CFG.HDR_SR_SHEET}", "${CFG.HDR_SR_EDITOR}", "${CFG.HDR_SR_UNI_EDITOR}"`);
  const data = sh.getRange(2, 1, lastRow-1, lastCol).getDisplayValues();
  return data.map(r => ({
    sheet:      (r[colSheet-1]||'').toString().trim(),
    editors:    _splitEmails_(r[colEditor-1]),
    uniEditors: _splitEmails_(r[colUniEditor-1])
  }));
}

function _getAllUniEditors_() {
  const rows = _getSheetRightsData_(), set = new Set();
  rows.forEach(r=>(r.uniEditors||[]).forEach(e=>set.add(e)));
  return [...set];
}

function _isInSheetRights_(sheetName) {
  try { return _getSheetRightsData_().some(r=>_norm_(r.sheet)===_norm_(sheetName)); }
  catch(e) { return false; }
}

function _canEditSheet_(sheetName, email) {
  const userEmail = _normalizeEmail_(email);
  if (!userEmail) return false;
  const rows = _getSheetRightsData_();
  if (new Set(_getAllUniEditors_()).has(userEmail)) return true;
  const row = rows.find(r=>_norm_(r.sheet)===_norm_(sheetName));
  if (!row) return false;
  return (row.editors||[]).includes(userEmail);
}

function _assertSheetAccess_(sheetName) {
  const email = _getEditorEmail_();
  if (!_canEditSheet_(sheetName, email))
    throw new Error(`Access denied. ${email} does not have permission for sheet: ${sheetName}`);
}

/* =========================
   EDIT/CLONE: Load by Vendor Subject
========================= */
function findByVendorSubject(vendorSubject, sheetName) {
  _assertSheetAccess_(sheetName);
  const sh = _getSheet_(sheetName);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const headers = _getHeaders_(sh);
  const vendorCol = _findHeaderIndex_(headers,[CFG.HDR_VENDOR_SUBJECT,CFG.HDR_VENDOR_SUBJECT_ALT]) || CFG.COL_VENDOR_SUBJ;
  const wanted = (vendorSubject||'').toString().trim();
  const values = sh.getRange(2, vendorCol, lastRow-1, 1).getDisplayValues().flat();
  const idx = values.findIndex(v=>(v||'').toString().trim()===wanted);
  if (idx===-1) return null;
  const row = idx+2;
  const r = sh.getRange(row, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  return JSON.parse(JSON.stringify({
    row,
    company:          r[CFG.COL_COMPANY-1]||'',
    account:          r[CFG.COL_ACCOUNT-1]||'',
    clientSubject:    r[CFG.COL_CLIENT_SUBJ-1]||'',
    vendorSubject:    r[vendorCol-1]||'',
    source:           r[CFG.COL_SOURCE-1]||'',
    target:           r[CFG.COL_TARGET-1]||'',
    emailUsed:        r[CFG.COL_EMAIL_USED-1]||'',
    translator:       r[CFG.COL_TRANSLATOR-1]||'',
    translatorEmail:  r[CFG.COL_TRANSLATOR_EMAIL-1]||'',
    translatorRate:   r[CFG.COL_TRANSLATOR_RATE-1]||'',
    vendorCurrency:   r[CFG.COL_VENDOR_CURRENCY-1]||'',
    translatorValue:  r[CFG.COL_TRANSLATOR_VALUE-1]||'',
    service:          r[CFG.COL_SERVICE-1]||'',
    startDate:        r[CFG.COL_START_DATE-1]||'',
    deliveryDate:     r[CFG.COL_DELIVERY_DATE-1]||'',
    words:            r[CFG.COL_WORDS-1]||'',
    hours:            r[CFG.COL_HOURS-1]||'',
    clientServiceType: r[CFG.COL_CLIENT_SERVICE_TYPE-1]||'',
    acceptedThrough:  r[CFG.COL_ACCEPTED_THROUGH-1]||'',
    remarks:          r[CFG.COL_REMARKS-1]||'',
    quality:          r[CFG.COL_QUALITY-1]||'',
    attachment:       r[CFG.COL_ATTACHMENT-1]||'',
    projectManager:   r[CFG.COL_PROJECT_MANAGER-1]||''
  }));
}

/* =========================
   SUBJECT GENERATOR
========================= */
function getNextAvailableRow_(sheet) {
  const startRow = 2, lastRow = sheet.getLastRow();
  if (lastRow < startRow) return startRow;
  const colVals = sheet.getRange(startRow, CFG.COL_SOURCE, lastRow-startRow+1, 1).getDisplayValues().flat();
  const idx = colVals.findIndex(v=>!String(v||'').trim());
  return idx===-1 ? lastRow+1 : startRow+idx;
}

function getNextGlobalProjectSequence_() {
  const ss = _getMainSS_();
  const allowedNames = new Set(getAllSheetNames().map(n=>_norm_(n)));
  let maxNum = 0;
  const re = /\/[A-Z](\d{6})-/;
  ss.getSheets().forEach(sh => {
    if (!allowedNames.has(_norm_(sh.getName().trim()))) return;
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;
    const headers = _getHeaders_(sh);
    const vendorCol = _findHeaderIndex_(headers,[CFG.HDR_VENDOR_SUBJECT_ALT,CFG.HDR_VENDOR_SUBJECT]) || CFG.COL_VENDOR_SUBJ;
    if (!vendorCol) return;
    const numRows = lastRow-1;
    if (numRows<=0) return;
    sh.getRange(2, vendorCol, numRows, 1).getDisplayValues().flat().forEach(s=>{
      const text=(s||'').toString().trim();
      if (!text) return;
      const m=text.match(re);
      if (!m) return;
      const num=Number(m[1]);
      if (num>maxNum) maxNum=num;
    });
  });
  return String(maxNum+1).padStart(6,'0');
}

function getMaxGlobalProjectSequence_() {
  const ss = _getMainSS_();
  const allowedNames = new Set(getAllSheetNames().map(n=>_norm_(n)));
  let maxNum = 0;
  const re = /\/[A-Z](\d{6})-/;
  ss.getSheets().forEach(sh => {
    if (!allowedNames.has(_norm_(sh.getName().trim()))) return;
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;
    const headers = _getHeaders_(sh);
    const vendorCol = _findHeaderIndex_(headers,[CFG.HDR_VENDOR_SUBJECT_ALT,CFG.HDR_VENDOR_SUBJECT]) || CFG.COL_VENDOR_SUBJ;
    if (!vendorCol) return;
    const numRows = lastRow-1;
    if (numRows<=0) return;
    sh.getRange(2, vendorCol, numRows, 1).getDisplayValues().flat().forEach(s=>{
      const text=(s||'').toString().trim();
      if (!text) return;
      const m=text.match(re);
      if (!m) return;
      const num=Number(m[1]);
      if (num>maxNum) maxNum=num;
    });
  });
  return maxNum;
}

function extractProjectNumber_(subject) {
  const m=(subject||'').match(/\/([A-Z]\d{6})-/);
  return m ? m[1] : null;
}

function getVendorSubjectMode_(p) {
  if (p.useHours) return { label:'Hour', qty:p.hours||'' };
  return { label:'WWC', qty:p.words||'' };
}

function generateVendorSubjectWithSameProject_(oldSubject, p) {
  const projectCode = extractProjectNumber_(oldSubject);
  if (!projectCode) throw new Error('Invalid vendor subject format. Cannot clone project.');
  const email = _getEditorEmail_();
  const user4 = (email.split('@')[0]||'').substring(0,4).toLowerCase();
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'ddMMyy');
  const mode = getVendorSubjectMode_(p);
  return `Off/${user4}/${dateStr}/${projectCode}- ${p.source} to ${p.target} ${mode.label}- ${mode.qty}`;
}

function generateVendorSubject(p) {
  const email = _getEditorEmail_();
  const user4 = (email.split('@')[0]||'').substring(0,4).toLowerCase();
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'ddMMyy');
  const projectLetter = ((p.sheetName||'').toString().trim().charAt(0)||'P').toUpperCase();
  const seq = getNextGlobalProjectSequence_();
  const mode = getVendorSubjectMode_(p);
  return `Off/${user4}/${dateStr}/${projectLetter}${seq}- ${p.source} to ${p.target} ${mode.label}- ${mode.qty}`;
}

function generateVendorSubjectWithSeq_(p, seq) {
  const email = _getEditorEmail_();
  const user4 = (email.split('@')[0]||'').substring(0,4).toLowerCase();
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'ddMMyy');
  const projectLetter = ((p.sheetName||'').toString().trim().charAt(0)||'P').toUpperCase();
  const mode = getVendorSubjectMode_(p);
  return `Off/${user4}/${dateStr}/${projectLetter}${seq}- ${p.source} to ${p.target} ${mode.label}- ${mode.qty}`;
}

/* =========================
   SUBMIT (single entry)
========================= */
function submit(p) {
  _assertSheetAccess_(p.sheetName);
  const sh = _getSheet_(p.sheetName);

  if (p.cloneProject && !p.editRow) throw new Error('Clone mode requires an existing project to be loaded.');
  if (!p.source || !p.target || !p.translator) throw new Error('Source, Target, and Translator are required.');

  const row = p.cloneProject
    ? getNextAvailableRow_(sh)
    : (p.editRow ? Number(p.editRow) : getNextAvailableRow_(sh));

  let vendorSubject;
  if (p.cloneProject)  vendorSubject = generateVendorSubjectWithSameProject_(p.vendorSubject, p);
  else if (p.editRow)  vendorSubject = p.vendorSubject;
  else                 vendorSubject = generateVendorSubject(p);

  _writeProjectRow_(sh, row, p, vendorSubject);

  if (p.cloneProject || !p.editRow) {
    const headers = _getHeaders_(sh);
    const onboardingCol = _findHeaderIndex_(headers,[CFG.HDR_DATE_ONBOARDING]);
    if (onboardingCol) sh.getRange(row, onboardingCol).setValue(new Date());
  }

  const result = { ok: true, vendorSubject, row };

  if (p.makeDraft) {
    const draftResult = _createGmailDraft_(p, vendorSubject);
    if (draftResult) Object.assign(result, draftResult);
  }

  return result;
}

/* =========================
   BULK SUBMIT
========================= */
function submitBulk(payload) {
  _assertSheetAccess_(payload.sheetName);
  const sh = _getSheet_(payload.sheetName);
  const rows = payload.rows || [];
  if (!rows.length) throw new Error('No rows to submit.');

  let seqNum = getMaxGlobalProjectSequence_();

  const headers = _getHeaders_(sh);
  const onboardingCol = _findHeaderIndex_(headers,[CFG.HDR_DATE_ONBOARDING]);
  const now = new Date();

  const submitted = [];
  const errors = [];

  rows.forEach((p, idx) => {
    try {
      if (!p.source || !p.target) throw new Error('Source and Target are required.');
      p.sheetName = payload.sheetName;

      seqNum += 1;
      const seq = String(seqNum).padStart(6,'0');
      const vendorSubject = generateVendorSubjectWithSeq_(p, seq);

      const row = getNextAvailableRow_(sh);
      _writeProjectRow_(sh, row, p, vendorSubject);
      if (onboardingCol) sh.getRange(row, onboardingCol).setValue(now);

      submitted.push({ index: idx, row, vendorSubject });
    } catch(e) {
      errors.push({ index: idx, error: e.message });
    }
  });

  return { ok: true, submitted: submitted.length, errors, results: submitted };
}

/**
 * Shared row-writing logic used by both submit() and submitBulk()
 */
function _writeProjectRow_(sh, row, p, vendorSubject) {
  sh.getRange(row, CFG.COL_COMPANY).setValue(p.company||'');
  sh.getRange(row, CFG.COL_ACCOUNT).setValue(p.account||'');
  sh.getRange(row, CFG.COL_CLIENT_SUBJ).setValue(p.clientSubject||'');
  sh.getRange(row, CFG.COL_VENDOR_SUBJ).setValue(vendorSubject);
  sh.getRange(row, CFG.COL_SOURCE).setValue(p.source||'');
  sh.getRange(row, CFG.COL_TARGET).setValue(p.target||'');
  sh.getRange(row, CFG.COL_EMAIL_USED).setValue(p.emailUsed||'');
  sh.getRange(row, CFG.COL_TRANSLATOR).setValue(p.translator||'');
  sh.getRange(row, CFG.COL_TRANSLATOR_EMAIL).setValue(p.translatorEmail||'');
  sh.getRange(row, CFG.COL_SERVICE).setValue(p.service||'');
  sh.getRange(row, CFG.COL_START_DATE).setValue(p.startDate||'');
  sh.getRange(row, CFG.COL_DELIVERY_DATE).setValue(p.deliveryDate||'');
  sh.getRange(row, CFG.COL_TRANSLATOR_RATE).setValue(p.translatorRate||'');
  sh.getRange(row, CFG.COL_VENDOR_CURRENCY).setValue(p.vendorCurrency||'');
  sh.getRange(row, CFG.COL_WORDS).setValue(p.words||'');
  sh.getRange(row, CFG.COL_HOURS).setValue(p.hours||'');
  sh.getRange(row, CFG.COL_CLIENT_SERVICE_TYPE).setValue(p.clientServiceType||'');
  sh.getRange(row, CFG.COL_ACCEPTED_THROUGH).setValue(p.acceptedThrough||'');
  sh.getRange(row, CFG.COL_QUALITY).setValue(p.quality||'');
  sh.getRange(row, CFG.COL_REMARKS).setValue(p.remarks||'');
  sh.getRange(row, CFG.COL_TRANSLATOR_VALUE).setValue(p.translatorValue||'');
  sh.getRange(row, CFG.COL_BASE_TRANSLATOR_VALUE).setValue(p.translatorValue||'');
  sh.getRange(row, CFG.COL_PROJECT_MANAGER).setValue(p.projectManager||'');
  sh.getRange(row, CFG.COL_ATTACHMENT).setValue(p.attachment||'');
}

/* =========================
   FILTER UI SUPPORT
========================= */
function getSheetColumns(sheetName) {
  _assertSheetAccess_(sheetName);
  const sh = _getSheet_(sheetName);
  return _getHeaders_(sh).filter(h=>(h??'').toString().trim()!=='');
}

/* =========================
   QUERY SHEET WITH FILTERS + PAGINATION
========================= */
function querySheetRows(payload) {
  _assertSheetAccess_(payload.sheetName);
  const sh = _getSheet_(payload.sheetName);
  const pageSize = Math.max(1, Number(payload.pageSize||50));
  const page     = Math.max(1, Number(payload.page||1));
  const headers  = _getHeaders_(sh).map(h=>(h??'').toString());
  const lastCol  = headers.length;
  const lastRow  = sh.getLastRow();
  if (lastRow < 2 || lastCol < 1) return { headers, total:0, page:1, pageSize, maxPage:1, rows:[] };
  const values = sh.getRange(2, 1, lastRow-1, lastCol).getDisplayValues();
  const headerMap = _headerMap_(headers);
  const filters = (payload.filters||[]).map(f=>{
    const colIndex = headerMap[_norm_(f.column)];
    if (!colIndex) return null;
    return { join:(f.join||'AND').toUpperCase(), colIndex, op:f.operator, value:(f.value??'').toString() };
  }).filter(Boolean);
  const all = values.map((r,i)=>({ rowNumber:i+2, values:r.map(v=>v??'') }));
  const matched = filters.length ? all.filter(ro=>_evalFiltersLeftToRight_(ro,filters)) : all;
  const total   = matched.length;
  const maxPage = Math.max(1, Math.ceil(total/pageSize));
  const safePage = Math.min(Math.max(1,page), maxPage);
  const start   = (safePage-1)*pageSize;
  return JSON.parse(JSON.stringify({ headers, total, page:safePage, pageSize, maxPage, rows:matched.slice(start,start+pageSize) }));
}

/* =========================
   UPDATE STATUS FOR ROWS
========================= */
function updateStatusForRows(payload) {
  _assertSheetAccess_(payload.sheetName);
  const sh = _getSheet_(payload.sheetName);
  const headers = _getHeaders_(sh);
  const statusCol = _findHeaderIndex_(headers,[CFG.HDR_PO_STATUS])||1;
  const rowNumbers = (payload.rowNumbers||[]).map(n=>Number(n)).filter(Boolean);
  const newStatus  = (payload.newStatus??'').toString().trim();
  if (!rowNumbers.length) throw new Error('No rows selected.');
  if (!newStatus) throw new Error('Status is required.');
  const allowedStatuses = getPOStatusOptions(payload.sheetName);
  if (allowedStatuses.length && !allowedStatuses.includes(newStatus)) throw new Error('Invalid status selected.');
  rowNumbers.forEach(rn=>{
    sh.getRange(rn, statusCol).setValue(newStatus);
    _applyPOStatusRowColor_(sh, rn);
    _applyPOStatusValueRules_(sh, rn);
  });
  const reason = (payload.reason??'').toString().trim();
  if (reason) _appendUpdationLog_(sh, headers, rowNumbers, reason);
  return { ok:true, updated:rowNumbers.length };
}

/* =========================
   BULK EDIT ROWS + DATE OF UPDATION LOG
========================= */
function saveEdits(payload) {
  _assertSheetAccess_(payload.sheetName);
  const sh = _getSheet_(payload.sheetName);
  const headers = _getHeaders_(sh), hm = _headerMap_(headers);
  const edits = payload.edits||[];
  if (!edits.length) throw new Error('No edits found.');
  let changedRows = new Set();
  edits.forEach(ed=>{
    const rn=Number(ed.rowNumber); if (!rn) return;
    Object.keys(ed.cells||{}).forEach(h=>{
      const col=hm[_norm_(h)]; if (!col) return;
      sh.getRange(rn, col).setValue(ed.cells[h]);
      changedRows.add(rn);
    });
  });
  const reasonMap = {};
  (payload.reasons||[]).forEach(r=>{ if (r&&r.rowNumber) reasonMap[Number(r.rowNumber)]=(r.reason??'').toString().trim(); });
  const updCol = _findHeaderIndex_(headers,[CFG.HDR_DATE_UPDATION]);
  if (updCol) {
    const email = _getEditorEmail_(), ts = _formatLogTimestamp_();
    [...changedRows].forEach(rn=>{
      _applyPOStatusRowColor_(sh, rn);
      _applyPOStatusValueRules_(sh, rn);
      _recalculateFinalTranslatorInvoiceValue_(sh, rn);
      const reason = reasonMap[rn]||'';
      const log = `[${ts}] ${email}:${reason||'updated'}`;
      const cell = sh.getRange(rn, updCol);
      const prev = (cell.getValue()??'').toString().trim();
      cell.setValue(prev ? prev+'\n'+log : log);
    });
  }
  return { ok:true, updated:changedRows.size };
}

function _appendUpdationLog_(sh, headers, rowNumbers, reason) {
  const updCol = _findHeaderIndex_(headers,[CFG.HDR_DATE_UPDATION]);
  if (!updCol) return;
  const email = _getEditorEmail_(), ts = _formatLogTimestamp_();
  rowNumbers.forEach(rn=>{
    const cell = sh.getRange(rn, updCol);
    const prev = (cell.getValue()??'').toString().trim();
    const log = `[${ts}] ${email}:${reason}`;
    cell.setValue(prev ? prev+'\n'+log : log);
  });
}

/* =========================
   TRANSLATOR DETAILS
========================= */
function getTranslatorDetails(src, tgt, translatorName, translatorEmail, service) {
  const nameNorm  = _norm_(translatorName||'');
  const emailNorm = _norm_(translatorEmail||'');
  if (nameNorm==='self'||emailNorm==='self') return { email:'Self', rate:'0', currency:'EUR' };
  const vdb = _getVdbSheet_();
  const last = vdb.getLastRow();
  if (last<2) return { email:'', rate:'', currency:'' };
  const data = vdb.getRange(2, 1, last-1, vdb.getLastColumn()).getDisplayValues();
  const rateCol = _getServiceRateColumn_(vdb, service);
  const srcNorm = _norm_(src||''), tgtNorm = _norm_(tgt||'');
  let row = data.find(r=>
    _norm_(r[CFG.COL_V_SRC-1])===srcNorm && _norm_(r[CFG.COL_V_TGT-1])===tgtNorm &&
    _norm_(r[CFG.COL_V_NAME-1])===nameNorm && _norm_(r[CFG.COL_V_EMAIL-1])===emailNorm);
  if (!row) row = data.find(r=>
    _norm_(r[CFG.COL_V_SRC-1])===srcNorm && _norm_(r[CFG.COL_V_TGT-1])===tgtNorm &&
    _norm_(r[CFG.COL_V_NAME-1])===nameNorm);
  if (!row) return { email:'', rate:'', currency:'' };
  return {
    email:    (row[CFG.COL_V_EMAIL-1]||'').toString().trim(),
    rate:     (row[rateCol-1]||'').toString().trim(),
    currency: (row[CFG.COL_V_CURRENCY-1]||'').toString().trim()
  };
}

/* =========================
   PO STATUS OPTIONS
========================= */
function getPOStatusOptions(sheetName) {
  const sh = _getSheet_(sheetName);
  const headers = _getHeaders_(sh);
  const statusCol = _findHeaderIndex_(headers,[CFG.HDR_PO_STATUS])||1;
  const rule = sh.getRange(2, statusCol).getDataValidation();
  if (rule) {
    const criteria = rule.getCriteriaType(), args = rule.getCriteriaValues();
    if (criteria===SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST)
      return ((args&&args[0])?args[0]:[]).map(v=>String(v).trim()).filter(Boolean);
    if (criteria===SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE && args&&args[0])
      return args[0].getDisplayValues().flat().map(v=>String(v).trim()).filter(Boolean);
  }
  return ['PO Pending','PO Issued','PO Issued but Not Updated','Project Cancelled for Vendor',
    'Project cancelled from client','Line Item Moved','Line Item Moved but Client PO Is Pending','PO not for us'];
}

function _applyPOStatusValueRules_(sh, row) {
  const status = _norm_(sh.getRange(row, CFG.COL_PO_STATUS).getDisplayValue()||'');
  if (status==='project cancelled for vendor') {
    sh.getRange(row, CFG.COL_TRANSLATOR_VALUE).setValue(0);
    sh.getRange(row, CFG.COL_BASE_TRANSLATOR_VALUE).setValue(0);
  }
}

/* =========================
   ON EDIT TRIGGER
========================= */
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (!_isInSheetRights_(sh.getName())) return;
    const startRow=e.range.getRow(), numRows=e.range.getNumRows();
    const startCol=e.range.getColumn(), numCols=e.range.getNumColumns();
    if (startRow<2) return;
    const endRow=startRow+numRows-1, endCol=startCol+numCols-1;
    if (startCol<=CFG.COL_TRANSLATOR_VALUE && endCol>=CFG.COL_TRANSLATOR_VALUE)
      for (let row=startRow;row<=endRow;row++)
        sh.getRange(row,CFG.COL_BASE_TRANSLATOR_VALUE).setValue(sh.getRange(row,CFG.COL_TRANSLATOR_VALUE).getValue());
    if (startCol<=CFG.COL_QUALITY && endCol>=CFG.COL_QUALITY)
      for (let row=startRow;row<=endRow;row++) _recalculateFinalTranslatorInvoiceValue_(sh, row);
    if (startCol<=CFG.COL_PO_STATUS && endCol>=CFG.COL_PO_STATUS)
      for (let row=startRow;row<=endRow;row++) { _applyPOStatusRowColor_(sh,row); _applyPOStatusValueRules_(sh,row); }
  } catch(err) { Logger.log(err); }
}
