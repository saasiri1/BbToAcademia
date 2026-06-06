
let lang           = 'ar';
let bbData         = null;
let ugWorkbook     = null;
let ugHeaderRow    = -1;
let bbColumns      = [];
let resultWorkbook = null;


let finalFileData    = null;
let finalFileHeaders = [];

let lastResults = null;


function t(key, ...args) {
  const val = translations[lang][key];
  return typeof val === 'function' ? val(...args) : (val || key);
}

function applyLang() {
  const html = document.documentElement;
  html.lang = lang;
  html.dir  = lang === 'ar' ? 'rtl' : 'ltr';
  document.title = lang === 'ar'
    ? 'محول درجات بلاك بورد'
    : 'Blackboard Grade Converter';

  document.querySelectorAll('[data-i18n]').forEach(el => {
    el.textContent = t(el.dataset.i18n);
  });

  document.getElementById('midtermMax').placeholder = t('placeholder');
  document.getElementById('finalMax').placeholder   = t('placeholder');

  // Refresh select placeholders
  ['finalFileStudentCol', 'finalFileGradeCol'].forEach(id => {
    const sel = document.getElementById(id);
    if (sel && sel.options[0]) sel.options[0].text = t('selectColPH');
  });

  if (document.getElementById('step2').classList.contains('visible')) updateFormulaInfo();
  if (document.getElementById('step3').classList.contains('visible')) renderStep3();
}

function toggleLang() {
  lang = lang === 'ar' ? 'en' : 'ar';
  applyLang();
}


const readAsArrayBuffer = f => new Promise((res, rej) => {
  const r = new FileReader();
  r.onload = e => res(e.target.result);
  r.onerror = rej;
  r.readAsArrayBuffer(f);
});

const readAsText = (f, enc) => new Promise((res, rej) => {
  const r = new FileReader();
  r.onload = e => res(e.target.result);
  r.onerror = rej;
  r.readAsText(f, enc);
});

// Strip zero-width / non-breaking Unicode spaces, trim, remove trailing .0
const normalizeNum = v =>
  String(v == null ? '' : v)
    .replace(/[​-‍﻿ ]/g, '')
    .trim()
    .replace(/\.0+$/, '');

// Round a grade value according to the chosen mode.
// Supported modes: 'round' (default), 'ceil', 'floor', 'none'
function roundGrade(value, mode) {
  switch (mode) {
    case 'ceil':  return Math.ceil(value);
    case 'floor': return Math.floor(value);
    case 'none':  return value;
    default:      return Math.round(value);
  }
}


['bb', 'ug'].forEach(id => {
  const zone   = document.getElementById(id + 'Zone');
  const input  = document.getElementById(id + 'File');
  const nameEl = document.getElementById(id + 'FileName');

  input.addEventListener('change', () => {
    if (input.files[0]) {
      nameEl.textContent = input.files[0].name;
      zone.classList.add('has-file');
      checkBothUploaded();
    }
  });

  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault(); zone.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f) {
      const dt = new DataTransfer(); dt.items.add(f);
      input.files = dt.files;
      nameEl.textContent = f.name;
      zone.classList.add('has-file');
      checkBothUploaded();
    }
  });
});

function checkBothUploaded() {
  document.getElementById('btnParse').disabled =
    !(document.getElementById('bbFile').files[0] && document.getElementById('ugFile').files[0]);
}


function toggleFinalSource() {
  const source   = document.querySelector('input[name="finalSource"]:checked').value;
  const hasFinal = source !== 'none';

  document.getElementById('finalWeightBox').style.display    = hasFinal ? '' : 'none';
  document.getElementById('finalPanel').style.display        = source === 'bb'   ? '' : 'none';
  document.getElementById('finalFileSection').style.display  = source === 'file' ? '' : 'none';

  if (!hasFinal) {
    ['midtermCols', 'extraCreditCols'].forEach(listId => {
      document.querySelectorAll(`#${listId} .col-check-item`).forEach(el => el.classList.remove('disabled'));
      document.querySelectorAll(`#${listId} input[type=checkbox]`).forEach(cb => cb.disabled = false);
    });
  }

  syncPanels();
  updateFormulaInfo();
}

function toggleExtraCredit() {
  const hasExtra = document.getElementById('hasExtraCredit').checked;
  document.getElementById('extraCreditSection').style.display = hasExtra ? '' : 'none';
  document.getElementById('extraCreditPanel').style.display   = hasExtra ? '' : 'none';
  syncPanels();
  updateFormulaInfo();
}

function updateFormulaInfo() {
  const hasFinal  = document.querySelector('input[name="finalSource"]:checked').value !== 'none';
  const hasExtra  = document.getElementById('hasExtraCredit').checked;
  const mw = document.getElementById('midtermWeight').value || 60;
  const fw = document.getElementById('finalWeight').value   || 40;

  let html = t('formulaMid', mw) + '<br>';

  if (hasFinal) {
    html += t('formulaFin', fw) + '<br>';
  } else {
    html += t('formulaNoFin') + '<br>';
  }

  if (hasExtra) html += t('formulaExtra') + '<br>';

  html += t('formulaTotal', hasFinal, hasExtra);
  document.getElementById('formulaInfo').innerHTML = html;
}

function detectGradeColumns(headers) {
  const firstGradeIdx = headers.findIndex(h =>
    h.includes('النقاط') || h.includes('[') || /\|\d+$/.test(h)
  );
  const startIdx = firstGradeIdx > 0 ? firstGradeIdx : 3;
  return headers.slice(startIdx);
}


function isBbLike(headers) {
  return headers.some(h => String(h).includes('النقاط') || String(h).includes('[') || /\|\d+$/.test(String(h)));
}
function isAcademiaLike(rows) {
  return rows.some(row => row.some(c => String(c).includes('رقم الطالب')));
}

async function parseFiles() {
  const bbFile = document.getElementById('bbFile').files[0];
  const ugFile = document.getElementById('ugFile').files[0];

  let bbHeaders = [];

  try {
    const text  = await readAsText(bbFile, 'UTF-16LE');
    const lines = text.split('\n').filter(l => l.trim());
    const parseLine = l => l.split('\t').map(c => c.replace(/^"|"$/g, '').trim());
    const headers = parseLine(lines[0]);
    bbHeaders = headers;

    bbData = lines.slice(1).map(line => {
      const vals = parseLine(line);
      const row  = {};
      headers.forEach((h, i) => row[h] = vals[i] || '');
      return row;
    }).filter(r => r[headers[0]]);

    bbColumns = detectGradeColumns(headers);
  } catch {
    const buf     = await readAsArrayBuffer(bbFile);
    const wb      = XLSX.read(buf, { type: 'array' });
    const ws      = wb.Sheets[wb.SheetNames[0]];
    const rows    = XLSX.utils.sheet_to_json(ws, { header: 1 });
    const headers = (rows[0] || []).map(String);
    bbHeaders = headers;

    bbData = rows.slice(1).map(vals => {
      const row = {};
      headers.forEach((h, i) => row[h] = vals[i] !== undefined ? String(vals[i]) : '');
      return row;
    }).filter(r => r[headers[0]]);

    bbColumns = detectGradeColumns(headers);
  }

  const ugBuf   = await readAsArrayBuffer(ugFile);
  ugWorkbook    = XLSX.read(ugBuf, { type: 'array', cellStyles: true });
  const ugSheet = ugWorkbook.Sheets[ugWorkbook.SheetNames[0]];
  const allRows = XLSX.utils.sheet_to_json(ugSheet, { header: 1, defval: '' });
  ugHeaderRow   = allRows.findIndex(row => row.some(c => String(c).includes('رقم الطالب')));

  // ── File validation ──────────────────────────────────────────────────────────
  const ugLooksLikeBb       = isBbLike(allRows[0] || []);
  const bbLooksLikeAcademia = isAcademiaLike([bbHeaders]);
  const hasUsername         = bbData.length > 0 && Object.keys(bbData[0]).some(k =>
    k.includes('اسم المستخدم') || k.toLowerCase().includes('username'));

  if (ugLooksLikeBb && bbLooksLikeAcademia) {
    alert(t('errFilesSwapped'));
    bbData = null; ugWorkbook = null; return;
  }
  if (ugLooksLikeBb) {
    alert(t('errUgIsBb'));
    bbData = null; ugWorkbook = null; return;
  }
  if (bbLooksLikeAcademia) {
    alert(t('errBbIsUg'));
    bbData = null; ugWorkbook = null; return;
  }
  if (!hasUsername || bbColumns.length === 0) {
    alert(t('errBbInvalid'));
    bbData = null; ugWorkbook = null; return;
  }
  if (ugHeaderRow === -1) {
    alert(t('errUgInvalid'));
    bbData = null; ugWorkbook = null; return;
  }

  ['midtermCols', 'finalCols', 'extraCreditCols'].forEach(id => {
    document.getElementById(id).innerHTML = '';
  });

  bbColumns.forEach((col, idx) => {
    if (!col) return;
    // Support both Arabic format "النقاط: 20" and English format "Total Pts: 20"
    // Also handles "Total Pts: up to 35" (capture group 1 = Arabic, group 2 = English)
    const m      = col.match(/النقاط:\s*([\d.]+)|Total Pts:\s*(?:up to\s*)?([\d.]+)/i);
    const maxPts = m ? parseFloat(m[1] ?? m[2]) : null;
    const short  = col.length > 65 ? col.substring(0, 62) + '...' : col;
    const maxLabel = maxPts !== null ? `<span class="col-max">${t('outOf', maxPts)}</span>` : '';

    const makeItem = (prefix, handler) => {
      const item = document.createElement('div');
      item.className   = 'col-check-item';
      item.dataset.col = col;
      item.innerHTML   = `
        <input type="checkbox" id="${prefix}_${idx}" data-col="${col}" data-max="${maxPts || 0}" onchange="${handler}()" />
        <label for="${prefix}_${idx}">${short}</label>${maxLabel}`;
      return item;
    };

    document.getElementById('midtermCols').appendChild(makeItem('m', 'onMidChange'));
    document.getElementById('finalCols').appendChild(makeItem('f', 'onFinChange'));
    document.getElementById('extraCreditCols').appendChild(makeItem('e', 'onExtChange'));
  });

  document.getElementById('step2').classList.add('visible');
  document.getElementById('step2').scrollIntoView({ behavior: 'smooth', block: 'start' });
  toggleFinalSource();
  updateFormulaInfo();
}


async function parseFinalFile() {
  const file = document.getElementById('finalExamFile').files[0];
  if (!file) return;

  document.getElementById('finalExamFileName').textContent = file.name;
  document.getElementById('finalFileZone').classList.add('has-file');

  let headers = [], rows = [];

  // .xlsx files are ZIP-based binaries — readAsText succeeds but returns garbage,
  // so we must route them directly to the XLSX binary parser.
  const isBinary = /\.xlsx$/i.test(file.name);

  if (isBinary) {
    // Binary XLSX path (also handles real .xls binary workbooks)
    const buf = await readAsArrayBuffer(file);
    const wb  = XLSX.read(buf, { type: 'array' });
    const ws  = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    headers   = (raw[0] || []).map(String);
    rows = raw.slice(1).map(vals => {
      const row = {};
      headers.forEach((h, i) => row[h] = vals[i] !== undefined ? String(vals[i]) : '');
      return row;
    }).filter(r => r[headers[0]]);
  } else {
    // Text path: Blackboard .xls exports are tab-separated UTF-16LE text, not real binary
    try {
      const text  = await readAsText(file, 'UTF-16LE');
      const lines = text.split('\n').filter(l => l.trim());
      const parseLine = l => l.split('\t').map(c => c.replace(/^"|"$/g, '').trim());
      headers = parseLine(lines[0]);
      rows = lines.slice(1).map(line => {
        const vals = parseLine(line);
        const row  = {};
        headers.forEach((h, i) => row[h] = vals[i] || '');
        return row;
      }).filter(r => r[headers[0]]);
    } catch {
      // Fallback to binary XLSX parser for any other format
      const buf = await readAsArrayBuffer(file);
      const wb  = XLSX.read(buf, { type: 'array' });
      const ws  = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      headers   = (raw[0] || []).map(String);
      rows = raw.slice(1).map(vals => {
        const row = {};
        headers.forEach((h, i) => row[h] = vals[i] !== undefined ? String(vals[i]) : '');
        return row;
      }).filter(r => r[headers[0]]);
    }
  }

  finalFileData    = rows;
  finalFileHeaders = headers;

  // Populate column selectors
  ['finalFileStudentCol', 'finalFileGradeCol'].forEach(selId => {
    const sel = document.getElementById(selId);
    sel.innerHTML = `<option value="">${t('selectColPH')}</option>`;
    headers.forEach(h => sel.add(new Option(h, h)));
  });

  document.getElementById('finalFileColSelectors').style.display = '';
}


function onFinChange() { syncPanels(); autoSumMax('fin', 'finalMax'); }
function onExtChange() { syncPanels(); }

function onMidChange() {
  const source    = document.querySelector('input[name="finalSource"]:checked')?.value;
  const hasFinal  = source !== 'none';
  const hasExtra  = document.getElementById('hasExtraCredit').checked;

  const checkedMid = new Set();
  document.querySelectorAll('#midtermCols input[type=checkbox]:checked').forEach(cb => checkedMid.add(cb.dataset.col));

  if (hasFinal && source === 'bb') {
    document.querySelectorAll('#finalCols .col-check-item').forEach(item => {
      const cb = item.querySelector('input');
      if (checkedMid.has(item.dataset.col)) { cb.checked = false; cb.disabled = true; item.classList.add('disabled'); }
      else { cb.disabled = false; item.classList.remove('disabled'); }
    });
    autoSumMax('fin', 'finalMax');
  }

  if (hasExtra) {
    document.querySelectorAll('#extraCreditCols .col-check-item').forEach(item => {
      const cb = item.querySelector('input');
      if (checkedMid.has(item.dataset.col)) { cb.checked = false; cb.disabled = true; item.classList.add('disabled'); }
      else { cb.disabled = false; item.classList.remove('disabled'); }
    });
  }

  autoSumMax('mid', 'midtermMax');
}

function syncPanels() {
  const source    = document.querySelector('input[name="finalSource"]:checked')?.value;
  const hasFinal  = source !== 'none';
  const hasExtra  = document.getElementById('hasExtraCredit').checked;

  // Collect all checked fin & extra cols
  const checkedFin = new Set();
  const checkedExt = new Set();
  if (hasFinal && source === 'bb')
    document.querySelectorAll('#finalCols input[type=checkbox]:checked').forEach(cb => checkedFin.add(cb.dataset.col));
  if (hasExtra)
    document.querySelectorAll('#extraCreditCols input[type=checkbox]:checked').forEach(cb => checkedExt.add(cb.dataset.col));

  const blocked = new Set([...checkedFin, ...checkedExt]);

  document.querySelectorAll('#midtermCols .col-check-item').forEach(item => {
    const cb = item.querySelector('input');
    if (blocked.has(item.dataset.col)) { cb.checked = false; cb.disabled = true; item.classList.add('disabled'); }
    else { cb.disabled = false; item.classList.remove('disabled'); }
  });

  const checkedMid = new Set();
  document.querySelectorAll('#midtermCols input[type=checkbox]:checked').forEach(cb => checkedMid.add(cb.dataset.col));

  if (hasFinal && source === 'bb') {
    document.querySelectorAll('#finalCols .col-check-item').forEach(item => {
      const cb = item.querySelector('input');
      const blocked2 = checkedMid.has(item.dataset.col) || checkedExt.has(item.dataset.col);
      if (blocked2) { cb.checked = false; cb.disabled = true; item.classList.add('disabled'); }
      else { cb.disabled = false; item.classList.remove('disabled'); }
    });
  }

  if (hasExtra) {
    document.querySelectorAll('#extraCreditCols .col-check-item').forEach(item => {
      const cb = item.querySelector('input');
      const blocked2 = checkedMid.has(item.dataset.col) || checkedFin.has(item.dataset.col);
      if (blocked2) { cb.checked = false; cb.disabled = true; item.classList.add('disabled'); }
      else { cb.disabled = false; item.classList.remove('disabled'); }
    });
  }

  autoSumMax('mid', 'midtermMax');
  if (hasFinal && source === 'bb') autoSumMax('fin', 'finalMax');
}

function autoSumMax(panel, inputId) {
  const map = { mid: 'midtermCols', fin: 'finalCols', ext: 'extraCreditCols' };
  let total = 0;
  document.querySelectorAll(`#${map[panel]} input[type=checkbox]:checked`).forEach(cb => {
    total += parseFloat(cb.dataset.max) || 0;
  });
  if (total > 0) document.getElementById(inputId).value = total;
}

function selectAll(panel) {
  const map = { mid: 'midtermCols', fin: 'finalCols', ext: 'extraCreditCols' };
  document.querySelectorAll(`#${map[panel]} input[type=checkbox]:not(:disabled)`).forEach(cb => cb.checked = true);
  if (panel === 'mid') onMidChange();
  else if (panel === 'fin') onFinChange();
  else onExtChange();
}

function deselectAll(panel) {
  const map = { mid: 'midtermCols', fin: 'finalCols', ext: 'extraCreditCols' };
  document.querySelectorAll(`#${map[panel]} input[type=checkbox]`).forEach(cb => cb.checked = false);
  if (panel === 'mid') onMidChange();
  else if (panel === 'fin') onFinChange();
  else onExtChange();
}


function letterGrade(total) {
  if (total >= 95) return 'A+';
  if (total >= 90) return 'A';
  if (total >= 85) return 'B+';
  if (total >= 80) return 'B';
  if (total >= 75) return 'C+';
  if (total >= 70) return 'C';
  if (total >= 65) return 'D+';
  if (total >= 60) return 'D';
  return 'F';
}


function processGrades() {
  const finalSource    = document.querySelector('input[name="finalSource"]:checked')?.value || 'bb';
  const hasFinal       = finalSource !== 'none';
  const hasExtra       = document.getElementById('hasExtraCredit').checked;
  const zeroMissing    = document.getElementById('zeroMissingStudents')?.checked ?? false;
  const roundingMode   = document.getElementById('roundingMode')?.value || 'round';

  const midtermWeight  = parseFloat(document.getElementById('midtermWeight').value) || 60;
  const finalWeight    = parseFloat(document.getElementById('finalWeight').value)   || 40;
  const midMax         = parseFloat(document.getElementById('midtermMax').value);
  const finalMax       = parseFloat(document.getElementById('finalMax').value);
  const extraCap       = parseFloat(document.getElementById('extraCreditCap').value) || Infinity;
  // Max points for a separate final-exam file.
  // If left blank, defaults to finalWeight (assumes file grades are already on the same scale).
  const _finalFileMaxRaw = parseFloat(document.getElementById('finalFileMax')?.value);
  const finalFileMaxVal  = (!isNaN(_finalFileMaxRaw) && _finalFileMaxRaw > 0) ? _finalFileMaxRaw : finalWeight;

  const selectedMidCols = [];
  const selectedFinCols = [];
  const selectedExtCols = [];
  document.querySelectorAll('#midtermCols input[type=checkbox]:checked').forEach(cb => selectedMidCols.push(cb.dataset.col));
  if (hasFinal && finalSource === 'bb')
    document.querySelectorAll('#finalCols input[type=checkbox]:checked').forEach(cb => selectedFinCols.push(cb.dataset.col));
  if (hasExtra)
    document.querySelectorAll('#extraCreditCols input[type=checkbox]:checked').forEach(cb => selectedExtCols.push(cb.dataset.col));

  // ── Validation ────────────────────────────────────────────────────────────
  if (selectedMidCols.length === 0)                        { alert(t('errMidCol')); return; }
  if (isNaN(midMax) || midMax <= 0)                        { alert(t('errMidMax')); return; }
  if (hasFinal && finalSource === 'bb') {
    if (selectedFinCols.length === 0)                      { alert(t('errFinCol')); return; }
    if (isNaN(finalMax) || finalMax <= 0)                  { alert(t('errFinMax')); return; }
  }
  if (hasFinal && finalSource === 'file') {
    if (!finalFileData)                                    { alert(t('errFinFile')); return; }
    const sc = document.getElementById('finalFileStudentCol').value;
    const gc = document.getElementById('finalFileGradeCol').value;
    if (!sc || !gc)                                        { alert(t('errFinFileCol')); return; }
    // finalFileMaxVal always has a valid value (defaults to finalWeight if blank)
  }

  // Detect username column in Blackboard
  const usernameKey = Object.keys(bbData[0]).find(k =>
    k.includes('اسم المستخدم') || k.toLowerCase().includes('username')
  );

  // Build separate-file final lookup: normalised student ID → raw grade
  const finalFileLookup = {};
  if (hasFinal && finalSource === 'file' && finalFileData) {
    const sc = document.getElementById('finalFileStudentCol').value;
    const gc = document.getElementById('finalFileGradeCol').value;
    finalFileData.forEach(row => {
      const num = normalizeNum(row[sc]);
      if (!num) return;
      const v = parseFloat(row[gc]);
      if (!isNaN(v)) finalFileLookup[num] = v;
    });
  }

  // Build Blackboard lookup: student ID → { midterm (scaled), final (scaled), extra }
  const bbLookup = {};
  bbData.forEach(row => {
    const num = normalizeNum(row[usernameKey]);
    if (!num) return;

    let midSum = 0, midHas = false;
    selectedMidCols.forEach(col => { const v = parseFloat(row[col]); if (!isNaN(v)) { midSum += v; midHas = true; } });

    let finSum = 0, finHas = false;
    if (hasFinal && finalSource === 'bb') {
      selectedFinCols.forEach(col => { const v = parseFloat(row[col]); if (!isNaN(v)) { finSum += v; finHas = true; } });
    }

    let extSum = 0;
    if (hasExtra) {
      selectedExtCols.forEach(col => { const v = parseFloat(row[col]); if (!isNaN(v)) extSum += v; });
    }

    // Scale: (student_score / bb_max) * university_weight
    bbLookup[num] = {
      midterm: midHas ? (midSum / midMax) * midtermWeight : null,
      final:   (hasFinal && finalSource === 'bb' && finHas) ? (finSum / finalMax) * finalWeight : null,
      extra:   hasExtra ? Math.min(extSum, extraCap) : 0,
    };
  });

  // ── Info messages: BB max differs from university weight (non-blocking) ────
  const infoMsgs = [];
  if (!isNaN(midMax) && midMax !== midtermWeight)
    infoMsgs.push(t('infoMaxMid', midMax, midtermWeight));
  if (hasFinal && finalSource === 'bb' && !isNaN(finalMax) && finalMax !== finalWeight)
    infoMsgs.push(t('infoMaxFin', finalMax, finalWeight));

  // ── Load university grade-sheet rows ──────────────────────────────────────
  const sheetName = ugWorkbook.SheetNames[0];
  const srcSheet  = ugWorkbook.Sheets[sheetName];
  const allRows   = XLSX.utils.sheet_to_json(srcSheet, { header: 1, defval: '' });
  const cleanRows = allRows.slice(ugHeaderRow);

  const hdr           = cleanRows[0];
  const colStudentNum = hdr.findIndex(c => String(c).includes('رقم الطالب'));
  const colMidterm    = hdr.findIndex(c => String(c).includes('فصلي'));
  const colFinal      = hdr.findIndex(c => String(c).includes('نهائي'));
  const colTotal      = hdr.findIndex(c => String(c).includes('الدرجة') && !String(c).includes('نهائي') && !String(c).includes('فصلي'));
  const colGrade      = hdr.findIndex(c => String(c).includes('التقدير'));

  let matched = 0, partial = 0, zeroed = 0, totalStudents = 0;
  // Track students missing from the separate final-exam file
  let finalFileMatched = 0;
  const finalFileMissedStudents = [];
  const tableRows = [];

  for (let i = 1; i < cleanRows.length; i++) {
    const row        = cleanRows[i];
    const studentNum = normalizeNum(row[colStudentNum]);
    if (!studentNum) continue;
    totalStudents++;

    const studentName = String(row[1] || '');
    const isExcused   = String(row[colGrade] || '').trim() === 'ع';

    if (isExcused) {
      if (colMidterm >= 0)            cleanRows[i][colMidterm] = 0;
      if (hasFinal && colFinal >= 0)  cleanRows[i][colFinal]   = 0;
      if (colTotal >= 0)              cleanRows[i][colTotal]   = 0;
      if (colGrade >= 0)              cleanRows[i][colGrade]   = 0;
      tableRows.push({ num: studentNum, name: studentName, mid: 0, fin: hasFinal ? 0 : '—', extra: hasExtra ? 0 : '—', total: 0, grade: 'F', status: 'excused' });
      continue;
    }

    const grades = bbLookup[studentNum];

    if (!grades) {
      // Student is in the university file but not in Blackboard
      zeroed++;
      if (zeroMissing) {
        if (colMidterm >= 0)            cleanRows[i][colMidterm] = 0;
        if (hasFinal && colFinal >= 0)  cleanRows[i][colFinal]   = 0;
        if (colTotal >= 0)              cleanRows[i][colTotal]   = 0;
        if (colGrade >= 0)              cleanRows[i][colGrade]   = 'F';
        tableRows.push({ num: studentNum, name: studentName, mid: 0, fin: hasFinal ? 0 : '—', extra: hasExtra ? 0 : '—', total: 0, grade: 'F', status: 'missing' });
      } else {
        // Leave original university-file values unchanged; flag as missing only
        const origMid   = colMidterm >= 0 ? row[colMidterm] : '—';
        const origFin   = colFinal   >= 0 ? row[colFinal]   : '—';
        const origTotal = colTotal   >= 0 ? row[colTotal]   : '—';
        const origGrade = colGrade   >= 0 ? row[colGrade]   : '—';
        tableRows.push({ num: studentNum, name: studentName, mid: origMid, fin: hasFinal ? origFin : '—', extra: '—', total: origTotal, grade: origGrade, status: 'missing' });
      }
      continue;
    }

    // Scale midterm: (score / bbMax) * uniWeight, then apply chosen rounding
    const midVal = grades.midterm !== null
      ? Math.min(midtermWeight, roundGrade(grades.midterm, roundingMode))
      : 0;

    // Final grade: from Blackboard columns or from a separate file
    let finVal = 0;
    // 'partial' = found in BB but final-file grade is missing → not fully updated
    let studentStatus = 'ok';

    if (hasFinal) {
      if (finalSource === 'bb') {
        finVal = grades.final !== null
          ? Math.min(finalWeight, roundGrade(grades.final, roundingMode))
          : 0;
      } else {
        // Scale the raw score from the separate file: (raw / fileMax) * uniWeight
        const rawFin = finalFileLookup[studentNum];
        if (rawFin !== undefined) {
          finalFileMatched++;
          finVal = Math.min(finalWeight, roundGrade((rawFin / finalFileMaxVal) * finalWeight, roundingMode));
        } else {
          // In BB but no entry in the final-exam file → partial record
          finalFileMissedStudents.push(studentNum);
          studentStatus = 'partial';
          finVal = 0;
        }
      }
    }

    // Count fully-updated vs partial separately so the summary is accurate
    if (studentStatus === 'ok') matched++;
    else                        partial++;

    const extraVal = hasExtra ? grades.extra : 0;
    const total    = Math.min(100, roundGrade(midVal + (hasFinal ? finVal : 0) + extraVal, roundingMode));

    if (colMidterm >= 0)            cleanRows[i][colMidterm] = midVal;
    if (hasFinal && colFinal >= 0)  cleanRows[i][colFinal]   = finVal;
    if (colTotal >= 0)              cleanRows[i][colTotal]   = total;
    const lg = letterGrade(total);
    if (colGrade >= 0)              cleanRows[i][colGrade]   = lg;

    tableRows.push({
      num: studentNum, name: studentName,
      mid: midVal,
      fin: hasFinal ? finVal : '—',
      extra: hasExtra ? extraVal : '—',
      total, grade: lg, status: studentStatus
    });
  }

  // ── Final-file match check ─────────────────────────────────────────────────
  // If every eligible student is missing from the final file, something is wrong
  // (wrong column selected, wrong file, or ID format mismatch). Block and warn.
  if (hasFinal && finalSource === 'file' && matched > 0 && finalFileMatched === 0) {
    alert(t('errFinFileNoMatch'));
    return;
  }

  // ── Output workbook ────────────────────────────────────────────────────────
  const keepCols = [colStudentNum, 1, colMidterm];
  if (hasFinal && colFinal >= 0) keepCols.push(colFinal);

  const dataRows = cleanRows.slice(1).map(row => keepCols.map(c => row[c] !== undefined ? row[c] : ''));

  // Custom header with grade-distribution percentages
  const midPct = Math.round(midtermWeight);
  const finPct = Math.round(finalWeight);
  const header = ['رقم الطالب', 'اسم الطالب', `فصلي (${midPct}%)`];
  if (hasFinal) header.push(`نهائي (${finPct}%)`);

  const outputRows = [header, ...dataRows];

  // Map student ID → output row index (for in-place midterm edits)
  const numToOutputRow = {};
  for (let i = 1; i < outputRows.length; i++) {
    const num = normalizeNum(outputRows[i][0]);
    if (num) numToOutputRow[num] = i;
  }

  const newSheet = XLSX.utils.aoa_to_sheet(outputRows);
  resultWorkbook  = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(resultWorkbook, newSheet, sheetName);

  // ── Save results state and render ─────────────────────────────────────────
  const missingStudents  = tableRows.filter(r => r.status === 'missing').map(r => r.num);
  const midtermWeightVal = parseFloat(document.getElementById('midtermWeight').value) || 60;
  lastResults = {
    totalStudents, matched, partial, zeroed, missingStudents, tableRows, hasFinal, hasExtra,
    outputRows, numToOutputRow, sheetName, midtermWeight: midtermWeightVal,
    infoMsgs,
    finalFileMissedStudents,  // students found in BB but missing from the final file
  };

  renderStep3();
  document.getElementById('step3').classList.add('visible');
  document.getElementById('step3').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderStep3() {
  if (!lastResults) return;
  const { totalStudents, matched, partial, zeroed, missingStudents, tableRows, hasFinal, hasExtra,
          infoMsgs, finalFileMissedStudents } = lastResults;
  const excusedCount = tableRows.filter(r => r.status === 'excused').length;

  document.getElementById('noteBox').textContent = t('noteCheck');

  // Non-blocking info: BB max differs from university weight
  const maxInfoEl = document.getElementById('maxInfoAlert');
  if (maxInfoEl) {
    maxInfoEl.innerHTML = (infoMsgs && infoMsgs.length > 0)
      ? `<div class="alert alert-info">${infoMsgs.map(m => `<div>${m}</div>`).join('')}</div>`
      : '';
  }

  // Build summary — partial box only appears when a separate final file is used
  let summaryHtml = `
    <div class="summary-box">
      <div class="num">${totalStudents}</div>
      <div class="lbl">${t('totalStudents')}</div>
    </div>
    <div class="summary-box">
      <div class="num" style="color:#16a34a">${matched}</div>
      <div class="lbl">${t('updated')}</div>
    </div>`;

  if (partial > 0) {
    summaryHtml += `
    <div class="summary-box">
      <div class="num" style="color:#d97706">${partial}</div>
      <div class="lbl">${t('partialUpdate')}</div>
    </div>`;
  }

  summaryHtml += `
    <div class="summary-box">
      <div class="num" style="color:#dc2626">${zeroed}</div>
      <div class="lbl">${t('notFound')}</div>
    </div>`;

  if (excusedCount > 0) {
    summaryHtml += `
    <div class="summary-box">
      <div class="num" style="color:#6b7280">${excusedCount}</div>
      <div class="lbl">${t('excusedCount')}</div>
    </div>`;
  }

  document.getElementById('summaryGrid').innerHTML = summaryHtml;

  // Build the alerts block
  let alertsHtml = '';

  // Students missing from Blackboard
  if (missingStudents.length > 0) {
    alertsHtml += `<div class="alert alert-warning">
      <strong>${t('warnTitle')}:</strong> ${t('warnMsg')}
      <ul class="missing-list">${missingStudents.map(s => `<li>${s}</li>`).join('')}</ul>
    </div>`;
  } else {
    alertsHtml += `<div class="alert alert-info">${t('successMsg')}</div>`;
  }

  // Students found in Blackboard but missing from the separate final-exam file
  if (finalFileMissedStudents && finalFileMissedStudents.length > 0) {
    alertsHtml += `<div class="alert alert-warning">
      <strong>${t('warnTitle')}:</strong> ${t('warnFinFileMissed', finalFileMissedStudents.length)}
      <ul class="missing-list">${finalFileMissedStudents.map(s => `<li>${s}</li>`).join('')}</ul>
    </div>`;
  }

  document.getElementById('missingAlert').innerHTML = alertsHtml;

  const finTh   = hasFinal ? `<th class="col-num">${t('tblFin')}</th>`   : '';
  const extraTh = hasExtra ? `<th class="col-num">${t('tblExtra')}</th>` : '';
  const finFilterTd   = hasFinal ? `<th></th>` : '';
  const extraFilterTd = hasExtra ? `<th></th>` : '';
  const colSpan = 3 + (hasFinal ? 1 : 0) + (hasExtra ? 1 : 0) + 2;
  document.getElementById('gradeTableHead').innerHTML = `
    <tr>
      <th class="col-id">${t('tblNum')}</th>
      <th class="col-name">${t('tblName')}</th>
      <th class="col-num">${t('tblMid')}</th>
      ${finTh}
      ${extraTh}
      <th class="col-num">${t('tblTotal')}</th>
      <th class="col-letter">${t('tblGrade')}</th>
      <th class="col-status">${t('tblStatus')}</th>
    </tr>
    <tr class="filter-row">
      <th class="col-id"><input type="text" id="filterSearch" class="col-filter-input" placeholder="${t('filterSearchPH')}" oninput="filterTable()" /></th>
      <th class="col-name"></th>
      <th></th>
      ${finFilterTd}
      ${extraFilterTd}
      <th></th>
      <th class="col-letter">
        <select id="filterGrade" class="col-filter-select" onchange="filterTable()">
          <option value="">${t('filterGradeAll')}</option>
          <option value="A">A</option>
          <option value="B+">B+</option>
          <option value="B">B</option>
          <option value="C+">C+</option>
          <option value="C">C</option>
          <option value="D+">D+</option>
          <option value="D">D</option>
          <option value="F">F</option>
        </select>
      </th>
      <th class="col-status">
        <select id="filterStatus" class="col-filter-select" onchange="filterTable()">
          <option value="">${t('filterStatusAll')}</option>
          <option value="ok">${t('statusOk')}</option>
          <option value="partial">${t('statusPartial')}</option>
          <option value="missing">${t('statusMissing')}</option>
          <option value="excused">${t('statusExcused')}</option>
        </select>
      </th>
    </tr>`;

  renderTableRows(tableRows);
}

function rowHighlightClass(r) {
  if (r.status === 'excused' || typeof r.total !== 'number') return '';
  const total = r.total;
  if (total >= 55 && total <= 59) return 'hl-fail-high';
  if (total >= 50 && total <= 54) return 'hl-fail-low';
  if (total >= 60) {
    const ones = Math.floor(total) % 10;
    if (ones === 4 || ones === 9) return 'hl-borderline';
  }
  return '';
}

function renderTableRows(rows) {
  if (!lastResults) return;
  const { hasFinal, hasExtra, midtermWeight } = lastResults;
  document.getElementById('gradeTableBody').innerHTML = rows.map((r, idx) => {
    const rowIdx  = lastResults.tableRows.indexOf(r);
    // fin cell: highlight zero when data is genuinely missing (not just a zero score)
    const finMissing = r.status === 'partial';
    const finTd   = hasFinal ? `<td class="col-num ${r.status === 'excused' ? 'excused' : (finMissing ? 'zero' : '')}">${r.fin}</td>` : '';
    const extraTd = hasExtra ? `<td class="col-num">${r.extra}</td>` : '';
    const isPass     = typeof r.total === 'number' && r.total >= 60 && r.status !== 'excused';
    const colorClass = r.status === 'excused' ? 'excused'
                     : r.status === 'partial'  ? 'partial'
                     : (isPass ? 'pass' : 'fail');
    const statusLbl  = r.status === 'ok'      ? t('statusOk')
                     : r.status === 'partial'  ? t('statusPartial')
                     : r.status === 'missing'  ? t('statusMissing')
                     :                           t('statusExcused');
    const hlClass    = rowHighlightClass(r);
    const midClass   = r.status === 'excused' ? 'excused' : (r.mid === 0 && r.status !== 'ok' && r.status !== 'partial' ? 'zero' : '');
    const canEdit    = r.status !== 'excused' && r.status !== 'missing';
    const editBtn    = canEdit
      ? `<button class="btn-edit-mid" onclick="startEditMid(this, ${rowIdx}, ${midtermWeight})" title="${t('editMidTitle')}">✏️</button>`
      : '';
    return `<tr class="${hlClass}">
      <td class="col-id">${r.num}</td>
      <td class="col-name">${r.name}</td>
      <td class="col-num mid-cell ${midClass}">${r.mid}${editBtn}</td>
      ${finTd}
      ${extraTd}
      <td class="col-num ${r.status === 'excused' ? 'excused' : (r.total === 0 && r.status !== 'ok' ? 'zero' : '')}">${r.total}</td>
      <td class="col-letter"><span class="grade-letter ${colorClass}">${r.grade}</span></td>
      <td class="col-status"><span class="status-chip ${colorClass}">${statusLbl}</span></td>
    </tr>`;
  }).join('');
}

function startEditMid(btn, rowIdx, maxMid) {
  const cell = btn.parentElement;
  const row  = lastResults.tableRows[rowIdx];
  cell.innerHTML = `
    <input type="number" class="mid-edit-input" value="${row.mid}" min="0" max="${maxMid}" step="0.5"
      onkeydown="handleMidEditKey(event, this, ${rowIdx}, ${maxMid})" />
    <button class="btn-edit-confirm" onclick="confirmEditMid(this.previousElementSibling, ${rowIdx}, ${maxMid})">✓</button>
    <button class="btn-edit-cancel"  onclick="cancelEditMid(${rowIdx})">✕</button>`;
  cell.querySelector('input').focus();
}

function handleMidEditKey(e, input, rowIdx, maxMid) {
  if (e.key === 'Enter') confirmEditMid(input, rowIdx, maxMid);
  if (e.key === 'Escape') cancelEditMid(rowIdx);
}

function confirmEditMid(input, rowIdx, maxMid) {
  const newMid = Math.min(maxMid, Math.max(0, parseFloat(input.value) || 0));
  const r = lastResults.tableRows[rowIdx];
  const newTotal = Math.round((newMid + (r.total - r.mid)) * 100) / 100;
  r.mid   = Math.round(newMid * 100) / 100;
  r.total = newTotal;
  r.grade = letterGrade(newTotal);

  // Write normalized midterm value back into outputRows (col 2)
  // outputRows already stores normalized values, so newMid is written directly.
  const { outputRows, numToOutputRow, sheetName } = lastResults;
  const outIdx = numToOutputRow[normalizeNum(String(r.num))];
  if (outIdx !== undefined) {
    outputRows[outIdx][2] = newMid;
    const newSheet = XLSX.utils.aoa_to_sheet(outputRows);
    resultWorkbook.Sheets[sheetName] = newSheet;
  }

  filterTable();
}

function cancelEditMid(rowIdx) {
  filterTable();
}

function filterTable() {
  if (!lastResults) return;
  const search = document.getElementById('filterSearch').value.trim().toLowerCase();
  const grade  = document.getElementById('filterGrade').value;
  const status = document.getElementById('filterStatus').value;
  const filtered = lastResults.tableRows.filter(r => {
    if (search && !String(r.num).toLowerCase().includes(search) && !r.name.toLowerCase().includes(search)) return false;
    if (grade  && r.grade  !== grade)  return false;
    if (status && r.status !== status) return false;
    return true;
  });
  renderTableRows(filtered);
}

// ── Download ──────────────────────────────────────────────────────────────────

function downloadResult() {
  if (!resultWorkbook) return;
  const checkbox = document.getElementById('disclaimerCheck');
  if (!checkbox.checked) {
    alert(t('disclaimerRequired'));
    checkbox.closest('.disclaimer-box').classList.add('disclaimer-error');
    return;
  }
  const filename = lang === 'ar' ? 'كشف_الدرجات_المحدّث.xls' : 'Updated_Grade_Sheet.xls';
  XLSX.writeFile(resultWorkbook, filename, { bookType: 'xls' });
}

function resetAll() {
  // Clear state
  bbData         = null;
  ugWorkbook     = null;
  ugHeaderRow    = -1;
  bbColumns      = [];
  resultWorkbook = null;
  finalFileData  = null;
  finalFileHeaders = [];
  lastResults    = null;

  // Reset file inputs and upload zones
  ['bb', 'ug'].forEach(id => {
    document.getElementById(id + 'File').value = '';
    document.getElementById(id + 'FileName').textContent = '';
    document.getElementById(id + 'Zone').classList.remove('has-file');
  });
  document.getElementById('finalExamFile').value = '';
  document.getElementById('finalExamFileName').textContent = '';
  document.getElementById('finalFileZone').classList.remove('has-file');
  document.getElementById('finalFileColSelectors').style.display = 'none';

  // Reset step 2 controls
  document.querySelector('input[name="finalSource"][value="bb"]').checked = true;
  document.getElementById('hasExtraCredit').checked = false;
  document.getElementById('extraCreditSection').style.display = 'none';
  document.getElementById('midtermWeight').value = 60;
  document.getElementById('finalWeight').value   = 40;
  document.getElementById('midtermMax').value    = '';
  document.getElementById('finalMax').value      = '';
  document.getElementById('extraCreditCap').value = '';
  ['midtermCols', 'finalCols', 'extraCreditCols'].forEach(id => {
    document.getElementById(id).innerHTML = '';
  });

  // Reset disclaimer
  const dc = document.getElementById('disclaimerCheck');
  if (dc) { dc.checked = false; dc.closest('.disclaimer-box').classList.remove('disclaimer-error'); }

  // Hide steps 2 & 3
  document.getElementById('step2').classList.remove('visible');
  document.getElementById('step3').classList.remove('visible');

  // Disable parse button
  document.getElementById('btnParse').disabled = true;

  // Scroll back to top
  document.getElementById('step1').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ── Init ──────────────────────────────────────────────────────────────────────

applyLang();
