
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

const normalizeNum = v => String(v || '').trim().replace(/\.0+$/, '');


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


async function parseFiles() {
  const bbFile = document.getElementById('bbFile').files[0];
  const ugFile = document.getElementById('ugFile').files[0];

  try {
    const text  = await readAsText(bbFile, 'UTF-16LE');
    const lines = text.split('\n').filter(l => l.trim());
    const parseLine = l => l.split('\t').map(c => c.replace(/^"|"$/g, '').trim());
    const headers = parseLine(lines[0]);

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
    const headers = rows[0];

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

  ['midtermCols', 'finalCols', 'extraCreditCols'].forEach(id => {
    document.getElementById(id).innerHTML = '';
  });

  bbColumns.forEach((col, idx) => {
    if (!col) return;
    const m      = col.match(/النقاط:\s*([\d.]+)/);
    const maxPts = m ? parseFloat(m[1]) : null;
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
    const buf = await readAsArrayBuffer(file);
    const wb  = XLSX.read(buf, { type: 'array' });
    const ws  = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    headers   = raw[0].map(String);
    rows = raw.slice(1).map(vals => {
      const row = {};
      headers.forEach((h, i) => row[h] = vals[i] !== undefined ? String(vals[i]) : '');
      return row;
    }).filter(r => r[headers[0]]);
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
  const finalSource = document.querySelector('input[name="finalSource"]:checked')?.value || 'bb';
  const hasFinal    = finalSource !== 'none';
  const hasExtra    = document.getElementById('hasExtraCredit').checked;

  const midtermWeight = parseFloat(document.getElementById('midtermWeight').value) || 60;
  const finalWeight   = parseFloat(document.getElementById('finalWeight').value)   || 40;
  const midMax        = parseFloat(document.getElementById('midtermMax').value);
  const finalMax      = parseFloat(document.getElementById('finalMax').value);
  const extraCap = parseFloat(document.getElementById('extraCreditCap').value) || Infinity;

  const selectedMidCols = [];
  const selectedFinCols = [];
  const selectedExtCols = [];
  document.querySelectorAll('#midtermCols input[type=checkbox]:checked').forEach(cb => selectedMidCols.push(cb.dataset.col));
  if (hasFinal && finalSource === 'bb')
    document.querySelectorAll('#finalCols input[type=checkbox]:checked').forEach(cb => selectedFinCols.push(cb.dataset.col));
  if (hasExtra)
    document.querySelectorAll('#extraCreditCols input[type=checkbox]:checked').forEach(cb => selectedExtCols.push(cb.dataset.col));

  // Validation
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
  }

  // Detect username column in Blackboard
  const usernameKey = Object.keys(bbData[0]).find(k =>
    k.includes('اسم المستخدم') || k.toLowerCase().includes('username')
  );

  // Build separate-file final lookup (student num → raw grade)
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

  // Build Blackboard lookup: student num → { midterm, final (BB), extraCredit }
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

    bbLookup[num] = {
      midterm: midHas ? (midSum / midMax) * midtermWeight : null,
      final:   (hasFinal && finalSource === 'bb' && finHas) ? (finSum / finalMax) * finalWeight : null,
      extra:   hasExtra ? Math.min(extSum, extraCap) : 0,
    };
  });

  // Load university grader rows (strip metadata header rows)
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

  let matched = 0, zeroed = 0, totalStudents = 0;
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
      zeroed++;
      if (colMidterm >= 0)            cleanRows[i][colMidterm] = 0;
      if (hasFinal && colFinal >= 0)  cleanRows[i][colFinal]   = 0;
      if (colTotal >= 0)              cleanRows[i][colTotal]   = 0;
      if (colGrade >= 0)              cleanRows[i][colGrade]   = 'F';
      tableRows.push({ num: studentNum, name: studentName, mid: 0, fin: hasFinal ? 0 : '—', extra: hasExtra ? 0 : '—', total: 0, grade: 'F', status: 'missing' });
      continue;
    }

    matched++;
    const midVal = grades.midterm !== null ? Math.ceil(grades.midterm) : 0;

    // Final: from BB or from separate file
    let finVal = 0;
    if (hasFinal) {
      if (finalSource === 'bb') {
        finVal = grades.final !== null ? Math.ceil(grades.final) : 0;
      } else {
        const rawFin = finalFileLookup[studentNum];
        finVal = rawFin !== undefined ? Math.ceil(rawFin) : 0;
      }
    }

    const extraVal = hasExtra ? grades.extra : 0;
    const total    = Math.ceil(midVal + (hasFinal ? finVal : 0) + extraVal);

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
      total, grade: lg, status: 'ok'
    });
  }

  // Output columns — only student number, name, midterm, final
  const keepCols = [colStudentNum, 1, colMidterm];
  if (hasFinal && colFinal >= 0) keepCols.push(colFinal);

  const dataRows = cleanRows.slice(1).map(row => keepCols.map(c => row[c] !== undefined ? row[c] : ''));

  // Custom header with percentages
  const midPct = Math.round(midtermWeight);
  const finPct = Math.round(finalWeight);
  const header = [
    'رقم الطالب',
    'اسم الطالب',
    `فصلي (${midPct}%)`,
  ];
  if (hasFinal) header.push(`نهائي (${finPct}%)`);

  const outputRows = [header, ...dataRows];

  // Map student number → output row index (skip header row 0)
  const numToOutputRow = {};
  for (let i = 1; i < outputRows.length; i++) {
    const num = normalizeNum(outputRows[i][0]);
    if (num) numToOutputRow[num] = i;
  }

  const newSheet = XLSX.utils.aoa_to_sheet(outputRows);
  resultWorkbook  = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(resultWorkbook, newSheet, sheetName);

  // ── Save results state and render ─────────────────────────────────────────

  const missingStudents = tableRows.filter(r => r.status === 'missing').map(r => r.num);
  lastResults = { totalStudents, matched, zeroed, missingStudents, tableRows, hasFinal, hasExtra,
                  outputRows, numToOutputRow, sheetName };

  renderStep3();
  document.getElementById('step3').classList.add('visible');
  document.getElementById('step3').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderStep3() {
  if (!lastResults) return;
  const { totalStudents, matched, zeroed, missingStudents, tableRows, hasFinal, hasExtra } = lastResults;

  document.getElementById('noteBox').textContent = t('noteCheck');

  document.getElementById('summaryGrid').innerHTML = `
    <div class="summary-box">
      <div class="num">${totalStudents}</div>
      <div class="lbl">${t('totalStudents')}</div>
    </div>
    <div class="summary-box">
      <div class="num" style="color:#16a34a">${matched}</div>
      <div class="lbl">${t('updated')}</div>
    </div>
    <div class="summary-box">
      <div class="num" style="color:#dc2626">${zeroed}</div>
      <div class="lbl">${t('notFound')}</div>
    </div>`;

  document.getElementById('missingAlert').innerHTML = missingStudents.length > 0
    ? `<div class="alert alert-warning">
        <strong>${t('warnTitle')}:</strong> ${t('warnMsg')}
        <ul class="missing-list">${missingStudents.map(s => `<li>${s}</li>`).join('')}</ul>
       </div>`
    : `<div class="alert alert-info">${t('successMsg')}</div>`;

  const finTh   = hasFinal ? `<th class="col-num">${t('tblFin')}</th>`   : '';
  const extraTh = hasExtra ? `<th class="col-num">${t('tblExtra')}</th>` : '';
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
    </tr>`;

  document.getElementById('gradeTableBody').innerHTML = tableRows.map(r => {
    const finTd   = hasFinal ? `<td class="col-num ${r.status === 'excused' ? 'excused' : (r.fin === 0 && r.status !== 'ok' ? 'zero' : '')}">${r.fin}</td>` : '';
    const extraTd = hasExtra ? `<td class="col-num">${r.extra}</td>` : '';
    const isPass     = typeof r.total === 'number' && r.total >= 60 && r.status !== 'excused';
    const colorClass = r.status === 'excused' ? 'excused' : (isPass ? 'pass' : 'fail');
    const statusLbl  = r.status === 'ok' ? t('statusOk') : r.status === 'missing' ? t('statusMissing') : t('statusExcused');
    return `<tr>
      <td class="col-id">${r.num}</td>
      <td class="col-name">${r.name}</td>
      <td class="col-num ${r.status === 'excused' ? 'excused' : (r.mid === 0 && r.status !== 'ok' ? 'zero' : '')}">${r.mid}</td>
      ${finTd}
      ${extraTd}
      <td class="col-num ${r.status === 'excused' ? 'excused' : (r.total === 0 && r.status !== 'ok' ? 'zero' : '')}">${r.total}</td>
      <td class="col-letter"><span class="grade-letter ${colorClass}">${r.grade}</span></td>
      <td class="col-status"><span class="status-chip ${colorClass}">${statusLbl}</span></td>
    </tr>`;
  }).join('');
}

// ── Download ──────────────────────────────────────────────────────────────────

function downloadResult() {
  if (!resultWorkbook) return;
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
