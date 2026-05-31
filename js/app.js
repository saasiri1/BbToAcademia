/**
 * app.js
 * Main application logic for the Blackboard Grade Converter.
 */

// ── State ─────────────────────────────────────────────────────────────────────

let lang          = 'ar';
let bbData        = null;
let ugWorkbook    = null;
let ugHeaderRow   = -1;
let bbColumns     = [];
let resultWorkbook = null;

// ── i18n ──────────────────────────────────────────────────────────────────────

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

  if (document.getElementById('step2').classList.contains('visible')) {
    updateFormulaInfo();
  }

  const noteBox = document.getElementById('noteBox');
  if (noteBox) noteBox.textContent = t('noteCheck');
}

function toggleLang() {
  lang = lang === 'ar' ? 'en' : 'ar';
  applyLang();
}

// ── Helpers ───────────────────────────────────────────────────────────────────

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

// ── Upload Zones ──────────────────────────────────────────────────────────────

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

  zone.addEventListener('dragover', e => {
    e.preventDefault();
    zone.classList.add('drag-over');
  });

  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));

  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const f = e.dataTransfer.files[0];
    if (f) {
      const dt = new DataTransfer();
      dt.items.add(f);
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

// ── Final Section Toggle ──────────────────────────────────────────────────────

function toggleFinalSection() {
  const hasFinal = document.querySelector('input[name="hasFinal"]:checked').value === 'yes';
  document.getElementById('finalPanel').style.display     = hasFinal ? '' : 'none';
  document.getElementById('finalWeightBox').style.display = hasFinal ? '' : 'none';

  if (!hasFinal) {
    document.querySelectorAll('#midtermCols .col-check-item').forEach(el => el.classList.remove('disabled'));
    document.querySelectorAll('#midtermCols input[type=checkbox]').forEach(cb => cb.disabled = false);
  }

  syncPanels();
  updateFormulaInfo();
}

function updateFormulaInfo() {
  const hasFinal = document.querySelector('input[name="hasFinal"]:checked').value === 'yes';
  const mw = document.getElementById('midtermWeight').value || 60;
  const fw = document.getElementById('finalWeight').value   || 40;

  document.getElementById('formulaInfo').innerHTML =
    t('formulaMid', mw) + '<br>' +
    (hasFinal ? t('formulaFin', fw) : t('formulaNoFin')) + '<br>' +
    t('formulaTotal', hasFinal);
}

// ── Parse Files ───────────────────────────────────────────────────────────────

async function parseFiles() {
  const bbFile = document.getElementById('bbFile').files[0];
  const ugFile = document.getElementById('ugFile').files[0];

  // Parse Blackboard — UTF-16 TSV (Blackboard export format) with xlsx fallback
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

    bbColumns = headers.slice(6);
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

    bbColumns = headers.slice(6);
  }

  // Parse university grader
  const ugBuf   = await readAsArrayBuffer(ugFile);
  ugWorkbook    = XLSX.read(ugBuf, { type: 'array', cellStyles: true });
  const ugSheet = ugWorkbook.Sheets[ugWorkbook.SheetNames[0]];
  const allRows = XLSX.utils.sheet_to_json(ugSheet, { header: 1, defval: '' });
  ugHeaderRow   = allRows.findIndex(row => row.some(c => String(c).includes('رقم الطالب')));

  // Build column checklists
  const midList = document.getElementById('midtermCols');
  const finList = document.getElementById('finalCols');
  midList.innerHTML = '';
  finList.innerHTML = '';

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

    midList.appendChild(makeItem('m', 'onMidChange'));
    finList.appendChild(makeItem('f', 'onFinChange'));
  });

  document.getElementById('step2').classList.add('visible');
  document.getElementById('step2').scrollIntoView({ behavior: 'smooth', block: 'start' });
  updateFormulaInfo();
}

// ── Panel Sync ────────────────────────────────────────────────────────────────

function onFinChange() {
  syncPanels();
  autoSumMax('fin', 'finalMax');
}

function onMidChange() {
  const hasFinal = document.querySelector('input[name="hasFinal"]:checked').value === 'yes';

  if (hasFinal) {
    const checkedMid = new Set();
    document.querySelectorAll('#midtermCols input[type=checkbox]:checked').forEach(cb => checkedMid.add(cb.dataset.col));

    document.querySelectorAll('#finalCols .col-check-item').forEach(item => {
      const cb = item.querySelector('input');
      if (checkedMid.has(item.dataset.col)) {
        cb.checked = false; cb.disabled = true; item.classList.add('disabled');
      } else {
        cb.disabled = false; item.classList.remove('disabled');
      }
    });

    autoSumMax('fin', 'finalMax');
  }

  autoSumMax('mid', 'midtermMax');
}

function syncPanels() {
  const hasFinal = document.querySelector('input[name="hasFinal"]:checked').value === 'yes';
  if (!hasFinal) return;

  const checkedFin = new Set();
  document.querySelectorAll('#finalCols input[type=checkbox]:checked').forEach(cb => checkedFin.add(cb.dataset.col));

  document.querySelectorAll('#midtermCols .col-check-item').forEach(item => {
    const cb = item.querySelector('input');
    if (checkedFin.has(item.dataset.col)) {
      cb.checked = false; cb.disabled = true; item.classList.add('disabled');
    } else {
      cb.disabled = false; item.classList.remove('disabled');
    }
  });

  autoSumMax('mid', 'midtermMax');
}

function autoSumMax(panel, inputId) {
  let total = 0;
  const selector = panel === 'mid' ? '#midtermCols' : '#finalCols';
  document.querySelectorAll(`${selector} input[type=checkbox]:checked`).forEach(cb => {
    total += parseFloat(cb.dataset.max) || 0;
  });
  if (total > 0) document.getElementById(inputId).value = total;
}

function selectAll(panel) {
  const selector = panel === 'mid' ? '#midtermCols' : '#finalCols';
  document.querySelectorAll(`${selector} input[type=checkbox]:not(:disabled)`).forEach(cb => cb.checked = true);
  panel === 'mid' ? onMidChange() : onFinChange();
}

function deselectAll(panel) {
  const selector = panel === 'mid' ? '#midtermCols' : '#finalCols';
  document.querySelectorAll(`${selector} input[type=checkbox]`).forEach(cb => cb.checked = false);
  panel === 'mid' ? onMidChange() : onFinChange();
}

// ── Letter Grade ──────────────────────────────────────────────────────────────

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

// ── Process Grades ────────────────────────────────────────────────────────────

function processGrades() {
  const hasFinal      = document.querySelector('input[name="hasFinal"]:checked').value === 'yes';
  const midtermWeight = parseFloat(document.getElementById('midtermWeight').value) || 60;
  const finalWeight   = parseFloat(document.getElementById('finalWeight').value)   || 40;
  const midMax        = parseFloat(document.getElementById('midtermMax').value);
  const finalMax      = parseFloat(document.getElementById('finalMax').value);

  const selectedMidCols = [];
  const selectedFinCols = [];
  document.querySelectorAll('#midtermCols input[type=checkbox]:checked').forEach(cb => selectedMidCols.push(cb.dataset.col));
  if (hasFinal) document.querySelectorAll('#finalCols input[type=checkbox]:checked').forEach(cb => selectedFinCols.push(cb.dataset.col));

  if (selectedMidCols.length === 0)                   { alert(t('errMidCol')); return; }
  if (isNaN(midMax) || midMax <= 0)                   { alert(t('errMidMax')); return; }
  if (hasFinal && selectedFinCols.length === 0)        { alert(t('errFinCol')); return; }
  if (hasFinal && (isNaN(finalMax) || finalMax <= 0)) { alert(t('errFinMax')); return; }

  // Detect username column
  const usernameKey = Object.keys(bbData[0]).find(k =>
    k.includes('اسم المستخدم') || k.toLowerCase().includes('username')
  );

  // Build student lookup from Blackboard
  const bbLookup = {};
  bbData.forEach(row => {
    const num = normalizeNum(row[usernameKey]);
    if (!num) return;

    let midSum = 0, midHas = false;
    selectedMidCols.forEach(col => {
      const v = parseFloat(row[col]);
      if (!isNaN(v)) { midSum += v; midHas = true; }
    });

    let finSum = 0, finHas = false;
    if (hasFinal) {
      selectedFinCols.forEach(col => {
        const v = parseFloat(row[col]);
        if (!isNaN(v)) { finSum += v; finHas = true; }
      });
    }

    bbLookup[num] = {
      midterm: midHas ? (midSum / midMax) * midtermWeight : null,
      final:   (hasFinal && finHas) ? (finSum / finalMax) * finalWeight : null,
    };
  });

  // Load university grader rows (strip metadata header rows)
  const sheetName = ugWorkbook.SheetNames[0];
  const srcSheet  = ugWorkbook.Sheets[sheetName];
  const allRows   = XLSX.utils.sheet_to_json(srcSheet, { header: 1, defval: '' });
  const cleanRows = allRows.slice(ugHeaderRow);

  // Locate columns by header name
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
      if (colMidterm >= 0)             cleanRows[i][colMidterm] = 0;
      if (hasFinal && colFinal >= 0)   cleanRows[i][colFinal]   = 0;
      if (colTotal >= 0)               cleanRows[i][colTotal]   = 0;
      if (colGrade >= 0)               cleanRows[i][colGrade]   = 0;
      tableRows.push({ num: studentNum, name: studentName, mid: 0, fin: hasFinal ? 0 : '—', total: 0, grade: 'F', status: 'excused' });
      continue;
    }

    const grades = bbLookup[studentNum];

    if (!grades) {
      zeroed++;
      if (colMidterm >= 0)             cleanRows[i][colMidterm] = 0;
      if (hasFinal && colFinal >= 0)   cleanRows[i][colFinal]   = 0;
      if (colTotal >= 0)               cleanRows[i][colTotal]   = 0;
      if (colGrade >= 0)               cleanRows[i][colGrade]   = 'F';
      tableRows.push({ num: studentNum, name: studentName, mid: 0, fin: hasFinal ? 0 : '—', total: 0, grade: 'F', status: 'missing' });
      continue;
    }

    matched++;
    const midVal = grades.midterm !== null ? Math.ceil(grades.midterm) : 0;
    const finVal = (hasFinal && grades.final !== null) ? Math.ceil(grades.final) : (hasFinal ? 0 : null);

    if (colMidterm >= 0)           cleanRows[i][colMidterm] = midVal;
    if (hasFinal && colFinal >= 0) cleanRows[i][colFinal]   = finVal;

    const total = Math.ceil(midVal + (hasFinal ? (finVal || 0) : 0));
    if (colTotal >= 0) cleanRows[i][colTotal] = total;

    const lg = letterGrade(total);
    if (colGrade >= 0) cleanRows[i][colGrade] = lg;

    tableRows.push({ num: studentNum, name: studentName, mid: midVal, fin: hasFinal ? finVal : '—', total, grade: lg, status: 'ok' });
  }

  // Output: only student number, name, midterm, (final if applicable)
  const keepCols = [colStudentNum, 1, colMidterm];
  if (hasFinal && colFinal >= 0) keepCols.push(colFinal);
  const outputRows = cleanRows.map(row => keepCols.map(c => row[c] !== undefined ? row[c] : ''));

  const newSheet = XLSX.utils.aoa_to_sheet(outputRows);
  resultWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(resultWorkbook, newSheet, sheetName);

  // Render results
  document.getElementById('noteBox').textContent = t('noteCheck');

  document.getElementById('summaryGrid').innerHTML = `
    <div class="summary-box">
      <div class="num">${totalStudents}</div>
      <div class="lbl">${t('totalStudents')}</div>
    </div>
    <div class="summary-box">
      <div class="num" style="color:#28a745">${matched}</div>
      <div class="lbl">${t('updated')}</div>
    </div>
    <div class="summary-box">
      <div class="num" style="color:#dc3545">${zeroed}</div>
      <div class="lbl">${t('notFound')}</div>
    </div>
  `;

  const missingStudents = tableRows.filter(r => r.status === 'missing').map(r => r.num);
  document.getElementById('missingAlert').innerHTML = missingStudents.length > 0
    ? `<div class="alert alert-warning">
        <strong>${t('warnTitle')}:</strong> ${t('warnMsg')}
        <ul class="missing-list">${missingStudents.map(s => `<li>${s}</li>`).join('')}</ul>
       </div>`
    : `<div class="alert alert-info">${t('successMsg')}</div>`;

  const finHeader = hasFinal ? `<th>${t('tblFin')}</th>` : '';
  document.getElementById('gradeTableHead').innerHTML = `
    <tr>
      <th>${t('tblNum')}</th>
      <th>${t('tblName')}</th>
      <th>${t('tblMid')}</th>
      ${finHeader}
      <th>${t('tblTotal')}</th>
      <th>${t('tblGrade')}</th>
      <th>${t('tblStatus')}</th>
    </tr>`;

  document.getElementById('gradeTableBody').innerHTML = tableRows.map(r => {
    const finCell  = hasFinal
      ? `<td class="${r.status === 'excused' ? 'excused' : (r.fin === 0 ? 'zero' : '')}">${r.fin}</td>`
      : '';
    const isPass   = typeof r.total === 'number' && r.total >= 60 && r.status !== 'excused';
    const badge    = r.status === 'excused' ? 'badge-excused' : (isPass ? 'badge-pass' : 'badge-fail');
    const statusLbl = r.status === 'ok' ? t('statusOk') : r.status === 'missing' ? t('statusMissing') : t('statusExcused');

    return `<tr>
      <td>${r.num}</td>
      <td>${r.name}</td>
      <td class="${r.status === 'excused' ? 'excused' : (r.mid === 0 && r.status !== 'ok' ? 'zero' : '')}">${r.mid}</td>
      ${finCell}
      <td class="${r.status === 'excused' ? 'excused' : (r.total === 0 && r.status !== 'ok' ? 'zero' : '')}">${r.total}</td>
      <td><span class="grade-badge ${badge}">${r.grade}</span></td>
      <td><span class="grade-badge ${badge}">${statusLbl}</span></td>
    </tr>`;
  }).join('');

  document.getElementById('step3').classList.add('visible');
  document.getElementById('step3').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ── Download ──────────────────────────────────────────────────────────────────

function downloadResult() {
  if (!resultWorkbook) return;
  const filename = lang === 'ar' ? 'كشف_الدرجات_المحدّث.xls' : 'Updated_Grade_Sheet.xls';
  XLSX.writeFile(resultWorkbook, filename, { bookType: 'xls' });
}

// ── Init ──────────────────────────────────────────────────────────────────────

applyLang();
