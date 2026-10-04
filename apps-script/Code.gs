// ═══════════════════════════════════════════════════════════════════
// RPM — Rekod Perkembangan Murid
// Google Apps Script v6.1
// SK Bandar Tawau | Guru: Harsidi bin Junick
//
// PERUBAHAN v6.1:
//   - Simpan (savePart) TIDAK lagi tulis ke Sheets setiap kali → simpan ~1-2 saat
//     (app panggil writeSheets sendiri ~20 saat selepas perubahan terakhir)
//   - writeSheets & importFromSheets kini boleh dipanggil melalui POST (app guna POST)
//   - Tolak simpan scores kosong jika server ada data (elak data hilang)
//   - writeSheets tidak padam Sheets jika tiada data TP
//   - importFromSheets: baca TP yang tersilap format tarikh (5 → 1900-01-04),
//     dan tarikh disimpan sebagai yyyy-MM-dd
//
// CARA DEPLOY:
//   1. Extensions → Apps Script → Tampal kod ini (gantikan SEMUA)
//   2. Simpan (Ctrl+S)
//   3. Deploy → New deployment (JANGAN edit deployment lama)
//   4. Type: Web app
//   5. Execute as: Me
//   6. Who has access: Anyone
//   7. Authorize → Salin URL baru → Tampal dalam RPM Tetapan & Sync → Simpan & Uji
// ═══════════════════════════════════════════════════════════════════

const SPREADSHEET_ID = ''; // kosong = auto (container-bound ATAU auto-create)
const KEY_SHEET_ID   = 'rpm_sheet_id';

function getSpreadsheet() {
  if (SPREADSHEET_ID) return SpreadsheetApp.openById(SPREADSHEET_ID);

  // 1. Cuba spreadsheet terikat (container-bound)
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (_) {}

  // 2. Cuba ID yang disimpan dalam ScriptProperties
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty(KEY_SHEET_ID);
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch (e) {
      Logger.log('Saved sheet id tidak sah, akan create baru: ' + e.message);
      props.deleteProperty(KEY_SHEET_ID);
    }
  }

  // 3. Auto-create spreadsheet baharu dan simpan ID
  const ss = SpreadsheetApp.create('RPM - Rekod Perkembangan Murid');
  props.setProperty(KEY_SHEET_ID, ss.getId());
  Logger.log('Spreadsheet baharu dicipta: ' + ss.getUrl());
  return ss;
}

function handleSheetInfo() {
  try {
    const ss = getSpreadsheet();
    return jsonResponse({ ok: true, id: ss.getId(), url: ss.getUrl(), name: ss.getName() });
  } catch (e) {
    return jsonResponse({ ok: false, error: e.message });
  }
}

// ─── KEYS ────────────────────────────────────────────────────────
const KEY_SCORES   = 'rpm_scores';
const KEY_POINTS   = 'rpm_points';
const KEY_BOOKS    = 'rpm_bookChecks';
const KEY_STUDENTS = 'rpm_students';
const KEY_SETTINGS = 'rpm_settings';
const CHUNK_SIZE   = 7500;

// ─── CURRICULUM DATA ─────────────────────────────────────────────
const SCIENCE_CLASSES = ['INOVATIF', 'INTELEK', 'KREATIF', 'KRITIS', 'PATRIOTIK'];

const SCIENCE_TOPICS = [
  { id: 't1',  title: '1.0 Kemahiran Saintifik',  standards: ['1.1.1','1.1.2','1.1.3','1.1.4','1.2.1','1.2.2','1.2.3','1.2.4','1.2.5'] },
  { id: 't2',  title: '2.0 Peraturan Bilik Sains', standards: ['2.1.1'] },
  { id: 't3',  title: '3.0 Manusia',               standards: ['3.1.1','3.1.2','3.1.3','3.1.4','3.1.5','3.1.6'] },
  { id: 't4',  title: '4.0 Haiwan',                standards: ['4.1.1','4.1.2','4.1.3','4.1.4','4.1.5','4.1.6','4.1.7'] },
  { id: 't5',  title: '5.0 Tumbuhan',              standards: ['5.1.1','5.1.2','5.1.3','5.1.4','5.1.5','5.1.6'] },
  { id: 't6',  title: '6.0 Terang dan Gelap',      standards: ['6.1.1','6.1.2','6.1.3','6.1.4','6.1.5','6.1.6'] },
  { id: 't7',  title: '7.0 Elektrik',              standards: ['7.1.1','7.1.2','7.1.3','7.1.4','7.1.5','7.1.6','7.1.7'] },
  { id: 't8',  title: '8.0 Campuran',              standards: ['8.1.1','8.1.2','8.1.3','8.1.4','8.1.5'] },
  { id: 't9',  title: '9.0 Bumi',                  standards: ['9.1.1','9.1.2','9.1.3','9.1.4','9.1.5','9.2.1','9.2.2','9.2.3','9.2.4','9.2.5','9.2.6'] },
  { id: 't10', title: '10.0 Teknologi',            standards: ['10.1.1','10.1.2','10.1.3','10.1.4','10.1.5','10.1.6'] }
];

const ENGLISH_SKILLS = [
  { id: 'ls', title: 'Listening',     standards: ['1.1.1','1.2.1','1.2.2','1.2.3','1.2.4','1.2.5','1.3.1'] },
  { id: 'sp', title: 'Speaking',      standards: ['2.1.1','2.1.2','2.1.3','2.1.4','2.1.5','2.2.1','2.2.2','2.3.1'] },
  { id: 'rd', title: 'Reading',       standards: ['3.2.1','3.2.2','3.2.3','3.2.4','3.3.1'] },
  { id: 'wr', title: 'Writing',       standards: ['4.1.2','4.2.1','4.2.2','4.2.3','4.2.4','4.2.5','4.3.1','4.3.2','4.3.3'] },
  { id: 'la', title: 'Language Arts', standards: ['5.1.1','5.1.2','5.2.1','5.3.1'] }
];

// ─── RESPONSE ────────────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function countScores(scores) {
  let n = 0;
  Object.values(scores || {}).forEach(subj =>
    Object.values(subj || {}).forEach(stu =>
      Object.values(stu || {}).forEach(topic =>
        { n += Object.keys(topic || {}).length; })));
  return n;
}

function handleWriteSheets() {
  try {
    const r = writeAllSheets();
    const ss = getSpreadsheet();
    return jsonResponse({ ok: true, written: r.written, records: r.records, url: ss.getUrl() });
  } catch (e) {
    return jsonResponse({ ok: false, error: e.message });
  }
}

// ═══════════════════════════════════════════════════════════════════
// GET HANDLER
// ═══════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || '';
    Logger.log('doGet action: ' + action);

    if (action === 'loadAll')          return handleLoadAll();
    if (action === 'ping')             return jsonResponse({ ok: true, version: '6.1', time: new Date().toISOString() });
    if (action === 'debug')            return handleDebug();
    if (action === 'cleanup')          return handleCleanup();
    if (action === 'importFromSheets') return handleImportFromSheets();
    if (action === 'sheetInfo')        return handleSheetInfo();
    if (action === 'resetPoints')      {
      const props = PropertiesService.getScriptProperties();
      const all = props.getProperties();
      const keys = Object.keys(all).filter(k => k === 'rpm_points' || k.startsWith('rpm_points_'));
      keys.forEach(k => props.deleteProperty(k));
      return jsonResponse({ ok: true, deleted: keys.length, keys });
    }
    if (action === 'writeSheets')      return handleWriteSheets();

    return jsonResponse({ ok: true, version: '6.1', message: 'RPM GAS v6.1. Guna ?action=ping atau ?action=loadAll' });
  } catch (err) {
    Logger.log('doGet error: ' + err.message);
    return jsonResponse({ error: err.message });
  }
}

// ═══════════════════════════════════════════════════════════════════
// POST HANDLER
// ═══════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    let payload = null;
    let parseMethod = '';

    // Method 1: JSON body
    if (e.postData && e.postData.type === 'application/json') {
      try {
        payload = JSON.parse(e.postData.contents);
        parseMethod = 'JSON body';
      } catch (pe) {
        Logger.log('JSON body parse gagal: ' + pe.message);
      }
    }

    // Method 2: form-encoded e.parameter.data (cara app RPM hantar)
    if (!payload && e.parameter && e.parameter.data) {
      try {
        payload = JSON.parse(e.parameter.data);
        parseMethod = 'e.parameter.data';
      } catch (pe) {
        Logger.log('e.parameter.data parse gagal: ' + pe.message);
      }
    }

    // Method 3: postData contents as raw JSON or URL-encoded
    if (!payload && e.postData && e.postData.contents) {
      const body = e.postData.contents;
      try {
        payload = JSON.parse(body);
        parseMethod = 'postData raw JSON';
      } catch (_) {
        try {
          const decoded = decodeURIComponent(body.replace(/\+/g, ' '));
          const pairs = decoded.split('&');
          for (const pair of pairs) {
            const eqIdx = pair.indexOf('=');
            if (eqIdx > -1 && pair.substring(0, eqIdx).trim() === 'data') {
              payload = JSON.parse(pair.substring(eqIdx + 1));
              parseMethod = 'URL-encoded body';
              break;
            }
          }
        } catch (ue) {
          Logger.log('URL-encoded parse gagal: ' + ue.message);
        }
      }
    }

    if (!payload) {
      return jsonResponse({ error: 'Tiada payload diterima.' });
    }

    Logger.log('Parse OK via: ' + parseMethod + ' | action: ' + (payload.action || 'tiada'));

    switch (payload.action || '') {
      case 'saveAll':          return handleSaveAll(payload);
      case 'savePart':         return handleSavePart(payload);
      case 'writeSheets':      return handleWriteSheets();
      case 'importFromSheets': return handleImportFromSheets();
      default:
        return jsonResponse({ error: 'Unknown action: ' + (payload.action || '') });
    }
  } catch (err) {
    Logger.log('doPost FATAL: ' + err.message + '\n' + err.stack);
    return jsonResponse({ error: err.message });
  }
}

// ═══════════════════════════════════════════════════════════════════
// PROPERTIES STORAGE (chunked for large data)
// ═══════════════════════════════════════════════════════════════════

function propSet(key, value) {
  const str   = JSON.stringify(value);
  const props = PropertiesService.getScriptProperties();

  // Delete old chunks
  const all   = props.getProperties();
  const toDel = Object.keys(all).filter(k =>
    k === key || k.startsWith(key + '_chunk_') || k === key + '_chunks'
  );
  if (toDel.length) toDel.forEach(k => props.deleteProperty(k));

  if (str.length <= CHUNK_SIZE) {
    props.setProperty(key, str);
    props.setProperty(key + '_chunks', '1');
  } else {
    const n = Math.ceil(str.length / CHUNK_SIZE);
    const batch = {};
    for (let i = 0; i < n; i++) {
      batch[key + '_chunk_' + i] = str.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
    }
    batch[key + '_chunks'] = String(n);
    props.setProperties(batch);  // single batch call — faster
    Logger.log('propSet ' + key + ': ' + n + ' chunks, ' + str.length + ' chars');
  }
}

function propGet(key) {
  try {
    const props      = PropertiesService.getScriptProperties();
    const chunkCount = parseInt(props.getProperty(key + '_chunks') || '0');

    if (chunkCount <= 1) {
      const val = props.getProperty(key);
      return val ? JSON.parse(val) : null;
    }

    let str = '';
    for (let i = 0; i < chunkCount; i++) {
      str += props.getProperty(key + '_chunk_' + i) || '';
    }
    return str ? JSON.parse(str) : null;
  } catch (e) {
    Logger.log('propGet ERROR ' + key + ': ' + e.message);
    return null;
  }
}

// ═══════════════════════════════════════════════════════════════════
// HANDLERS
// ═══════════════════════════════════════════════════════════════════

// ─── LOAD ALL ────────────────────────────────────────────────────
function handleLoadAll() {
  const scores     = propGet(KEY_SCORES)   || { science: {}, english: {} };
  const points     = propGet(KEY_POINTS)   || {};
  const bookChecks = propGet(KEY_BOOKS)    || {};
  const students   = propGet(KEY_STUDENTS);
  const settings   = propGet(KEY_SETTINGS) || {};

  const scoreCount = countScores(scores);
  Logger.log('loadAll: ' + scoreCount + ' rekod TP');

  return jsonResponse({
    ok: true,
    scores,
    points,
    bookChecks,
    students,
    settings,
    _meta: { scoreCount, version: '6.1' }
  });
}

// ─── SAVE ALL (versi lama app — tidak digunakan oleh app semasa) ──
function handleSaveAll(payload) {
  const t0 = new Date();

  // PERLINDUNGAN: Tolak jika scores kosong — jangan timpa data sedia ada
  const incomingCount = countScores(payload.scores);
  if (incomingCount === 0) {
    const existingCount = countScores(propGet(KEY_SCORES));
    if (existingCount > 0) {
      Logger.log('DITOLAK: Cuba timpa ' + existingCount + ' rekod dengan data kosong!');
      return jsonResponse({
        ok: false,
        error: 'Ditolak: data masuk kosong (0 rekod) tetapi server ada ' + existingCount + ' rekod.',
        existingCount
      });
    }
  }

  let saved = [];

  if (payload.scores)     { propSet(KEY_SCORES, payload.scores);     saved.push('scores'); }
  if (payload.bookChecks) { propSet(KEY_BOOKS, payload.bookChecks);  saved.push('bookChecks'); }
  if (payload.students)   { propSet(KEY_STUDENTS, payload.students); saved.push('students'); }
  if (payload.settings)   { propSet(KEY_SETTINGS, payload.settings); saved.push('settings'); }

  if (payload.points) {
    try {
      const trimmed = {};
      Object.entries(payload.points).forEach(([key, val]) => {
        trimmed[key] = { total: val.total || 0, log: (val.log || []).slice(0, 5) };
      });
      propSet(KEY_POINTS, trimmed);
      saved.push('points');
    } catch (e) {
      Logger.log('Points save gagal (quota?): ' + e.message + ' — skip');
    }
  }

  let sheetsResult = null;
  try {
    sheetsResult = writeAllSheets();
  } catch (e) {
    Logger.log('writeSheets error (non-fatal): ' + e.message);
  }

  return jsonResponse({
    ok: true,
    saved: saved.join('+'),
    sheets: sheetsResult ? { written: sheetsResult.written, records: sheetsResult.records } : null,
    ms: new Date() - t0
  });
}

// ─── SAVE PART (pantas — Properties sahaja) ──────────────────────
function handleSavePart(payload) {
  const part = payload.part || '';
  let saved = [];

  if (part === 'scores') {
    if (payload.scores !== undefined) {
      // PERLINDUNGAN: jangan timpa data server dengan scores kosong
      const existingCount = countScores(propGet(KEY_SCORES));
      if (countScores(payload.scores) === 0 && existingCount > 0) {
        Logger.log('savePart DITOLAK: scores kosong, server ada ' + existingCount);
        return jsonResponse({ ok: false, error: 'Ditolak: data TP kosong tetapi server ada ' + existingCount + ' rekod. Muat semula app.' });
      }
      propSet(KEY_SCORES, payload.scores);
      saved.push('scores');
    }
    if (payload.points !== undefined) {
      try {
        const trimmed = {};
        Object.entries(payload.points).forEach(([k, v]) => {
          trimmed[k] = { total: v.total || 0, log: (v.log || []).slice(0, 5) };
        });
        propSet(KEY_POINTS, trimmed);
        saved.push('points');
      } catch (e) { Logger.log('Points save gagal: ' + e.message); }
    }
    if (payload.bookChecks !== undefined) { propSet(KEY_BOOKS, payload.bookChecks); saved.push('bookChecks'); }
  } else if (part === 'meta') {
    if (payload.students !== undefined)   { propSet(KEY_STUDENTS, payload.students); saved.push('students'); }
    if (payload.settings !== undefined)   { propSet(KEY_SETTINGS, payload.settings); saved.push('settings'); }
  } else {
    return jsonResponse({ error: 'Unknown part: ' + part });
  }

  // Sheets TIDAK ditulis di sini — app panggil writeSheets selepas guru berhenti menaip
  return jsonResponse({ ok: true, saved: saved.join('+'), part });
}

// ═══════════════════════════════════════════════════════════════════
// WRITE TO GOOGLE SHEETS
// ═══════════════════════════════════════════════════════════════════

function writeAllSheets() {
  const ss       = getSpreadsheet();
  const scores   = propGet(KEY_SCORES)   || { science: {}, english: {} };
  const points   = propGet(KEY_POINTS)   || {};
  const books    = propGet(KEY_BOOKS)    || {};
  const students = propGet(KEY_STUDENTS);
  const settings = propGet(KEY_SETTINGS) || {};

  // Sheets ialah salinan sandaran — jangan padam dengan data kosong
  if (countScores(scores) === 0) throw new Error('Tiada data TP di server — Sheets tidak ditimpa.');

  function getOrCreate(name) {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    s.clearContents();
    s.clearFormats();
    return s;
  }

  function getTP(subject, cls, name, topicId, sp) {
    const key   = subject === 'science' ? cls + '__' + name : name;
    return (((scores[subject] || {})[key] || {})[topicId] || {})[sp] || null;
  }

  function studs(subject, cls) {
    if (!students) return [];
    return subject === 'science'
      ? ((students.science || {})[cls] || [])
      : (students.english || []);
  }

  function styleHeader(sheet, row, cols, bg) {
    sheet.getRange(row, 1, 1, cols)
      .setBackground(bg).setFontColor('#ffffff').setFontWeight('bold').setFontSize(10);
  }

  let written = 0;
  let totalRecords = 0;

  // ── Sains setiap kelas ─────────────────────────────────────
  SCIENCE_CLASSES.forEach(cls => {
    const sheet = getOrCreate('SAINS_' + cls);
    const list  = studs('science', cls);

    const h1 = ['Bil', 'Nama Murid'];
    const h2 = ['', ''];
    SCIENCE_TOPICS.forEach(t => {
      h1.push(t.title);
      for (let i = 1; i < t.standards.length; i++) h1.push('');
      t.standards.forEach(sp => h2.push(sp));
    });
    h1.push('Purata TP'); h2.push('');

    const rows = [h1, h2];
    list.forEach((name, idx) => {
      const row = [idx + 1, name];
      let sum = 0, cnt = 0;
      SCIENCE_TOPICS.forEach(t => {
        t.standards.forEach(sp => {
          const sc = getTP('science', cls, name, t.id, sp);
          const tp = sc ? sc.tp : '';
          row.push(tp);
          if (tp !== '') { sum += Number(tp); cnt++; totalRecords++; }
        });
      });
      row.push(cnt > 0 ? Math.round(sum / cnt * 100) / 100 : '');
      rows.push(row);
    });

    if (rows.length > 2) {
      sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
      styleHeader(sheet, 1, rows[0].length, '#7c3aed');
      styleHeader(sheet, 2, rows[0].length, '#9d5cf0');
      sheet.setFrozenRows(2);
      sheet.setFrozenColumns(2);
      try { sheet.autoResizeColumns(1, Math.min(rows[0].length, 26)); } catch (_) {}
    }
    written++;
  });

  // ── BI Tahun 3 ────────────────────────────────────────────
  const biSheet = getOrCreate('BI_TAHUN3');
  const biList  = studs('english', null);
  const bh1 = ['Bil', 'Nama Murid'];
  const bh2 = ['', ''];
  ENGLISH_SKILLS.forEach(s => {
    bh1.push(s.title);
    for (let i = 1; i < s.standards.length; i++) bh1.push('');
    s.standards.forEach(sp => bh2.push(sp));
  });
  bh1.push('Purata PL'); bh2.push('');

  const biRows = [bh1, bh2];
  biList.forEach((name, idx) => {
    const row = [idx + 1, name];
    let sum = 0, cnt = 0;
    ENGLISH_SKILLS.forEach(s => {
      s.standards.forEach(sp => {
        const sc = getTP('english', null, name, s.id, sp);
        const tp = sc ? sc.tp : '';
        row.push(tp);
        if (tp !== '') { sum += Number(tp); cnt++; totalRecords++; }
      });
    });
    row.push(cnt > 0 ? Math.round(sum / cnt * 100) / 100 : '');
    biRows.push(row);
  });

  if (biRows.length > 2) {
    biSheet.getRange(1, 1, biRows.length, biRows[0].length).setValues(biRows);
    styleHeader(biSheet, 1, biRows[0].length, '#0891b2');
    styleHeader(biSheet, 2, biRows[0].length, '#0ea5c9');
    biSheet.setFrozenRows(2); biSheet.setFrozenColumns(2);
    try { biSheet.autoResizeColumns(1, Math.min(biRows[0].length, 26)); } catch (_) {}
  }
  written++;

  // ── Tarikh Taksiran ───────────────────────────────────────
  const tarSheet = getOrCreate('TARIKH_TAKSIRAN');
  const tarRows  = [['Bil', 'Subjek', 'Kelas', 'Nama Murid', 'Topik', 'SP', 'TP', 'Tarikh']];
  let tarIdx = 0;

  SCIENCE_CLASSES.forEach(cls => {
    studs('science', cls).forEach(name => {
      SCIENCE_TOPICS.forEach(t => {
        t.standards.forEach(sp => {
          const sc = getTP('science', cls, name, t.id, sp);
          if (sc) tarRows.push([++tarIdx, 'Sains', cls, name, t.title, sp, sc.tp, sc.date || '']);
        });
      });
    });
  });
  studs('english', null).forEach(name => {
    ENGLISH_SKILLS.forEach(s => {
      s.standards.forEach(sp => {
        const sc = getTP('english', null, name, s.id, sp);
        if (sc) tarRows.push([++tarIdx, 'BI', 'Tahun 3', name, s.title, sp, sc.tp, sc.date || '']);
      });
    });
  });

  // SP & Tarikh sebagai teks, TP sebagai nombor — elak Sheets tukar jadi tarikh
  tarSheet.getRange(1, 6, tarRows.length, 1).setNumberFormat('@');
  tarSheet.getRange(1, 7, tarRows.length, 1).setNumberFormat('0');
  tarSheet.getRange(1, 8, tarRows.length, 1).setNumberFormat('@');
  tarSheet.getRange(1, 1, tarRows.length, 8).setValues(tarRows);
  styleHeader(tarSheet, 1, 8, '#059669');
  tarSheet.setFrozenRows(1);
  try { tarSheet.autoResizeColumns(1, 8); } catch (_) {}
  written++;

  // ── Mata Terkumpul ────────────────────────────────────────
  const ptsSheet = getOrCreate('MATA_TERKUMPUL');
  const ptsRows  = [['Bil', 'Nama Murid', 'Kelas/Subjek', 'Jumlah Mata', 'Mata TP', 'Mata Buku', 'Mata Manual', 'Mata Tolak']];
  const allPts   = [];

  SCIENCE_CLASSES.forEach(cls => {
    studs('science', cls).forEach(name => {
      const p   = points[cls + '__' + name] || { total: 0, log: [] };
      const log = p.log || [];
      allPts.push([0, name, 'Sains - ' + cls, p.total || 0,
        log.filter(l => l.type === 'tp').reduce((a, l) => a + (l.pts || 0), 0),
        log.filter(l => l.type === 'book').reduce((a, l) => a + (l.pts || 0), 0),
        log.filter(l => l.type === 'manual').reduce((a, l) => a + (l.pts || 0), 0),
        log.filter(l => l.type === 'deduct').reduce((a, l) => a + Math.abs(l.pts || 0), 0)
      ]);
    });
  });
  studs('english', null).forEach(name => {
    const p   = points[name] || { total: 0, log: [] };
    const log = p.log || [];
    allPts.push([0, name, 'BI Tahun 3', p.total || 0,
      log.filter(l => l.type === 'tp').reduce((a, l) => a + (l.pts || 0), 0),
      log.filter(l => l.type === 'book').reduce((a, l) => a + (l.pts || 0), 0),
      log.filter(l => l.type === 'manual').reduce((a, l) => a + (l.pts || 0), 0),
      log.filter(l => l.type === 'deduct').reduce((a, l) => a + Math.abs(l.pts || 0), 0)
    ]);
  });

  allPts.sort((a, b) => b[3] - a[3]);
  allPts.forEach((r, i) => { r[0] = i + 1; ptsRows.push(r); });
  ptsSheet.getRange(1, 1, ptsRows.length, 8).setValues(ptsRows);
  styleHeader(ptsSheet, 1, 8, '#7c3aed');
  ptsSheet.setFrozenRows(1);
  written++;

  // ── Semakan Buku ──────────────────────────────────────────
  const bukuSheet = getOrCreate('SEMAKAN_BUKU');
  const bukuRows  = [['Bil', 'Tarikh', 'Nama Murid', 'Kelas/Subjek', 'Nota', 'Mata']];
  const allBuku   = [];

  SCIENCE_CLASSES.forEach(cls => {
    studs('science', cls).forEach(name => {
      ((books[cls + '__' + name] || {}).records || []).forEach(r =>
        allBuku.push([0, r.date || '', name, 'Sains - ' + cls, r.note || '', r.pts || 2])
      );
    });
  });
  studs('english', null).forEach(name => {
    ((books[name] || {}).records || []).forEach(r =>
      allBuku.push([0, r.date || '', name, 'BI Tahun 3', r.note || '', r.pts || 2])
    );
  });

  allBuku.sort((a, b) => String(b[1]).localeCompare(String(a[1])));
  allBuku.forEach((r, i) => { r[0] = i + 1; bukuRows.push(r); });
  bukuSheet.getRange(1, 1, bukuRows.length, 6).setValues(bukuRows);
  styleHeader(bukuSheet, 1, 6, '#d97706');
  bukuSheet.setFrozenRows(1);
  written++;

  // ── Status / Meta ─────────────────────────────────────────
  const metaSheet = getOrCreate('RPM_META');
  const now = new Date();
  metaSheet.getRange(1, 1, 7, 2).setValues([
    ['Kunci',            'Nilai'],
    ['Versi GAS',        '6.1'],
    ['Sekolah',          settings.school || 'SK Bandar Tawau'],
    ['Guru',             settings.teacher || 'Harsidi bin Junick'],
    ['Terakhir dikemas', now.toLocaleString('ms-MY')],
    ['Tab ditulis',      written],
    ['Rekod TP',         tarIdx]
  ]);
  styleHeader(metaSheet, 1, 2, '#0e0f11');
  written++;

  return { written, records: tarIdx };
}

// ─── IMPORT FROM SHEETS → PropertiesService ──────────────────────
function handleImportFromSheets() {
  Logger.log('handleImportFromSheets dipanggil');

  // Auto-cleanup: padam key lama untuk bebaskan ruang
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const keepPrefixes = ['rpm_scores', 'rpm_students', 'rpm_settings', 'rpm_bookChecks', 'rpm_points', 'rpm_sheet_id', 'ANTHROPIC_API_KEY'];
  let cleaned = 0;
  Object.keys(all).forEach(k => {
    if (!keepPrefixes.some(p => k.startsWith(p))) {
      props.deleteProperty(k);
      cleaned++;
    }
  });
  Logger.log('Cleanup: removed ' + cleaned + ' old keys');

  const ss = getSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  // Build SP → topicId lookup
  const spMap = {};
  SCIENCE_TOPICS.forEach(t => t.standards.forEach(sp => { spMap['sci_' + sp] = t.id; }));
  ENGLISH_SKILLS.forEach(s => s.standards.forEach(sp => { spMap['eng_' + sp] = s.id; }));

  const tarSheet = ss.getSheetByName('TARIKH_TAKSIRAN');
  if (!tarSheet) return jsonResponse({ error: 'Tab TARIKH_TAKSIRAN tidak dijumpai.' });

  const data = tarSheet.getDataRange().getValues();
  if (data.length < 2) return jsonResponse({ error: 'Tab TARIKH_TAKSIRAN kosong.' });

  // TP yang tersilap format sebagai tarikh (5 → 1900-01-04) → nombor semula
  function tpValue(v) {
    if (v instanceof Date) v = Math.round((v.getTime() - new Date(1899, 11, 30).getTime()) / 864e5);
    return parseInt(v);
  }
  function dateValue(v) {
    if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    return v ? String(v).trim() : '';
  }

  const newScores = { science: {}, english: {} };
  let imported = 0, skipped = 0;

  for (let i = 1; i < data.length; i++) {
    const row     = data[i];
    const isSains = String(row[1] || '').toLowerCase().includes('sains');
    const cls     = String(row[2] || '').trim();
    const name    = String(row[3] || '').trim();
    const sp      = String(row[5] || '').trim();
    const tp      = tpValue(row[6]);
    const date    = dateValue(row[7]);

    if (!name || !sp || isNaN(tp) || tp < 1 || tp > 6) { skipped++; continue; }

    const subject = isSains ? 'science' : 'english';
    const topicId = spMap[(isSains ? 'sci_' : 'eng_') + sp];
    if (!topicId) { skipped++; continue; }

    if (subject === 'science') {
      const key = cls + '__' + name;
      if (!newScores.science[key]) newScores.science[key] = {};
      if (!newScores.science[key][topicId]) newScores.science[key][topicId] = {};
      newScores.science[key][topicId][sp] = { tp, date };
    } else {
      if (!newScores.english[name]) newScores.english[name] = {};
      if (!newScores.english[name][topicId]) newScores.english[name][topicId] = {};
      newScores.english[name][topicId][sp] = { tp, date };
    }
    imported++;
  }

  if (imported === 0) return jsonResponse({ ok: true, imported: 0, skipped, error: 'Tiada rekod TP sah dalam TARIKH_TAKSIRAN.' });
  propSet(KEY_SCORES, newScores);

  // Also import student lists from Sheet tabs
  const stud = { science: {}, english: [] };
  SCIENCE_CLASSES.forEach(cls => {
    const sh = ss.getSheetByName('SAINS_' + cls);
    if (!sh) return;
    const rows = sh.getDataRange().getValues();
    const names = [];
    for (let i = 2; i < rows.length; i++) {
      const n = String(rows[i][1] || '').trim();
      if (n) names.push(n);
    }
    if (names.length) stud.science[cls] = names;
  });

  const biSh = ss.getSheetByName('BI_TAHUN3');
  if (biSh) {
    const rows = biSh.getDataRange().getValues();
    for (let i = 2; i < rows.length; i++) {
      const n = String(rows[i][1] || '').trim();
      if (n) stud.english.push(n);
    }
  }

  if (Object.keys(stud.science).length || stud.english.length) {
    propSet(KEY_STUDENTS, stud);
  }

  Logger.log('importFromSheets: ' + imported + ' rekod, ' + skipped + ' dilangkau');
  return jsonResponse({ ok: true, imported, skipped });
}

// ─── CLEANUP — padam key lama untuk bebaskan ruang ──────────────
function handleCleanup() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const keepPrefixes = ['rpm_scores', 'rpm_students', 'rpm_settings', 'rpm_bookChecks', 'rpm_points', 'rpm_sheet_id', 'ANTHROPIC_API_KEY'];
  let deleted = [];
  Object.keys(all).forEach(k => {
    const keep = keepPrefixes.some(p => k.startsWith(p));
    if (!keep) {
      props.deleteProperty(k);
      deleted.push(k);
    }
  });
  Logger.log('Cleanup: deleted ' + deleted.length + ' keys');
  return jsonResponse({ ok: true, deleted: deleted.length, keys: deleted });
}

// ─── DEBUG ───────────────────────────────────────────────────────
function handleDebug() {
  const props   = PropertiesService.getScriptProperties().getProperties();
  const summary = {};
  Object.keys(props).forEach(k => { summary[k] = (props[k] || '').length + ' chars'; });
  const scoreCount = countScores(propGet(KEY_SCORES));
  return jsonResponse({ ok: true, version: '6.1', keys: summary, scoreCount, time: new Date().toISOString() });
}

// ═══════════════════════════════════════════════════════════════════
// TEST FUNCTIONS — jalankan dari Apps Script editor
// ═══════════════════════════════════════════════════════════════════

function testConnection() {
  try {
    const ss = getSpreadsheet();
    Logger.log('✅ Spreadsheet: ' + ss.getName());
    propSet('_test', { hello: 'world', t: new Date().toISOString() });
    const v = propGet('_test');
    Logger.log('✅ Properties OK: ' + JSON.stringify(v));
    PropertiesService.getScriptProperties().deleteProperty('_test');
    PropertiesService.getScriptProperties().deleteProperty('_test_chunks');
    Logger.log('✅ Semua OK!');
  } catch (e) {
    Logger.log('❌ ERROR: ' + e.message);
  }
}

function checkStoredData() {
  Logger.log('Rekod TP: ' + countScores(propGet(KEY_SCORES)));
  Logger.log('Keys: ' + Object.keys(PropertiesService.getScriptProperties().getProperties()).join(', '));
}

/** AWAS: padam semua data */
function clearAllData() {
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log('✅ Semua data dipadam.');
}
