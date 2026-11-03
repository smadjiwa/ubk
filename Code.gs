const SPREADSHEET_ID = '1Myw1o9dh1D_HgAtAGZEKu_YpT3bGq0TyKNOejtTIWLM';
const DATABASE_SHEET = 'PROFIL MURID';
const USER_SHEET = 'USER';
const IMAGE_SHEET = 'RegGambar';
const AUTH_CACHE_PREFIX = 'UBK_AUTH_';
const AUTH_TTL = 21600; // 6 jam
const MODULE_TAB_MAP = {
  1: 'PROFIL MURID',
  2: 'GAYA BELAJAR',
  3: 'PERSONALITI WARNA',
  4: 'MINAT KERJAYA',
  5: 'TRET PERSONALITI',
  6: 'KECERDASAN PELBAGAI',
  7: 'AKADEMIK',
  8: 'MINDA SIHAT'
};
const DATABASE_COLUMNS = 66; // A:BN

// Data yang diperlukan untuk dashboard awal.
// Sensitif PHQ/GAD dan NO KP tidak dihantar pada loading awal.
const DASHBOARD_HEADERS = [
  'BIL','NAMA MURID','JANTINA','TINGKATAN','KELAS','RUMAH','MENTOR','DORM','JAWATAN',
  'GAYA BELAJAR','PERSONALITI WARNA',
  'REALISTIK','INVESTIGATIF','ARTISTIK','SOSIAL','ENTERPRISING','KONVENSIONAL',
  'AUTONOMI','KREATIF','AGRESIF','EKSTROVERT','PENCAPAIAN','KEPELBAGAIAN','INTELEKTUAL',
  'KEPEMIMPINAN','STRUKTUR','RESILIENCE','MENOLONG','ANALITIKAL','KRITIK DIRI','WAWASAN','KETELUSAN',
  'VLM','VLI','LM','VR','MZK','NAT','INTRA','INTER','KINES','EKS',
  'BM (P1)','BI (P1)','SEJ (P1)','GEO (P1)','MM (P1)','SN (P1)','PAI (P1)','BAK (P1)',
  'RBT (P1)','ASK (P1)','FIZ (P1)','KIM (P1)','BIO (P1)','MT (P1)','MRL (P1)','PQS (P1)','PSI (P1)'
];

const MINDA_HEADERS = [
  'NAMA MURID','JANTINA','TINGKATAN','KELAS',
  'TAHAP PHQ-9','TINDAKAN (PHQ-9)','KRISIS (PHQ-9)',
  'TAHAP (GAD-7)','TINDAKAN (GAD-7)','KRISIS (GAD-7)'
];

function doGet(e) {
  // Endpoint kecil untuk bridge apabila frontend dihoskan di GitHub Pages.
  if (e && e.parameter && e.parameter.bridge === '1') {
    return HtmlService.createHtmlOutput(getGitHubBridgeHtml_())
      .setTitle('UBK API Bridge')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard UBK - SMSS Sultan Mohamad Jiwa')
    .setFaviconUrl('https://mylink.la/storage/avatars/a7/51148-a7263e64de7d87e7.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Halaman bridge minimal; tiada data atau UI dashboard dipaparkan di iframe. */
function getGitHubBridgeHtml_() {
  const allowedMethods = [
    'authenticateUser','logoutUser','getSessionUser','getAdminSettings',
    'updateUpperSecondaryGrades','updateLowerSecondaryGrades','getCoreData',
    'getFullDashboardData','getAkademikExamsByStudent','getDormAnalysisData',
    'getSummaryData','getMindaSihatData','getStudentProfile',
    'getStudentAcademicP1','checkMindaAccess'
  ];
  return `<!doctype html><html><head><meta charset="utf-8"></head><body><script>
    const ALLOWED = ${JSON.stringify(allowedMethods)};
    const TRUSTED = /^https:\/\/[a-z0-9-]+\.github\.io$/i;
    function send(data, origin){ try { parent.postMessage(data, origin); } catch(e) {} }
    addEventListener('message', function(ev) {
      if (ev.source !== parent || !TRUSTED.test(ev.origin)) return;
      const d = ev.data || {};
      if (d.type !== 'UBK_CALL' || !ALLOWED.includes(d.method) || !Array.isArray(d.args)) return;
      let runner = google.script.run
        .withSuccessHandler(function(result){ send({type:'UBK_RESULT',id:d.id,ok:true,result:result}, ev.origin); })
        .withFailureHandler(function(error){ send({type:'UBK_RESULT',id:d.id,ok:false,error:error && error.message ? error.message : String(error)}, ev.origin); });
      try { runner[d.method].apply(runner, d.args); }
      catch(err) { send({type:'UBK_RESULT',id:d.id,ok:false,error:err.message || String(err)}, ev.origin); }
    });
    try { parent.postMessage({type:'UBK_READY'}, '*'); } catch(e) {}
  <\/script></body></html>`;
}


function getUserSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(USER_SHEET);
  if (!sheet) throw new Error('Tab USER tidak dijumpai. Sila cipta tab USER dalam Google Sheet.');
  return sheet;
}

function normalizeHeader_(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, '');
}

function authenticateUser(username, password) {
  const user = String(username || '').trim();
  const pass = String(password || '');
  if (!user || !pass) return { ok: false, message: 'Sila masukkan username dan password.' };

  const sheet = getUserSheet_();
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 10);
  if (lastRow < 2) return { ok: false, message: 'Tiada akaun pengguna dalam tab USER.' };

  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = values[0].map(normalizeHeader_);
  const userCol = headers.indexOf('user');
  const passCol = headers.indexOf('password');
  if (userCol < 0 || passCol < 0) throw new Error('Tab USER mesti mempunyai kolum user dan password.');

  const moduleCols = {};
  for (let i = 1; i <= 8; i++) moduleCols[i] = headers.indexOf('modul' + i);

  let matched = null;
  for (let r = 1; r < values.length; r++) {
    const rowUser = String(values[r][userCol] || '').trim();
    if (rowUser && rowUser.toLowerCase() === user.toLowerCase()) {
      matched = values[r];
      break;
    }
  }

  if (!matched) return { ok: false, message: 'Username atau password tidak sah.' };
  if (String(matched[passCol] || '') !== pass) return { ok: false, message: 'Username atau password tidak sah.' };

  const permissions = {};
  for (let i = 1; i <= 8; i++) {
    const col = moduleCols[i];
    permissions['Modul' + i] = col >= 0 && String(matched[col] || '').trim() === '1' ? 1 : 0;
  }

  const token = Utilities.getUuid();
  const session = {
    username: String(matched[userCol] || user).trim(),
    permissions: permissions,
    created: new Date().toISOString()
  };
  CacheService.getScriptCache().put(AUTH_CACHE_PREFIX + token, JSON.stringify(session), AUTH_TTL);

  return { ok: true, token: token, user: { username: session.username }, permissions: permissions, expiresIn: AUTH_TTL };
}

function isAdminSession_(token) {
  const session = validateSession_(token);
  return String(session.username || '').trim().toLowerCase() === 'admin';
}

function requireAdmin_(token) {
  const session = validateSession_(token);
  if (String(session.username || '').trim().toLowerCase() !== 'admin') {
    throw new Error('Akses ditolak. Modul Setting hanya untuk user admin.');
  }
  return session;
}

function getAdminSettings(token) {
  const session = requireAdmin_(token);
  return {
    ok: true,
    username: session.username,
    permissions: session.permissions
  };
}


/**
 * ================================================================
 * KEMASKINI GRED MENENGAH ATAS
 * Sumber  : Markah_MenAtas
 * Sasaran : PROFIL MURID
 * Key     : NO KP (dinormalisasi 12 digit)
 *
 * Mapping:
 * F BM    -> AX BM (P1)
 * G BI    -> AY BI (P1)
 * H PI    -> BD PAI (P1)
 * I PM(A) -> BL MRL (P1)
 * J SEJ   -> AZ SEJ (P1)
 * K MAT   -> BB MM (P1)
 * L MT    -> BK MT (P1)
 * M FIZ   -> BH FIZ (P1)
 * N KIM   -> BI KIM (P1)
 * O BIO   -> BJ BIO (P1)
 * P/Q/R/S -> BE BAK (P1), guna nilai pertama yang tidak kosong
 * ================================================================
 */
function updateUpperSecondaryGrades(token) {
  requireAdmin_(token);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const source = ss.getSheetByName('Markah_MenAtas');
  const target = ss.getSheetByName('PROFIL MURID');

  if (!source) throw new Error('Tab Markah_MenAtas tidak dijumpai.');
  if (!target) throw new Error('Tab PROFIL MURID tidak dijumpai.');

  const sourceLastRow = source.getLastRow();
  const targetLastRow = target.getLastRow();

  if (sourceLastRow < 2) {
    return {
      ok: true,
      updated: 0,
      notFound: 0,
      skipped: 0,
      sourceRows: 0,
      message: 'Tiada rekod dalam Markah_MenAtas.'
    };
  }
  if (targetLastRow < 2) {
    throw new Error('Tiada data murid dalam PROFIL MURID.');
  }

  // Baca header + data sumber. F:S ialah markah/gred.
  const sourceValues = source.getRange(1, 1, sourceLastRow, Math.max(source.getLastColumn(), 19)).getDisplayValues();
  const sourceHeaders = sourceValues[0].map(h => String(h || '').trim());

  // Cari kolum NO KP berdasarkan header supaya tidak bergantung pada kedudukan.
  const kpAliases = [
    'no kp', 'no. kp', 'no.kp', 'no kp murid', 'no. kp murid',
    'nokp', 'no_kp', 'kp', 'ic', 'no kad pengenalan'
  ];
  const normalized = sourceHeaders.map(normalizeHeader_);
  let sourceKPCol = -1;
  for (let i = 0; i < normalized.length; i++) {
    if (kpAliases.indexOf(normalized[i]) >= 0) {
      sourceKPCol = i;
      break;
    }
  }
  if (sourceKPCol < 0) {
    throw new Error('Kolum No KP tidak dijumpai dalam tab Markah_MenAtas. Letakkan header "No KP" / "No. KP" pada baris pertama.');
  }

  // PROFIL MURID: C = No KP, AX:BN = markah/gred.
  const targetKPValues = target.getRange(2, 3, targetLastRow - 1, 1).getDisplayValues();
  const targetGradeValues = target.getRange(2, 50, targetLastRow - 1, 17).getValues(); // AX:BN

  const targetByKP = new Map();
  targetKPValues.forEach((row, idx) => {
    const kp = normalizeKP12_(row[0]);
    if (kp && !targetByKP.has(kp)) targetByKP.set(kp, idx);
  });

  // Source columns are 0-based: F=5 ... S=18.
  // Source F:S -> Target AX:BN
  // AX=0, AY=1, AZ=2, BA=3, BB=4, BC=5, BD=6,
  // BE=7, BF=8, BG=9, BH=10, BI=11, BJ=12,
  // BK=13, BL=14, BM=15, BN=16
  const mappings = [
    [5, 0],   // F BM    -> AX BM (P1)
    [6, 1],   // G BI    -> AY BI (P1)
    [7, 6],   // H PI    -> BD PAI (P1)
    [8, 14],  // I PM(A) -> BL MRL (P1)
    [9, 2],   // J SEJ   -> AZ SEJ (P1)
    [10, 4],  // K MAT   -> BB MM (P1)
    [11, 13], // L MT    -> BK MT (P1)
    [12, 10], // M FIZ   -> BH FIZ (P1)
    [13, 11], // N KIM   -> BI KIM (P1)
    [14, 12]  // O BIO   -> BJ BIO (P1)
  ];

  let updated = 0;
  let notFound = 0;
  let skipped = 0;
  const missing = [];
  const duplicateTargetKP = [];

  for (let r = 1; r < sourceValues.length; r++) {
    const row = sourceValues[r];
    const kp = normalizeKP12_(row[sourceKPCol]);

    if (!kp) {
      skipped++;
      continue;
    }

    const targetIdx = targetByKP.get(kp);
    if (targetIdx === undefined) {
      notFound++;
      if (missing.length < 100) missing.push(kp);
      continue;
    }

    let changed = false;

    // F:O -> AX,AY,BD,BL,AZ,BB,BK,BH,BI,BJ.
    mappings.forEach(([sourceIdx, targetGradeIdx]) => {
      const value = String(row[sourceIdx] ?? '').trim();
      if (value !== '') {
        targetGradeValues[targetIdx][targetGradeIdx] = value;
        changed = true;
      }
    });

    // P/Q/R/S -> BE. Ambil nilai pertama yang tidak kosong (P > Q > R > S).
    const altCandidates = [row[15], row[16], row[17], row[18]]
      .map(v => String(v ?? '').trim())
      .filter(v => v !== '');
    if (altCandidates.length) {
      targetGradeValues[targetIdx][3] = altCandidates[0]; // BE = AX index 3
      changed = true;
    }

    if (changed) updated++;
  }

  // Satu write batch untuk AX:BN, mengurangkan panggilan ke Spreadsheet.
  target.getRange(2, 50, targetLastRow - 1, 17).setValues(targetGradeValues);

  return {
    ok: true,
    updated: updated,
    notFound: notFound,
    skipped: skipped,
    sourceRows: sourceLastRow - 1,
    targetRows: targetLastRow - 1,
    missingNoKP: missing,
    mapping: {
      'F BM': 'AX BM (P1)',
      'G BI': 'AY BI (P1)',
      'H PI': 'BD PAI (P1)',
      'I PM(A)': 'BL MRL (P1)',
      'J SEJ': 'AZ SEJ (P1)',
      'K MAT': 'BB MM (P1)',
      'L MT': 'BK MT (P1)',
      'M FIZ': 'BH FIZ (P1)',
      'N KIM': 'BI KIM (P1)',
      'O BIO': 'BJ BIO (P1)',
      'P/Q/R/S': 'BE BAK (P1)'
    },
    message: 'Kemaskini gred Menengah Atas selesai.'
  };
}

function normalizeKP12_(value) {
  const digits = String(value == null ? '' : value).replace(/\D/g, '');
  if (!digits) return '';
  return digits.length < 12 ? digits.padStart(12, '0') : digits.slice(-12);
}



/**
 * ================================================================
 * KEMASKINI GRED MENENGAH RENDAH
 * Key: KAD PENGENALAN / NO KP -> PROFIL MURID!C
 *
 * Sumber:
 * C KAD PENGENALAN
 * F BM | G BI | H PI | I PM(R) | J SEJ | K GEO | L SNS
 * M MAT | N RBT | O ASK | P PSV | Q PSM | R BJ | S BC | T BA | U BG
 *
 * Target PROFIL MURID:
 * AX BM | AY BI | AZ SEJ | BA GEO | BB MM | BC SN | BD PAI
 * BE BAK | BF RBT | BG ASK | BM PQS | BN PSI | BL MRL
 *
 * R/S/T/U -> BE BAK, keutamaan R > S > T > U.
 * ================================================================
 */
function updateLowerSecondaryGrades(token) {
  requireAdmin_(token);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const target = ss.getSheetByName('PROFIL MURID');
  if (!target) throw new Error('Tab PROFIL MURID tidak dijumpai.');

  const source = findLowerSecondarySheet_(ss);
  if (!source) {
    throw new Error(
      'Tab markah Menengah Rendah tidak dijumpai. ' +
      'Pastikan struktur C=KAD PENGENALAN dan F:U mengandungi BM, BI, PI, PM(R), SEJ, GEO, SNS, MAT, RBT, ASK, PSV, PSM, BJ, BC, BA, BG.'
    );
  }

  const sourceLastRow = source.getLastRow();
  const targetLastRow = target.getLastRow();

  if (sourceLastRow < 2) {
    return {
      ok: true, updated: 0, notFound: 0, skipped: 0,
      sourceRows: 0, sourceSheet: source.getName(),
      message: 'Tiada rekod dalam ' + source.getName() + '.'
    };
  }
  if (targetLastRow < 2) {
    throw new Error('Tiada data murid dalam PROFIL MURID.');
  }

  // C:U, 19 kolum.
  const sourceValues = source.getRange(1, 3, sourceLastRow, 19).getDisplayValues();

  // PROFIL MURID: C = NO KP, AX:BN = 17 kolum.
  const targetKPValues = target.getRange(2, 3, targetLastRow - 1, 1).getDisplayValues();
  const targetGradeValues = target.getRange(2, 50, targetLastRow - 1, 17).getValues();

  const targetByKP = new Map();
  targetKPValues.forEach((row, idx) => {
    const kp = normalizeKP12_(row[0]);
    if (kp && !targetByKP.has(kp)) targetByKP.set(kp, idx);
  });

  // Dalam C:U: C=0, F=3 ... U=18.
  // Dalam AX:BN: AX=0 ... BN=16.
  const mappings = [
    [3, 0],   // F BM    -> AX
    [4, 1],   // G BI    -> AY
    [5, 6],   // H PI    -> BD
    [6, 14],  // I PM(R) -> BL
    [7, 2],   // J SEJ   -> AZ
    [8, 3],   // K GEO   -> BA
    [9, 5],   // L SNS   -> BC
    [10, 4],  // M MAT   -> BB
    [11, 8],  // N RBT   -> BF
    [12, 9],  // O ASK   -> BG
    [13, 15], // P PSV   -> BM
    [14, 16]  // Q PSM   -> BN
  ];

  let updated = 0;
  let notFound = 0;
  let skipped = 0;
  const missing = [];

  for (let r = 1; r < sourceValues.length; r++) {
    const row = sourceValues[r];
    const kp = normalizeKP12_(row[0]); // C

    if (!kp) {
      skipped++;
      continue;
    }

    const targetIdx = targetByKP.get(kp);
    if (targetIdx === undefined) {
      notFound++;
      if (missing.length < 100) missing.push(kp);
      continue;
    }

    let changed = false;

    mappings.forEach(([sourceIdx, targetGradeIdx]) => {
      const value = String(row[sourceIdx] ?? '').trim();
      if (value !== '') {
        targetGradeValues[targetIdx][targetGradeIdx] = value;
        changed = true;
      }
    });

    // R/S/T/U -> BE BAK (P1), first non-empty wins.
    const bakCandidates = [row[15], row[16], row[17], row[18]]
      .map(v => String(v ?? '').trim())
      .filter(v => v !== '');

    if (bakCandidates.length) {
      targetGradeValues[targetIdx][7] = bakCandidates[0]; // BE
      changed = true;
    }

    if (changed) updated++;
  }

  target.getRange(2, 50, targetLastRow - 1, 17).setValues(targetGradeValues);

  return {
    ok: true,
    updated: updated,
    notFound: notFound,
    skipped: skipped,
    sourceRows: sourceLastRow - 1,
    targetRows: targetLastRow - 1,
    sourceSheet: source.getName(),
    missingNoKP: missing,
    mapping: {
      'F BM': 'AX BM (P1)',
      'G BI': 'AY BI (P1)',
      'H PI': 'BD PAI (P1)',
      'I PM(R)': 'BL MRL (P1)',
      'J SEJ': 'AZ SEJ (P1)',
      'K GEO': 'BA GEO (P1)',
      'L SNS': 'BC SN (P1)',
      'M MAT': 'BB MM (P1)',
      'N RBT': 'BF RBT (P1)',
      'O ASK': 'BG ASK (P1)',
      'P PSV': 'BM PQS (P1)',
      'Q PSM': 'BN PSI (P1)',
      'R/S/T/U': 'BE BAK (P1)'
    },
    message: 'Kemaskini gred Menengah Rendah selesai.'
  };
}

/**
 * Cari tab Menengah Rendah berdasarkan nama biasa atau struktur header.
 */
function findLowerSecondarySheet_(ss) {
  const preferredNames = [
    'Markah_MenRendah',
    'Markah_MenengahRendah',
    'Markah Menengah Rendah',
    'Menengah_Rendah',
    'Menengah Rendah'
  ];

  for (const name of preferredNames) {
    const sh = ss.getSheetByName(name);
    if (sh && lowerSecondaryHeaderMatch_(sh)) return sh;
  }

  for (const sh of ss.getSheets()) {
    if (sh.getName() === 'PROFIL MURID') continue;
    if (lowerSecondaryHeaderMatch_(sh)) return sh;
  }

  return null;
}

function lowerSecondaryHeaderMatch_(sh) {
  if (sh.getLastRow() < 1 || sh.getLastColumn() < 21) return false;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn())
    .getDisplayValues()[0]
    .map(v => String(v || '').trim().toUpperCase());

  const expected = ['BM','BI','PI','PM(R)','SEJ','GEO','SNS','MAT','RBT','ASK','PSV','PSM','BJ','BC','BA','BG'];
  const expectedCols = [5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20];

  const kp = headers[2] || '';
  if (!/KAD PENGENALAN|NO\.?\s*KP|NO KAD PENGENALAN/.test(kp)) return false;

  return expected.every((label, i) => headers[expectedCols[i]] === label);
}


function getSessionUser(token) {
  const session = validateSession_(token);
  return { ok: true, user: { username: session.username }, permissions: session.permissions };
}

function validateSession_(token) {
  const t = String(token || '').trim();
  if (!t) throw new Error('Sesi log masuk tidak sah. Sila log masuk semula.');
  const raw = CacheService.getScriptCache().get(AUTH_CACHE_PREFIX + t);
  if (!raw) throw new Error('Sesi telah tamat. Sila log masuk semula.');
  const session = JSON.parse(raw);
  session.token = t;
  return session;
}

function requireModuleAccess_(token, moduleNo) {
  const session = validateSession_(token);
  const key = 'Modul' + Number(moduleNo);
  if (String(session.permissions[key] || 0) !== '1') {
    const moduleName = MODULE_TAB_MAP[moduleNo] || ('Modul ' + moduleNo);
    throw new Error('AKSES DISEKAT: Anda tidak dibenarkan mengakses ' + moduleName + '.');
  }
  return session;
}

function logoutUser(token) {
  const t = String(token || '').trim();
  if (t) CacheService.getScriptCache().remove(AUTH_CACHE_PREFIX + t);
  return { ok: true };
}

function getDatabaseSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = DATABASE_SHEET
    ? ss.getSheetByName(DATABASE_SHEET)
    : ss.getSheets()[0];

  if (!sheet) throw new Error('Sheet database tidak dijumpai: ' + DATABASE_SHEET);
  return sheet;
}

function readDatabase_() {
  const sheet = getDatabaseSheet_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return { headers: [], rows: [] };

  // Hanya A:BN. Elak getDataRange() membaca kolum/ruang yang tidak diperlukan.
  const values = sheet.getRange(1, 1, lastRow, DATABASE_COLUMNS).getDisplayValues();
  const headers = values[0].map(h => String(h).trim());

  return { headers, rows: values.slice(1) };
}

function makeColumnMap_(headers) {
  const map = {};
  headers.forEach((h, i) => { if (h) map[h] = i; });
  return map;
}

function projectRows_(raw, wantedHeaders) {
  const map = makeColumnMap_(raw.headers);
  const indices = wantedHeaders.map(h => map[h]);

  const rows = [];
  for (let r = 0; r < raw.rows.length; r++) {
    const row = raw.rows[r];
    const nameIndex = map['NAMA MURID'];
    if (nameIndex == null || !row[nameIndex]) continue;
    rows.push(indices.map(i => i == null ? '' : (row[i] || '')));
  }
  return { headers: wantedHeaders, rows };
}

/**
 * ============================================================
 * V8 FAST SNAPSHOT ARCHITECTURE
 *
 * getCoreData() TIDAK membaca Google Sheet jika snapshot sudah ada.
 * Dashboard startup hanya membaca snapshot A:I.
 *
 * Jalankan refreshCoreSnapshot() sekali selepas setup/deploy.
 * Kemudian dashboard tidak perlu membaca Sheet pada setiap open.
 * ============================================================
 */
const V8_CACHE_KEY = 'V8_CORE_CACHE';
const V8_META_KEY = 'V8_CORE_META';
const V8_PROP_PREFIX = 'V8_CORE_';
const V8_CHUNK_SIZE = 8000; // PropertiesService value limit ~9KB
const V8_CACHE_TTL = 21600; // 6 jam

// Susunan sebenar PROFIL MURID!A:I
// A=BIL, B=NAMA MURID, C=NO KP, D=JANTINA, E=TINGKATAN,
// F=KELAS, G=RUMAH, H=MENTOR, I=DORM.
// MESTI sepadan dengan susunan values yang dibaca dalam refreshCoreSnapshot().
const CORE_HEADERS = [
  'BIL','NAMA MURID','NO KP','JANTINA','TINGKATAN',
  'KELAS','RUMAH','MENTOR','DORM'
];

/**
 * Dashboard startup.
 * 1. Script Cache
 * 2. Persistent Script Properties snapshot
 * 3. Hanya jika kedua-duanya tiada -> baca Google Sheet dan bina snapshot.
 */
function getCoreData(token, moduleNo) {
  validateSession_(token);
  if (moduleNo) requireModuleAccess_(token, moduleNo);
  const cache = CacheService.getScriptCache();

  // Fast path #1: Script Cache
  const cached = cache.get(V8_CACHE_KEY);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {}
  }

  // Fast path #2: persistent snapshot
  const snapshot = readCoreSnapshot_();
  if (snapshot) {
    try {
      cache.put(V8_CACHE_KEY, JSON.stringify(snapshot), V8_CACHE_TTL);
    } catch (e) {}
    return snapshot;
  }

  // Cold start sahaja: baca Sheet dan bina snapshot.
  return refreshCoreSnapshot();
}

/**
 * Baca A:I dari PROFIL MURID dan simpan snapshot secara persistent.
 * Jalankan fungsi ini selepas data database berubah jika mahu data terkini.
 */
function refreshCoreSnapshot() {
  const sheet = getDatabaseSheet_();
  const lastRow = sheet.getLastRow();

  let payload = { headers: CORE_HEADERS, rows: [] };

  if (lastRow >= 2) {
    // Hanya A:I. Tiada A:BN.
    const values = sheet.getRange(1, 1, lastRow, 9).getValues();

    const rows = [];
    for (let r = 1; r < values.length; r++) {
      // NAMA MURID = kolum B
      if (!values[r][1]) continue;
      rows.push(values[r].map(v => v == null ? '' : String(v)));
    }

    payload.rows = rows;
  }

  writeCoreSnapshot_(payload);

  return payload;
}

/**
 * Simpan snapshot ke Script Properties dalam chunks.
 * Script Properties kekal walaupun CacheService expired.
 */
function writeCoreSnapshot_(payload) {
  const props = PropertiesService.getScriptProperties();
  const json = JSON.stringify(payload);
  const count = Math.ceil(json.length / V8_CHUNK_SIZE);

  // Buang snapshot lama dahulu.
  const oldMeta = props.getProperty(V8_META_KEY);
  if (oldMeta) {
    try {
      const old = JSON.parse(oldMeta);
      const oldKeys = [];
      for (let i = 0; i < (old.count || 0); i++) {
        oldKeys.push(V8_PROP_PREFIX + i);
      }
      if (oldKeys.length) props.deleteProperties(oldKeys);
    } catch (e) {}
  }

  // PropertiesService tidak menerima payload besar sebagai satu value.
  const entries = {};
  for (let i = 0; i < count; i++) {
    entries[V8_PROP_PREFIX + i] = json.substring(
      i * V8_CHUNK_SIZE,
      (i + 1) * V8_CHUNK_SIZE
    );
  }

  entries[V8_META_KEY] = JSON.stringify({
    count: count,
    updated: new Date().toISOString(),
    rowCount: payload.rows.length
  });

  props.setProperties(entries, false);

  // Fast cache untuk request berikutnya.
  const cache = CacheService.getScriptCache();
  try {
    cache.put(V8_CACHE_KEY, json, V8_CACHE_TTL);
  } catch (e) {}
}

/** Baca snapshot persistent daripada Script Properties. */
function readCoreSnapshot_() {
  const props = PropertiesService.getScriptProperties();
  const metaText = props.getProperty(V8_META_KEY);
  if (!metaText) return null;

  try {
    const meta = JSON.parse(metaText);
    const parts = [];

    for (let i = 0; i < meta.count; i++) {
      const part = props.getProperty(V8_PROP_PREFIX + i);
      if (part == null) return null;
      parts.push(part);
    }

    return JSON.parse(parts.join(''));
  } catch (e) {
    return null;
  }
}

/** Maklumat status snapshot untuk troubleshooting. */
function getSnapshotStatus() {
  const props = PropertiesService.getScriptProperties();
  const meta = props.getProperty(V8_META_KEY);
  if (!meta) {
    return { ready: false, message: 'Snapshot V8 belum dibuat.' };
  }

  try {
    const m = JSON.parse(meta);
    return {
      ready: true,
      rows: m.rowCount || 0,
      chunks: m.count || 0,
      updated: m.updated || ''
    };
  } catch (e) {
    return { ready: false, message: 'Snapshot V8 rosak.' };
  }
}

/**
 * Setup sekali sahaja.
 * Run fungsi ini selepas paste Code.gs V8.
 */
function setupV8() {
  const payload = refreshCoreSnapshot();
  return {
    ok: true,
    sheet: DATABASE_SHEET,
    rows: payload.rows.length,
    message: 'V8 snapshot berjaya dibina. Dashboard startup kini menggunakan snapshot.'
  };
}

/** Padam snapshot V8 jika mahu bina semula dari Sheet. */
function clearV8Snapshot() {
  const props = PropertiesService.getScriptProperties();
  const metaText = props.getProperty(V8_META_KEY);

  if (metaText) {
    try {
      const meta = JSON.parse(metaText);
      const keys = [V8_META_KEY];
      for (let i = 0; i < (meta.count || 0); i++) {
        keys.push(V8_PROP_PREFIX + i);
      }
      props.deleteProperties(keys);
    } catch (e) {
      props.deleteProperty(V8_META_KEY);
    }
  }

  CacheService.getScriptCache().remove(V8_CACHE_KEY);
  return 'Snapshot V8 telah dibersihkan.';
}

/**
 * SUMMARY: dipanggil selepas dashboard asas sudah dipaparkan.
 * Hanya A:AW (49 kolum), bukan keseluruhan A:BN.
 */
function getSummaryData(token) {
  validateSession_(token);
  const cache = CacheService.getScriptCache();
  const key = 'SUMMARY_V6';
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const sheet = getDatabaseSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { mindaSummary: {phq:{}, gad:{}}, summary:{} };
  }

  // Semua field summary yang diperlukan berada dalam A:AW.
  const SUMMARY_COLUMNS = 49; // A:AW
  const values = sheet.getRange(1, 1, lastRow, SUMMARY_COLUMNS).getDisplayValues();
  const headers = values[0].map(h => String(h).trim());
  const map = makeColumnMap_(headers);
  const rows = values.slice(1).filter(r => map['NAMA MURID'] != null && r[map['NAMA MURID']]);

  const get = (row, field) => map[field] == null ? '' : String(row[map[field]] || '').trim();
  const rumahCounts = {}, rumahNames = {};
  const gbCounts = {}, gbNames = {};
  const pwCounts = {}, pwNames = {};
  const mkCounts = {}, mkNames = {};
  const phq = {}, gad = {};
  const tpFields = ['AUTONOMI','KREATIF','AGRESIF','EKSTROVERT','PENCAPAIAN','KEPELBAGAIAN','INTELEKTUAL','KEPEMIMPINAN','STRUKTUR','RESILIENCE','MENOLONG','ANALITIKAL','KRITIK DIRI','WAWASAN','KETELUSAN'];
  const kpFields = ['VLM','VLI','LM','VR','MZK','NAT','INTRA','INTER','KINES','EKS'];
  const mkFields = ['REALISTIK','INVESTIGATIF','ARTISTIK','SOSIAL','ENTERPRISING','KONVENSIONAL'];
  const tpSum = {}, tpN = {}, kpSum = {}, kpN = {};

  const addGroup = (counts, names, key, name) => {
    key = key || 'TIADA';
    counts[key] = (counts[key] || 0) + 1;
    if (!names[key]) names[key] = [];
    names[key].push(name);
  };

  rows.forEach(row => {
    const name = get(row,'NAMA MURID');
    addGroup(rumahCounts, rumahNames, get(row,'RUMAH'), name);
    addGroup(gbCounts, gbNames, get(row,'GAYA BELAJAR'), name);
    addGroup(pwCounts, pwNames, get(row,'PERSONALITI WARNA'), name);

    mkFields.forEach(f => {
      const v = Number(get(row,f));
      if (!isNaN(v) && v >= 70) addGroup(mkCounts, mkNames, f, name);
    });

    tpFields.forEach(f => {
      const v = Number(get(row,f));
      if (!isNaN(v)) { tpSum[f]=(tpSum[f]||0)+v; tpN[f]=(tpN[f]||0)+1; }
    });
    kpFields.forEach(f => {
      const v = Number(get(row,f));
      if (!isNaN(v)) { kpSum[f]=(kpSum[f]||0)+v; kpN[f]=(kpN[f]||0)+1; }
    });

    const p = get(row,'TAHAP PHQ-9');
    const g = get(row,'TAHAP (GAD-7)');
    if (p) phq[p]=(phq[p]||0)+1;
    if (g) gad[g]=(gad[g]||0)+1;
  });

  const tpAvg = {}, kpAvg = {};
  tpFields.forEach(f => tpAvg[f] = tpN[f] ? Math.round(tpSum[f]/tpN[f]) : 0);
  kpFields.forEach(f => kpAvg[f] = kpN[f] ? Math.round(kpSum[f]/kpN[f]) : 0);

  const payload = {
    mindaSummary: { phq, gad },
    summary: { rumahCounts, rumahNames, gbCounts, gbNames, pwCounts, pwNames, mkCounts, mkNames, tpAvg, kpAvg }
  };

  const json = JSON.stringify(payload);
  if (json.length < 95000) cache.put(key, json, 600);
  return payload;
}

/**
 * FULL DATA
 * Dipanggil hanya apabila tab psikometrik / akademik diperlukan.
 */
function getFullDashboardData(token, moduleNo) {
  requireModuleAccess_(token, moduleNo);
  return getDashboardData();
}


function getDashboardData() {
  const cache = CacheService.getScriptCache();
  const metaKey = 'DASHBOARD_V3_META';
  const meta = cache.get(metaKey);

  // Cache dipecahkan kepada beberapa bahagian supaya payload besar
  // masih boleh dicache. Ini menjadikan reload seterusnya jauh lebih laju.
  if (meta) {
    try {
      const info = JSON.parse(meta);
      const chunks = [];
      for (let i = 0; i < info.count; i++) {
        const part = cache.get('DASHBOARD_V3_' + i);
        if (!part) throw new Error('Cache chunk missing');
        chunks.push(part);
      }
      return JSON.parse(chunks.join(''));
    } catch (e) {
      // Cache rosak/tamat tempoh -> bina semula dari Sheet.
    }
  }

  const raw = readDatabase_();
  const payload = projectRows_(raw, DASHBOARD_HEADERS);

  const map = makeColumnMap_(raw.headers);
  const phqIndex = map['TAHAP PHQ-9'];
  const gadIndex = map['TAHAP (GAD-7)'];
  const nameIndex = map['NAMA MURID'];
  const phq = {};
  const gad = {};

  raw.rows.forEach(function(row) {
    if (nameIndex == null || !row[nameIndex]) return;
    const p = phqIndex == null ? '' : String(row[phqIndex] || '').trim();
    const g = gadIndex == null ? '' : String(row[gadIndex] || '').trim();
    if (p) phq[p] = (phq[p] || 0) + 1;
    if (g) gad[g] = (gad[g] || 0) + 1;
  });

  payload.mindaSummary = { phq: phq, gad: gad };

  const json = JSON.stringify(payload);

  // Cache chunks ~80KB setiap satu.
  const chunkSize = 80000;
  const count = Math.ceil(json.length / chunkSize);
  const writes = {};

  for (let i = 0; i < count; i++) {
    writes['DASHBOARD_V3_' + i] = json.substring(i * chunkSize, (i + 1) * chunkSize);
  }
  writes[metaKey] = JSON.stringify({ count: count });

  // CacheService lebih stabil jika ditulis satu persatu.
  Object.keys(writes).forEach(function(key) {
    cache.put(key, writes[key], 600); // 10 minit
  });

  return payload;
}

/**
 * Clear temporary caches.
 * Snapshot V8 PERSISTENT tidak dipadamkan.
 * Jika data Sheet berubah, gunakan refreshCoreSnapshot().
 */
function clearDashboardCache() {
  const cache = CacheService.getScriptCache();
  const keys = [
    V8_CACHE_KEY,
    'SUMMARY_V6',
    'CORE_V4',
    'CORE_V5',
    'CORE_V6',
    'DASHBOARD_V3_META'
  ];

  // Baca metadata dahulu untuk tahu bilangan chunk cache Dashboard V3.
  const metaText = cache.get('DASHBOARD_V3_META');
  if (metaText) {
    try {
      const meta = JSON.parse(metaText);
      for (let i = 0; i < (meta.count || 0); i++) {
        keys.push('DASHBOARD_V3_' + i);
      }
    } catch (e) {}
  }

  cache.removeAll(keys);
  return 'Cache sementara dashboard dibersihkan. Snapshot V8 persistent dikekalkan.';
}

/**
 * Bina semula snapshot A:I daripada PROFIL MURID.
 * Jalankan sekali selepas pembetulan CORE_HEADERS supaya data Mentor/Dorm
 * tidak lagi menggunakan pemetaan lama.
 */
function rebuildCoreSnapshotNow() {
  CacheService.getScriptCache().remove(V8_CACHE_KEY);
  const payload = refreshCoreSnapshot();
  clearDashboardCache();
  return {
    ok: true,
    rows: payload.rows.length,
    headers: payload.headers,
    message: 'Snapshot A:I dibina semula menggunakan pemetaan PROFIL MURID yang betul.'
  };
}


/**
 * Data Akademik untuk MODUL AKADEMIK sahaja.
 * Filter murid kekal daripada PROFIL MURID:
 *   B = NAMA MURID, C = NO KP, E = TINGKATAN, F = KELAS.
 * Gred diambil terus daripada Markah_MenBawah / Markah_MenAtas.
 * Untuk setiap murid, rekod peperiksaan paling terkini dipilih berdasarkan
 * TAHUN kemudian kedudukan baris.
 */
/** Akademik: semua rekod peperiksaan murid yang dipilih. Sumber terus Markah_MenBawah/MenAtas. */
function getAkademikExamsByStudent(token, noKp, studentName, tingkatan, kelas) {
  requireModuleAccess_(token, 7);

  // Modul Akademik tidak menghantar NO KP dalam allData untuk loading awal.
  // Jika NO KP kosong/000000000000, cari semula NO KP sebenar terus daripada
  // PROFIL MURID berdasarkan Nama + Tingkatan + Kelas.
  let kp = normalizeNoKp_(noKp);
  const kpMissing = !kp || /^0{12}$/.test(kp);

  if (kpMissing && String(studentName || '').trim()) {
    const sheet = getDatabaseSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const name = String(studentName || '').trim();
      const targetTing = String(tingkatan || '').replace(/\D/g, '');
      const targetKelas = String(kelas || '').trim().toUpperCase();

      const nameRange = sheet.getRange(2, 2, lastRow - 1, 1); // B = NAMA MURID
      const finder = nameRange.createTextFinder(name)
        .matchEntireCell(true)
        .matchCase(false);

      let cell = finder.findNext();
      while (cell) {
        const row = cell.getRow();
        const vals = sheet.getRange(row, 3, 1, 4).getDisplayValues()[0]; // C:F
        const rowKp = normalizeNoKp_(vals[0]);
        const rowTing = String(vals[2] || '').replace(/\D/g, '');
        const rowKelas = String(vals[3] || '').trim().toUpperCase();

        if (rowKp && !/^0{12}$/.test(rowKp) &&
            (!targetTing || rowTing === targetTing) &&
            (!targetKelas || rowKelas === targetKelas)) {
          kp = rowKp;
          break;
        }
        cell = finder.findNext(cell);
      }
    }
  }

  if (!kp || /^0{12}$/.test(kp)) {
    return { menBawah: [], menAtas: [], noKp: '' };
  }

  const result = getAcademicP1_(kp);
  result.noKp = kp;
  return result;
}

function getAkademikModuleData(token, tingkatan, kelas, murid) {
  requireModuleAccess_(token, 7);

  const clean = function(v) { return String(v == null ? '' : v).trim(); };
  const targetTing = clean(tingkatan);
  const targetKelas = clean(kelas);
  const targetMurid = clean(murid);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const profileSheet = ss.getSheetByName('PROFIL MURID');
  if (!profileSheet || profileSheet.getLastRow() < 2) return [];

  // B:F = Nama, No KP, Jantina, Tingkatan, Kelas
  const profileRows = profileSheet.getRange(2, 2, profileSheet.getLastRow() - 1, 5).getDisplayValues();
  const students = [];

  profileRows.forEach(function(r) {
    const nama = clean(r[0]);       // B
    const noKp = normalizeNoKp_(r[1]); // C
    const ting = clean(r[3]);       // E
    const kls = clean(r[4]);        // F

    if (!nama || !noKp) return;
    if (targetTing && targetTing !== 'SEMUA' && ting !== targetTing) return;
    if (targetKelas && targetKelas !== 'SEMUA' && kls !== targetKelas) return;
    if (targetMurid && targetMurid !== 'SEMUA' && nama !== targetMurid) return;

    students.push({ nama: nama, noKp: noKp, ting: ting, kelas: kls });
  });

  if (!students.length) return [];

  const wanted = {};
  students.forEach(function(s) { wanted[s.noKp] = s; });

  // Paparan P1 kekal seperti struktur modul sedia ada.
  const emptyGrades = function() {
    return {
      'BM (P1)': '', 'BI (P1)': '', 'SEJ (P1)': '', 'GEO (P1)': '',
      'MM (P1)': '', 'SN (P1)': '', 'PAI (P1)': '', 'BAK (P1)': '',
      'RBT (P1)': '', 'ASK (P1)': '', 'FIZ (P1)': '', 'KIM (P1)': '',
      'BIO (P1)': '', 'MT (P1)': '', 'MRL (P1)': '', 'PQS (P1)': '',
      'PSI (P1)': ''
    };
  };

  // Source column -> P1 label.
  const lowerMap = {
    'BM':'BM (P1)', 'BI':'BI (P1)', 'SEJ':'SEJ (P1)', 'GEO':'GEO (P1)',
    'MAT':'MM (P1)', 'SNS':'SN (P1)', 'PI':'PAI (P1)', 'PM(R)':'MRL (P1)',
    'RBT':'RBT (P1)', 'ASK':'ASK (P1)', 'PSV':'PQS (P1)', 'PSM':'PSI (P1)',
    'BG':'BAK (P1)', 'BA':'BAK (P1)', 'BC':'BAK (P1)', 'BJ':'BAK (P1)'
  };

  const upperMap = {
    'BM':'BM (P1)', 'BI':'BI (P1)', 'SEJ':'SEJ (P1)', 'MAT':'MM (P1)',
    'MT':'MT (P1)', 'PI':'PAI (P1)', 'PM(A)':'MRL (P1)', 'FIZ':'FIZ (P1)',
    'KIM':'KIM (P1)', 'BIO':'BIO (P1)', 'BJ':'BAK (P1)', 'BG':'BAK (P1)',
    'BC':'BAK (P1)', 'BA':'BAK (P1)'
  };

  const lowerCols = {
    'BM':9, 'BI':11, 'PI':13, 'PM(R)':15, 'SEJ':17, 'GEO':19, 'SNS':21,
    'MAT':23, 'RBT':25, 'ASK':27, 'PSV':29, 'PSM':31, 'BG':33, 'BA':35,
    'BC':37, 'BJ':39, 'PJK':41
  };
  const upperCols = {
    'BM':9, 'BI':11, 'PI':13, 'PM(A)':15, 'SEJ':17, 'MAT':19, 'MT':21,
    'FIZ':23, 'KIM':25, 'BIO':27, 'BJ':29, 'BG':31, 'BC':33, 'BA':35
  };

  const latest = {};

  function yearNumber(v) {
    const m = clean(v).match(/\d{4}/);
    return m ? parseInt(m[0], 10) : 0;
  }

  function readMarks(sheetName, cols, map, isUpper) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;

    const lastRow = sh.getLastRow();
    const values = sh.getRange(2, 1, lastRow - 1, 41).getDisplayValues();

    values.forEach(function(row, i) {
      const kp = normalizeNoKp_(row[4]); // E = KAD PENGENALAN
      if (!kp || !wanted[kp]) return;
      const tingValue = clean(wanted[kp].ting).replace(/\D/g, '');
      if (isUpper && tingValue && !['4','5'].includes(tingValue)) return;
      if (!isUpper && tingValue && !['1','2','3'].includes(tingValue)) return;

      const rec = {
        tahun: clean(row[0]),
        exam: clean(row[1]),
        rowNo: i + 2,
        grades: emptyGrades()
      };

      Object.keys(cols).forEach(function(subject) {
        const target = map[subject];
        if (!target) return;
        const grade = clean(row[cols[subject] - 1]);
        if (grade) rec.grades[target] = grade;
      });

      const old = latest[kp];
      const newYear = yearNumber(rec.tahun);
      const oldYear = old ? yearNumber(old.tahun) : -1;

      if (!old || newYear > oldYear || (newYear === oldYear && rec.rowNo > old.rowNo)) {
        latest[kp] = rec;
      }
    });
  }

  // OPTIMASI: hanya baca tab peperiksaan yang berkaitan dengan tingkatan.
  // Ini mengelakkan setiap permintaan membaca kedua-dua tab sekali gus.
  // Tingkatan 1-3 -> Menengah Rendah; Tingkatan 4-5 -> Menengah Atas.
  const tingDigit = targetTing.replace(/\D/g, '');
  if (['1','2','3'].includes(tingDigit)) {
    readMarks('Markah_MenBawah', lowerCols, lowerMap, false);
  } else if (['4','5'].includes(tingDigit)) {
    readMarks('Markah_MenAtas', upperCols, upperMap, true);
  } else {
    // Hanya untuk pilihan "Semua Tingkatan".
    readMarks('Markah_MenBawah', lowerCols, lowerMap, false);
    readMarks('Markah_MenAtas', upperCols, upperMap, true);
  }

  return students.map(function(s) {
    const rec = latest[s.noKp];
    return {
      nama: s.nama,
      noKp: s.noKp,
      tingkatan: s.ting,
      kelas: s.kelas,
      tahun: rec ? rec.tahun : '',
      exam: rec ? rec.exam : '',
      grades: rec ? rec.grades : emptyGrades()
    };
  });
}

/**
 * Ambil satu profil penuh apabila user memilih nama murid.
 * Optimized: cari nama pada kolum B dahulu, kemudian baca SATU row A:BN.
 * Tidak lagi membaca seluruh A:BN setiap kali profil dipilih.
 */
function normalizeNoKp_(value) {
  const digits = String(value == null ? '' : value).replace(/\D/g, '');
  return digits ? digits.padStart(12, '0').slice(-12) : '';
}

function getImageSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(IMAGE_SHEET);
  if (!sheet) return null;
  return sheet;
}

/**
 * Ambil gambar murid daripada tab RegGambar.
 * A = Nama Fail, B = ID Fail Gambar, C = No. KP.
 * ID Fail Gambar dianggap sebagai Google Drive File ID.
 */
function getStudentImage_(noKp) {
  const targetKp = normalizeNoKp_(noKp);
  if (!targetKp) return { found: false, url: '', fileId: '' };

  const sheet = getImageSheet_();
  if (!sheet || sheet.getLastRow() < 2) {
    return { found: false, url: '', fileId: '' };
  }

  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const fileName = String(values[i][0] || '').trim();
    const fileId = String(values[i][1] || '').trim();
    const rowKp = normalizeNoKp_(values[i][2]);
    if (rowKp && rowKp === targetKp && fileId) {
      return {
        found: true,
        fileName: fileName,
        fileId: fileId,
        url: 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w600'
      };
    }
  }

  return { found: false, url: '', fileId: '' };
}


/**
 * Ambil maklumat waris murid daripada tab Daftar menggunakan NO KP murid.
 * Daftar:
 * D  = No. KP murid
 * O  = No. KP Penjaga 1
 * P  = Nama Penjaga 1
 * Q  = Hubungan Penjaga 1
 * U  = Pekerjaan Penjaga 1
 * V  = No. Tel Bimbit Penjaga 1
 * W  = Pendapatan Penjaga 1
 * X  = Jumlah Tanggungan
 * Y  = No. KP Penjaga 2
 * Z  = Nama Penjaga 2
 * AA = Hubungan Penjaga 2
 * AE = Pekerjaan Penjaga 2
 * AF = No. Tel Bimbit Penjaga 2
 *
 * Warganegara, Kaum dan Agama sengaja tidak dipulangkan.
 */
function getStudentWaris_(noKp) {
  const targetKp = normalizeNoKp_(noKp);
  const empty = {
    found: false,
    penjaga1: null,
    penjaga2: null
  };
  if (!targetKp) return empty;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Daftar');
  if (!sheet || sheet.getLastRow() < 2) return empty;

  // D sahaja untuk carian NO KP murid.
  const lastRow = sheet.getLastRow();
  const kpValues = sheet.getRange(2, 4, lastRow - 1, 1).getDisplayValues();
  let rowNumber = 0;

  for (let i = 0; i < kpValues.length; i++) {
    if (normalizeNoKp_(kpValues[i][0]) === targetKp) {
      rowNumber = i + 2;
      break;
    }
  }
  if (!rowNumber) return empty;

  // B:AF (31 kolum) — cukup untuk semua medan waris yang diperlukan.
  const row = sheet.getRange(rowNumber, 2, 1, 31).getDisplayValues()[0];
  const val = function(index) { return String(row[index] == null ? '' : row[index]).trim(); };
  const upper = function(value) { return String(value || '').trim().toUpperCase(); };

  const penjaga1 = {
    noKp: normalizeNoKp_(val(13)),       // O
    nama: upper(val(14)),                // P
    hubungan: upper(val(15)),            // Q
    pekerjaan: val(19),                  // U
    telefon: val(20),                    // V
    pendapatan: val(21),                 // W
    jumlahTanggungan: val(22)            // X
  };

  const penjaga2 = {
    noKp: normalizeNoKp_(val(23)),       // Y
    nama: upper(val(24)),                // Z
    hubungan: upper(val(25)),            // AA
    pekerjaan: val(29),                  // AE
    telefon: val(30)                     // AF
  };

  const has1 = Boolean(penjaga1.nama || penjaga1.noKp || penjaga1.hubungan || penjaga1.pekerjaan || penjaga1.telefon);
  const has2 = Boolean(penjaga2.nama || penjaga2.noKp || penjaga2.hubungan || penjaga2.pekerjaan || penjaga2.telefon);

  return {
    found: has1 || has2,
    penjaga1: has1 ? penjaga1 : null,
    penjaga2: has2 ? penjaga2 : null
  };
}

/**
 * Ambil alamat rumah murid daripada tab Daftar, kolum N.
 * Key: NO KP murid = Daftar!D:D.
 */
function getStudentAddress_(noKp) {
  const targetKp = normalizeNoKp_(noKp);
  if (!targetKp) return '';

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Daftar');
  if (!sheet || sheet.getLastRow() < 2) return '';

  const lastRow = sheet.getLastRow();
  const kpValues = sheet.getRange(2, 4, lastRow - 1, 1).getDisplayValues(); // D = No. KP
  let rowNumber = 0;

  for (let i = 0; i < kpValues.length; i++) {
    if (normalizeNoKp_(kpValues[i][0]) === targetKp) {
      rowNumber = i + 2;
      break;
    }
  }

  if (!rowNumber) return '';
  return String(sheet.getRange(rowNumber, 14).getDisplayValue() || '').trim(); // N = Alamat Rumah
}

/**
 * Prestasi Akademik (P1) untuk PROFIL MURID.
 * Sumber: Markah_MenBawah dan Markah_MenAtas.
 * Key: NO KP PROFIL MURID <-> KAD PENGENALAN.
 * Semua No KP dinormalisasi kepada 12 digit.
 * Hanya TAHUN, EXAM dan gred subjek dipulangkan.
 */
function getAcademicP1_(noKp) {
  const targetKp = normalizeNoKp_(noKp);
  const result = { menBawah: [], menAtas: [] };
  if (!targetKp) return result;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Struktur kolum adalah tetap berdasarkan sheet semasa.
  // Index di sini ialah index 0-based selepas membaca A:AO.
  const CONFIG = {
    Markah_MenBawah: {
      subjects: [
        ['BM', 9], ['BI', 11], ['PI', 13], ['PM(R)', 15], ['SEJ', 17],
        ['GEO', 19], ['SNS', 21], ['MAT', 23], ['RBT', 25], ['ASK', 27],
        ['PSV', 29], ['PSM', 31], ['BG', 33], ['BA', 35], ['BC', 37],
        ['BJ', 39], ['PJK', 41]
      ],
      output: 'menBawah'
    },
    Markah_MenAtas: {
      subjects: [
        ['BM', 9], ['BI', 11], ['PI', 13], ['PM(A)', 15], ['SEJ', 17],
        ['MAT', 19], ['MT', 21], ['FIZ', 23], ['KIM', 25], ['BIO', 27],
        ['BJ', 29], ['BG', 31], ['BC', 33], ['BA', 35]
      ],
      output: 'menAtas'
    }
  };

  /**
   * Cari KP terus pada kolum E sahaja.
   * Tidak membaca keseluruhan sheet.
   * Jika KP dalam sheet tidak mempunyai 0 di hadapan, cuba juga versi tanpa
   * leading zero kerana semua KP dinormalisasi kepada 12 digit.
   */
  const findMatchingRows = function(sheet, target) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const kpRange = sheet.getRange(2, 5, lastRow - 1, 1); // E:E
    const matches = [];

    const searchValues = [target];
    const unpadded = String(target).replace(/^0+/, '');
    if (unpadded && unpadded !== target) searchValues.push(unpadded);

    searchValues.forEach(function(searchValue) {
      const found = kpRange.createTextFinder(searchValue)
        .matchEntireCell(true)
        .matchCase(false)
        .findAll();
      found.forEach(function(cell) {
        const row = cell.getRow();
        if (matches.indexOf(row) === -1) matches.push(row);
      });
    });

    matches.sort(function(a, b) { return a - b; });
    return matches;
  };

  const readExamSheet = function(sheetName, config) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const rowsFound = findMatchingRows(sheet, targetKp);
    if (!rowsFound.length) return [];

    const rows = [];

    rowsFound.forEach(function(rowNumber) {
      // A:AO = 41 kolum. Hanya row yang sudah dipastikan milik murid dibaca.
      const row = sheet.getRange(rowNumber, 1, 1, 41).getDisplayValues()[0];
      const grades = {};

      config.subjects.forEach(function(item) {
        const label = item[0];
        const colNumber = item[1]; // 1-based sheet column
        grades[label] = String(row[colNumber - 1] || '').trim();
      });

      rows.push({
        tahun: String(row[0] || '').trim(), // A
        exam: String(row[1] || '').trim(),  // B
        kelas: String(row[6] || '').trim(), // G = KELAS
        grades: grades,
        _row: rowNumber
      });
    });

    rows.sort(function(a, b) {
      const yaMatch = String(a.tahun).match(/\d{4}/);
      const ybMatch = String(b.tahun).match(/\d{4}/);
      const ya = yaMatch ? parseInt(yaMatch[0], 10) : 0;
      const yb = ybMatch ? parseInt(ybMatch[0], 10) : 0;
      if (yb !== ya) return yb - ya;
      return b._row - a._row;
    });

    rows.forEach(function(item) { delete item._row; });
    return rows;
  };

  result.menBawah = readExamSheet('Markah_MenBawah', CONFIG.Markah_MenBawah);
  result.menAtas = readExamSheet('Markah_MenAtas', CONFIG.Markah_MenAtas);

  return result;
}


/**
 * Ambil satu profil penuh apabila user memilih nama murid.
 * Cari nama pada kolum B, kemudian baca SATU row A:BN.
 * No KP dinormalisasi kepada 12 digit dan gambar dipadankan melalui RegGambar!C:C.
 */
function getStudentProfile(token, name) {
  requireModuleAccess_(token, 1);
  const target = String(name || '').trim();
  if (!target) return null;

  const sheet = getDatabaseSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const range = sheet.getRange(2, 2, lastRow - 1, 1);
  const finder = range.createTextFinder(target)
    .matchEntireCell(true)
    .matchCase(false);
  const cell = finder.findNext();
  if (!cell) return null;

  const rowNumber = cell.getRow();
  const values = sheet.getRange(rowNumber, 1, 1, DATABASE_COLUMNS).getDisplayValues()[0];
  const headers = sheet.getRange(1, 1, 1, DATABASE_COLUMNS).getDisplayValues()[0]
    .map(h => String(h).trim());

  const obj = {};
  headers.forEach((h, i) => {
    if (h) obj[h] = values[i] || '';
  });

  // Pastikan No KP sentiasa 12 digit untuk paparan dan padanan gambar.
  if (Object.prototype.hasOwnProperty.call(obj, 'NO KP')) {
    obj['NO KP'] = normalizeNoKp_(obj['NO KP']);
  }

  const image = getStudentImage_(obj['NO KP']);
  obj['GAMBAR_MURID'] = image.url || '';
  obj['GAMBAR_MURID_ID'] = image.fileId || '';
  obj['GAMBAR_MURID_NAMA_FAIL'] = image.fileName || '';

  // Maklumat waris dipaparkan dalam Profil Murid.
  // Carian dibuat berdasarkan NO KP murid pada tab Daftar.
  obj['WARIS'] = getStudentWaris_(obj['NO KP']);

  // Alamat rumah diambil terus daripada tab Daftar, kolum N.
  obj['ALAMAT_RUMAH'] = getStudentAddress_(obj['NO KP']);

  return obj;
}

/**
 * Lazy-load Prestasi Akademik P1 selepas profil asas sudah dipaparkan.
 * Ini mengelakkan getStudentProfile membaca Markah_MenBawah + Markah_MenAtas
 * pada request yang sama.
 */
function getStudentAcademicP1(token, noKp) {
  requireModuleAccess_(token, 1);
  const kp = normalizeNoKp_(noKp);
  if (!kp) return { menBawah: [], menAtas: [] };

  const cache = CacheService.getScriptCache();
  const key = 'UBK_ACAD_P1_' + kp;
  const cached = cache.get(key);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const result = getAcademicP1_(kp);
  try {
    const json = JSON.stringify(result);
    if (json.length <= 95000) cache.put(key, json, 600);
  } catch (e) {}
  return result;
}

/** Ambil hanya data PHQ/GAD + filter asas selepas kata kunci disahkan. */
function getMindaSihatData(token) {
  requireModuleAccess_(token, 8);
  const raw = readDatabase_();
  return projectRows_(raw, MINDA_HEADERS);
}

/** Password Minda Sihat disimpan di Script Properties, bukan Index.html. */
function setupMindaSihatPassword() {
  PropertiesService.getScriptProperties().setProperty('MINDA_SIHAT_PASSWORD', 'KAUNSELING');
}

function checkMindaAccess(token, password) {
  requireModuleAccess_(token, 8);
  const saved = PropertiesService.getScriptProperties().getProperty('MINDA_SIHAT_PASSWORD');
  if (!saved) {
    throw new Error('MINDA_SIHAT_PASSWORD belum ditetapkan. Jalankan setupMindaSihatPassword() dahulu.');
  }
  return String(password || '').trim().toUpperCase() === String(saved).trim().toUpperCase();
}

function ping() {
  const sheet = getDatabaseSheet_();
  return {
    ok: true,
    spreadsheetId: SPREADSHEET_ID,
    sheetName: sheet.getName(),
    lastRow: sheet.getLastRow(),
    lastColumn: sheet.getLastColumn(),
    snapshot: getSnapshotStatus()
  };
}

function testV8Snapshot() {
  const status = getSnapshotStatus();
  Logger.log(JSON.stringify(status, null, 2));
  return status;
}

function testGetDashboardData() {
  const result = getDashboardData();
  Logger.log('Sheet: ' + DATABASE_SHEET);
  Logger.log('Murid: ' + result.rows.length);
  Logger.log('Kolum dashboard: ' + result.headers.length);
  return { murid: result.rows.length, kolum: result.headers.length };
}

function testGetStudentProfile() {
  return { ok: false, message: 'Gunakan Web App selepas log masuk untuk menguji getStudentProfile.' };
}


/**
 * Data khusus modul DORM.
 * Membaca terus daripada PROFIL MURID supaya kolum Dorm yang terkini
 * tidak bergantung pada cache dashboard lama.
 * PROFIL MURID: B=NAMA, D=JANTINA, E=TINGKATAN, F=KELAS, I=DORM.
 */
function getDormAnalysisData(token, moduleNo) {
  validateSession_(token);
  const sheet = getDatabaseSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { headers: ['NAMA MURID','JANTINA','TINGKATAN','KELAS','DORM'], rows: [] };

  // B:I supaya data diambil terus daripada kolum sumber sebenar.
  const values = sheet.getRange(1, 2, lastRow, 8).getDisplayValues();
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const name = String(row[0] || '').trim();      // B
    if (!name) continue;
    rows.push([
      name,                                        // B: NAMA MURID
      String(row[2] || '').trim(),                 // D: JANTINA
      String(row[3] || '').trim(),                 // E: TINGKATAN
      String(row[4] || '').trim(),                 // F: KELAS
      String(row[7] || '').trim()                  // I: DORM
    ]);
  }
  return { headers: ['NAMA MURID','JANTINA','TINGKATAN','KELAS','DORM'], rows: rows };
}

