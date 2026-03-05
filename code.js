// ============================================================
// CONFIGURATION
// ============================================================
const MAIN_FOLDER_ID    = '1yJM6SmvbDddE1iYkvow-C5-FBZJL79Pc';
const GOOGLE_CLOUD_API_KEY = 'AIzaSyBWAYkfgkgOBH6EZkIzQurymTr7u7vBD64';


// ============================================================
// ENTRY POINT
// ============================================================
function processAllYears() {
  clearAllSheetDataExceptHeaders();
  const clearedSheets = new Set();
  const mainFolder    = DriveApp.getFolderById(MAIN_FOLDER_ID);
  const yearFolders   = mainFolder.getFolders();

  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    let parsedSheetFile = null;
    let masterListFile  = null;

    const files = yearFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) {
      const file     = files.next();
      const fileName = file.getName().toLowerCase();
      if (fileName.includes('master_list')) {
        masterListFile = file;
      } else {
        parsedSheetFile = file;
      }
    }

    if (!parsedSheetFile || !masterListFile) continue;

    const parsedSpreadsheet = SpreadsheetApp.open(parsedSheetFile);
    const branchFolders     = yearFolder.getFolders();

    while (branchFolders.hasNext()) {
      const branchFolder = branchFolders.next();
      const branchName   = branchFolder.getName();
      const sheetTab     = getOrCreateSheetTab(parsedSpreadsheet, branchName);

      if (!clearedSheets.has(branchName)) {
        ensureHeaders(sheetTab);
        clearedSheets.add(branchName);
      }
      copyRollNumbersFromMasterListToParsedSheet(parsedSheetFile, masterListFile);
      processBranch(branchFolder, sheetTab);
    }
  }
}


// ============================================================
// SHEET HELPERS
// ============================================================
function getOrCreateSheetTab(spreadsheet, branchName) {
  const tabs     = spreadsheet.getSheets();
  let   sheetTab = tabs.find(s => s.getName() === branchName);
  if (!sheetTab) sheetTab = spreadsheet.insertSheet(branchName);
  return sheetTab;
}

function ensureHeaders(sheetTab) {
  const headers = [
    'College Roll No.',
    'Name',

    'Name as per 10th marksheet',
    '10th Board Name',
    '10th Board Roll No.',
    "10th Mother's Name",
    "10th Father's Name",
    "10th Academic Year",

    '10th Subject 1',
    '10th Subject 1 Marks',
    '10th Subject 2',
    '10th Subject 2 Marks',
    '10th Subject 3',
    '10th Subject 3 Marks',
    '10th Subject 4',
    '10th Subject 4 Marks',
    '10th Subject 5',
    '10th Subject 5 Marks',
    '10th Subject 6',
    '10th Subject 6 Marks',

    '10th Percentage',
    '10th Result',

    'Name as per 12th marksheet',
    '12th Board Name',
    '12th Board Roll No.',
    "12th Mother's Name",
    "12th Father's Name",
    "12th Academic Year",

    '12th Subject 1',
    '12th Subject 1 Marks',
    '12th Subject 2',
    '12th Subject 2 Marks',
    '12th Subject 3',
    '12th Subject 3 Marks',
    '12th Subject 4',
    '12th Subject 4 Marks',
    '12th Subject 5',
    '12th Subject 5 Marks',
    '12th Subject 6',
    '12th Subject 6 Marks',

    '12th Percentage',
    '12th Result'
  ];

  const range = sheetTab.getRange(1, 1, 1, headers.length);
  range.clearContent();
  range.setValues([headers]);
  range.setFontWeight('bold');
  range.setBackground('#4285F4');
  range.setFontColor('#FFFFFF');
  range.setWrap(true);
}

function copyRollNumbersFromMasterListToParsedSheet(parsedSheetFile, masterListFile) {
  const parsedSpreadsheet = SpreadsheetApp.open(parsedSheetFile);
  const masterSpreadsheet = SpreadsheetApp.open(masterListFile);

  masterSpreadsheet.getSheets().forEach(masterBranchSheet => {
    const branchName      = masterBranchSheet.getName();
    const parsedBranchSheet = parsedSpreadsheet.getSheetByName(branchName);
    if (!parsedBranchSheet) return;

    const masterData    = masterBranchSheet.getDataRange().getValues();
    const masterHeaders = masterData[0];
    const rollColIndex  = masterHeaders.indexOf('ROLL NO.');
    const nameIndex     = masterHeaders.indexOf('Name');

    if (rollColIndex === -1 || nameIndex === -1) return;

    const rollNumbers = masterData.slice(1).map(r => r[rollColIndex]).filter(r => r);
    const nameList    = masterData.slice(1).map(r => r[nameIndex]).filter(r => r);

    if (parsedBranchSheet.getLastColumn() === 0) ensureHeaders(parsedBranchSheet);

    const parsedHeaders = parsedBranchSheet
      .getRange(1, 1, 1, parsedBranchSheet.getLastColumn())
      .getValues()[0];

    let colIndex     = parsedHeaders.indexOf('College Roll No.');
    let namecolIndex = parsedHeaders.indexOf('Name');

    if (colIndex === -1)     { colIndex = parsedHeaders.length;     parsedBranchSheet.getRange(1, colIndex + 1).setValue('College Roll No.'); }
    if (namecolIndex === -1) { namecolIndex = parsedHeaders.length; parsedBranchSheet.getRange(1, namecolIndex + 1).setValue('Name'); }

    rollNumbers.forEach((rn, i) => parsedBranchSheet.getRange(i + 2, colIndex + 1).setValue(rn));
    nameList.forEach((nm, i)    => parsedBranchSheet.getRange(i + 2, namecolIndex + 1).setValue(nm));
  });
}


// ============================================================
// GOOGLE VISION OCR
// ============================================================
function parseMarksheetFromFile(file) {
  const bytes      = file.getBlob().getBytes();
  const base64Image = Utilities.base64Encode(bytes);

  const payload = {
    requests: [{
      image:    { content: base64Image },
      features: [{ type: 'TEXT_DETECTION' }]
    }]
  };

  const response = UrlFetchApp.fetch(
    'https://vision.googleapis.com/v1/images:annotate?key=' + GOOGLE_CLOUD_API_KEY,
    {
      method:          'post',
      contentType:     'application/json',
      payload:         JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  const json = JSON.parse(response.getContentText());

  if (!json.responses ||
      !json.responses[0] ||
      !json.responses[0].fullTextAnnotation) {
    Logger.log('❌ No text found in: ' + file.getName());
    return null;
  }

  return json.responses[0].fullTextAnnotation.text;
}


// ============================================================
// BOARD DETECTION
// ============================================================
function detectBoard(text) {
  const t = text.toLowerCase();
  if (t.includes('central board of secondary education') || t.includes('cbse')) return 'CBSE';
  if (t.includes('council for the indian school certificate') ||
      t.includes('icse') ||
      t.includes('indian school certificate')) return 'ICSE';
  if (t.includes('maharashtra state board') ||
      t.includes('state board') ||
      t.includes('divisional board') ||
      t.includes('konkan divisional')) return 'STATE';
  return 'Unknown';
}


// ============================================================
// DISPATCHER
// ============================================================
function parseMarksheet(text, examType) {
  if (!text) {
    Logger.log('❌ parseMarksheet received null/undefined text');
    return null;
  }

  const board = detectBoard(text);
  Logger.log('✓ Board detected: ' + board + ' | Exam: ' + examType);

  if (board === 'CBSE')  return examType === '10th' ? parseCBSE10th(text)  : parseCBSE12th(text);
  if (board === 'STATE') return examType === '10th' ? parseState10th(text) : parseState12th(text);
  if (board === 'ICSE')  return examType === '10th' ? parseICSE10th(text)  : parseISC12th(text);

  Logger.log('⚠️ Unknown board — skipping');
  return null;
}


// ============================================================
// HELPER: safe regex extract (returns '' on no match)
// ============================================================
function reExtract(text, regex) {
  if (!text) return 'Not Found';
  const m = text.match(regex);
  return m ? m[1].trim() : 'Not Found';
}


// ============================================================
// PARSER 1 — CBSE 10th
// Table columns: SUB.CODE | SUBJECT | THEORY | IA/PR | TOTAL | IN WORDS | GRADE
// Co-scholastic codes 500-504 have no numeric marks — excluded by code filter.
// CBSE 10th percentage = best 5 of scored subjects.
// ============================================================
function parseCBSE10th(text) {
  Logger.log('🔍 Parsing CBSE 10th');

  // ── Student info ──────────────────────────────────────────
  const name       = reExtract(text, /This is to certify that\s+([A-Z][A-Z\s]+?)(?:\n|Roll)/i);
  const rollNo     = reExtract(text, /Roll\s*No[.\s]+(\d+)/i);
  const motherName = reExtract(text, /Mother'?s?\s*Name\s+([A-Z][A-Z\s]+?)(?:\n|Father)/i);
  const fatherName = reExtract(text, /Father'?s?\s*(?:\/\s*Guardian'?s?\s*)?Name\s+([A-Z][A-Z\s]+?)(?:\n|Date|School|विद्यालय)/i);
  const year       = reExtract(text, /(?:SECONDARY SCHOOL EXAMINATION|EXAMINATION)[,\s]+(\d{4})/i);

  // ── Subjects: line-by-line (Vision API outputs each cell on its own line) ─
  // Layout per subject row:
  //   184               ← standalone 3-digit code
  //   ENGLISH LNG & LIT.← subject name (all-caps, no digits)
  //   077               ← theory
  //   020               ← IA/PR
  //   097               ← total
  //   NINETY SEVEN      ← words (skip)
  //   A1                ← grade (skip)
  //
  // Guards: skip co-scholastic codes 500-504, skip number-word subject lines,
  // use usedLines set so total-value line is never re-read as a new code.
  const NUMBER_WORD_RE = /^(NINETY|EIGHTY|SEVENTY|SIXTY|FIFTY|FORTY|THIRTY|TWENTY|TEN|ELEVEN|TWELVE|THIRTEEN|FOURTEEN|FIFTEEN|SIXTEEN|SEVENTEEN|EIGHTEEN|NINETEEN|ONE HUNDRED)/;
  const skipCodes  = new Set(['500', '501', '502', '503', '504']);
  const lines      = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  const subjects   = [];
  const seenCodes  = new Set();
  const usedLines  = new Set();

  for (let i = 0; i < lines.length; i++) {
    if (usedLines.has(i)) continue;
    if (!/^\d{3}$/.test(lines[i])) continue;
    const code = lines[i];
    if (skipCodes.has(code) || seenCodes.has(code)) continue;

    // Next unused line = subject name
    let si = i + 1;
    while (si < lines.length && usedLines.has(si)) si++;
    if (si >= lines.length) continue;
    const subjectLine = lines[si].trim();
    if (!/^[A-Z][A-Z\s&.()\-\/]+$/.test(subjectLine)) continue;
    if (NUMBER_WORD_RE.test(subjectLine)) continue;

    // Collect next 3 standalone 2-3 digit numbers: theory, pr, total
    const nums = [];
    for (let j = si + 1; j <= si + 8 && j < lines.length && nums.length < 3; j++) {
      if (usedLines.has(j)) continue;
      if (/^\d{2,3}$/.test(lines[j])) {
        nums.push({ val: parseInt(lines[j]), idx: j });
      } else if (nums.length > 0) {
        break;
      }
    }
    if (nums.length < 3) continue;

    const theory = nums[0].val, pr = nums[1].val, total = nums[2].val;
    if (total < 0 || total > 100) continue;
    if (Math.abs(total - (theory + pr)) > 2) continue;

    usedLines.add(i); usedLines.add(si);
    nums.forEach(n => usedLines.add(n.idx));
    seenCodes.add(code);

    subjects.push({ name: subjectLine, marks: total });
    Logger.log('  ✅ [' + code + '] ' + subjectLine + ': ' + total + ' (T:' + theory + ' PR:' + pr + ')');
  }

  const top5 = [...subjects].sort((a, b) => b.marks - a.marks).slice(0, 5);
  const percentage = top5.length > 0
    ? ((top5.reduce((s, x) => s + x.marks, 0) / (top5.length * 100)) * 100).toFixed(2)
    : '0.00';

  const result = /Result\s+PASS/i.test(text) ? 'PASS' : 'FAIL';
  Logger.log('  📊 ' + subjects.length + ' subjects | %: ' + percentage);
  return { name, rollNo, motherName, fatherName, year, subjects, percentage, result };
}


// ============================================================
// PARSER 2 — CBSE 12th
// Same table structure as 10th.
// Subjects: ENGLISH CORE, MATHEMATICS, PHYSICS, CHEMISTRY, COMPUTER SCIENCE etc.
// WORK EXPERIENCE (500), HEALTH & PHYSICAL EDUCATION (502), GENERAL STUDIES (503) — grade only, skip.
// ============================================================
function parseCBSE12th(text) {
  Logger.log('🔍 Parsing CBSE 12th');

  // ── Student info ──────────────────────────────────────────
  const name       = reExtract(text, /This is to certify that\s+([A-Z][A-Z\s]+?)(?:\n|Roll)/i);
  const rollNo     = reExtract(text, /Roll\s*No[.\s]+(\d+)/i);
  const motherName = reExtract(text, /Mother'?s?\s*Name\s+([A-Z][A-Z\s]+?)(?:\n|Father)/i);
  const fatherName = reExtract(text, /Father'?s?\s*(?:\/\s*Guardian'?s?\s*)?Name\s+([A-Z][A-Z\s]+?)(?:\n|School|Date|विद्यालय)/i);
  const year       = reExtract(text, /(?:SENIOR SCHOOL CERTIFICATE EXAMINATION|EXAMINATION)[,\s]+(\d{4})/i);

  // ── Subjects: same line-by-line approach as CBSE 10th ────
  const NUMBER_WORD_RE = /^(NINETY|EIGHTY|SEVENTY|SIXTY|FIFTY|FORTY|THIRTY|TWENTY|TEN|ELEVEN|TWELVE|THIRTEEN|FOURTEEN|FIFTEEN|SIXTEEN|SEVENTEEN|EIGHTEEN|NINETEEN|ONE HUNDRED)/;
  const skipCodes  = new Set(['500', '501', '502', '503', '504']);
  const lines      = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
  const subjects   = [];
  const seenCodes  = new Set();
  const usedLines  = new Set();

  for (let i = 0; i < lines.length; i++) {
    if (usedLines.has(i)) continue;
    if (!/^\d{3}$/.test(lines[i])) continue;
    const code = lines[i];
    if (skipCodes.has(code) || seenCodes.has(code)) continue;

    let si = i + 1;
    while (si < lines.length && usedLines.has(si)) si++;
    if (si >= lines.length) continue;
    const subjectLine = lines[si].trim();
    if (!/^[A-Z][A-Z\s&.()\-\/]+$/.test(subjectLine)) continue;
    if (NUMBER_WORD_RE.test(subjectLine)) continue;

    const nums = [];
    for (let j = si + 1; j <= si + 8 && j < lines.length && nums.length < 3; j++) {
      if (usedLines.has(j)) continue;
      if (/^\d{2,3}$/.test(lines[j])) {
        nums.push({ val: parseInt(lines[j]), idx: j });
      } else if (nums.length > 0) {
        break;
      }
    }
    if (nums.length < 3) continue;

    const theory = nums[0].val, pr = nums[1].val, total = nums[2].val;
    if (total < 0 || total > 100) continue;
    if (Math.abs(total - (theory + pr)) > 2) continue;

    usedLines.add(i); usedLines.add(si);
    nums.forEach(n => usedLines.add(n.idx));
    seenCodes.add(code);

    subjects.push({ name: subjectLine, marks: total });
    Logger.log('  ✅ [' + code + '] ' + subjectLine + ': ' + total + ' (T:' + theory + ' PR:' + pr + ')');
  }

  const top5 = [...subjects].sort((a, b) => b.marks - a.marks).slice(0, 5);
  const percentage = top5.length > 0
    ? ((top5.reduce((s, x) => s + x.marks, 0) / (top5.length * 100)) * 100).toFixed(2)
    : '0.00';

  const result = /Result\s+PASS/i.test(text) ? 'PASS' : 'FAIL';
  Logger.log('  📊 ' + subjects.length + ' subjects | %: ' + percentage);
  return { name, rollNo, motherName, fatherName, year, subjects, percentage, result };
}


// ============================================================
// PARSER 3 — Maharashtra State Board 10th (SSC)
// Table: SubCode  SubjectName  MaxMarks(100)  ObtainedFigure  ObtainedWords
// Example row: "03 ENGLISH (1ST LANG)   100   078   SEVENTYEIGHT"
// Percentage is printed on sheet as "Percentage/टक्केवारी  91.40"
// ============================================================
function parseState10th(text) {
  Logger.log('🔍 Parsing State Board 10th');

  // Candidate name appears after "CANDIDATE'S FULL NAME (SURNAME FIRST)" heading
  const nameMatch = text.match(/CANDIDATE'?S?\s+FULL\s+NAME[^\n]*\n([^\n]+)/i);
  const name = nameMatch
    ? nameMatch[1].replace(/\(SURNAME FIRST\)/i, '').trim()
    : 'Not Found';

  const rollNo     = reExtract(text, /SEAT\s*NO\.?\s*[:\s]+([A-Z0-9]+)/i);
  const motherName = reExtract(text, /MOTHER'?S?\s*NAME\s+([A-Z][A-Za-z\s]+?)(?:\n|$)/im);
  const fatherName = ''; // SSC sheets do not include father's name
  const year       = reExtract(text,
    /((?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)[-–]\d{4})/i);

  // ── Subject rows ──
  // Pattern: 2-digit-code  SUBJECT NAME(caps,spaces,brackets)  100  obtained  WORDWORD
  // We require "100" as max-marks anchor to avoid false matches.
  const subjects = [];
  const rowRe    = /\b(\d{2})\s+([A-Z][A-Z\s()\/&]+?)\s{2,}100\s{1,6}(\d{2,3})\s+[A-Z]{3,}/gm;
  let m;
  while ((m = rowRe.exec(text)) !== null) {
    const subject = m[2].trim().replace(/\s{2,}/g, ' ');
    const marks   = parseInt(m[3]);
    if (marks >= 0 && marks <= 100) {
      subjects.push({ name: subject, marks });
      Logger.log('  ✅ ' + subject + ': ' + marks);
    }
  }

  // Fallback line-by-line if regex caught nothing (heavily OCR-warped text)
  if (subjects.length === 0) {
    Logger.log('  ⚠️ Regex found 0 subjects — running line-by-line fallback');
    const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l);

    const subjectList = [];
    let inSubjects = false;
    for (let i = 0; i < lines.length; i++) {
      if (/Subject Code No\.|विषयाचा सांकेतिक/i.test(lines[i])) { inSubjects = true; continue; }
      if (inSubjects && /कमाल|Max\./i.test(lines[i])) { inSubjects = false; break; }
      if (inSubjects) {
        const sub = lines[i].match(/^(\d{2})\s+(.+)$/);
        if (sub) subjectList.push(sub[2].trim());
      }
    }

    const marksList = [];
    for (let i = 0; i < lines.length - 1; i++) {
      if (lines[i] === '100') {
        const next = parseInt(lines[i + 1]);
        if (!isNaN(next) && next >= 0 && next <= 100) marksList.push(next);
      }
    }

    const n = Math.min(subjectList.length, marksList.length);
    for (let i = 0; i < n; i++) {
      subjects.push({ name: subjectList[i], marks: marksList[i] });
      Logger.log('  ✅ [fb] ' + subjectList[i] + ': ' + marksList[i]);
    }
  }

  // Percentage from sheet; compute if not found
  let percentage = '0.00';
  const pctM = text.match(/(?:Percentage|टक्केवारी)\s*[\/|]?\s*(?:टक्केवारी\s*)?(\d{1,3}\.\d{2})/i)
            || text.match(/टक्केवारी\s+(\d{1,3}\.\d{2})/);
  if (pctM) {
    percentage = pctM[1];
  } else if (subjects.length > 0) {
    const grand = subjects.reduce((s, x) => s + x.marks, 0);
    percentage  = ((grand / (subjects.length * 100)) * 100).toFixed(2);
  }

  const result = /PASS/i.test(text) ? 'PASS' : 'FAIL';
  Logger.log('  📊 ' + subjects.length + ' subjects | %: ' + percentage);
  return { name, rollNo, motherName, fatherName, year, subjects, percentage, result };
}


// ============================================================
// PARSER 4 — Maharashtra State Board 12th (HSC)
//
// Actual OCR layout from marksheet image:
//   STREAM   SEAT NO.   CENTRE NO.   DIST...   MONTH & YEAR   SR.NO.
//   ARTS     M174465    4237         ...        FEBRUARY-24    341845
//
//   CANDIDATE'S FULL NAME (SURNAME FIRST)
//   Chorge Arnav Prashant                    ← mixed-case, own line
//
//   CANDIDATE'S MOTHER'S NAME   Pranali      ← inline after label
//
// Subject table has 5 columns — OCR outputs ALL on ONE line per subject:
//   "01 ENGLISH ENG 100 075 SEVENTYFIVE"
//   "33 SANSKRIT ENG 100 082 EIGHTYTWO"
//   "31 ENV. EDU. & WATER SECURITY ENG - A"  ← dash = grade only, skip
//   "30 HEALTH & PHYSICAL EDUCATION ENG - A" ← skip
//
// Key difference from SSC: there is a MEDIUM column (ENG/MAR/URD/SEM)
// between subject name and max marks. We must skip over it.
//
// Percentage: "टक्केवारी/ Percentage   83.50" printed near bottom.
// ============================================================
function parseState12th(text) {
  Logger.log('🔍 Parsing State Board 12th');

  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);

  // ── Student name ──────────────────────────────────────────
  // The heading line contains "CANDIDATE'S FULL NAME (SURNAME FIRST)"
  // The actual name is on the NEXT non-empty line (mixed-case).
  let name = 'Not Found';
  for (let i = 0; i < lines.length; i++) {
    if (/CANDIDATE'?S?\s+FULL\s+NAME/i.test(lines[i])) {
      // Check if name is appended inline after the closing bracket
      const afterBracket = lines[i].replace(/^.*\)\s*/i, '').trim();
      if (afterBracket.length > 2 && !/CANDIDATE|FULL|NAME|SURNAME/i.test(afterBracket)) {
        name = afterBracket;
      } else {
        // Name is on the next line
        for (let j = i + 1; j < lines.length; j++) {
          const candidate = lines[j].trim();
          // Name is mixed-case (not all caps, not a label line)
          if (candidate.length > 2 && !/^[A-Z\s\/&.]+$/.test(candidate) && !/MOTHER|FATHER|SEAT|CENTRE/i.test(candidate)) {
            name = candidate;
            break;
          }
          // Also accept all-caps names
          if (/^[A-Z][A-Za-z\s]+$/.test(candidate) && candidate.length > 4 && !/MOTHER|FATHER|SEAT|CENTRE|CANDIDATE/i.test(candidate)) {
            name = candidate;
            break;
          }
        }
      }
      break;
    }
  }

  // ── Seat / Roll number ────────────────────────────────────
  // OCR outputs the header row and data row column by column:
  //   STREAM / SEAT NO. / CENTRE NO. / ... (headers, each on own line)
  //   ARTS   / M174465  / 4237       / ... (values, each on own line)
  //
  // Three strategies, in priority order:
  //   1. Inline: "SEAT NO. M174465" on same line
  //   2. Post-stream scan: find ARTS/SCIENCE/COMMERCE value line, grab next alphanumeric
  //   3. Header-offset: find SEAT NO. position in header block, apply same offset to data block
  let rollNo = 'Not Found';

  // Strategy 1: inline
  const inlineSeat = text.match(/SEAT\s*NO\.?\s+([A-Z]{0,2}\d{4,})/i);
  if (inlineSeat) {
    rollNo = inlineSeat[1];
  }

  // Strategy 2: post-stream-value scan
  if (rollNo === 'Not Found') {
    for (let i = 0; i < lines.length; i++) {
      if (/^(?:ARTS|SCIENCE|COMMERCE|VOCATIONAL)$/i.test(lines[i])) {
        for (let j = i + 1; j <= i + 4 && j < lines.length; j++) {
          if (/^[A-Z]{0,2}\d{5,7}$/.test(lines[j])) { rollNo = lines[j]; break; }
        }
        break;
      }
      // Also catch single-line data row: "ARTS M174465 4237 ..."
      if (/^(?:ARTS|SCIENCE|COMMERCE|VOCATIONAL)\s+([A-Z]{0,2}\d{5,})/i.test(lines[i])) {
        const dm = lines[i].match(/^(?:ARTS|SCIENCE|COMMERCE|VOCATIONAL)\s+([A-Z]{0,2}\d{5,})/i);
        if (dm) { rollNo = dm[1]; break; }
      }
    }
  }

  // Strategy 3: header-offset (SEAT NO. is Nth header → Nth value after stream value)
  if (rollNo === 'Not Found') {
    let streamHdrIdx = -1, seatHdrIdx = -1;
    for (let i = 0; i < lines.length; i++) {
      if (/^STREAM$/i.test(lines[i]) && streamHdrIdx === -1) streamHdrIdx = i;
      if (/^SEAT\s*NO\.?$/i.test(lines[i]) && seatHdrIdx === -1) seatHdrIdx = i;
    }
    if (streamHdrIdx !== -1 && seatHdrIdx !== -1) {
      const offset = seatHdrIdx - streamHdrIdx;
      for (let i = 0; i < lines.length; i++) {
        if (/^(?:ARTS|SCIENCE|COMMERCE|VOCATIONAL)$/i.test(lines[i])) {
          const candidate = lines[i + offset] || '';
          if (/^[A-Z]{0,2}\d{4,}$/.test(candidate)) rollNo = candidate;
          break;
        }
      }
    }
  }

  // ── Mother's name ─────────────────────────────────────────
  // "CANDIDATE'S MOTHER'S NAME   Pranali"  — always on the same line after label
  let motherName = 'Not Found';
  for (let i = 0; i < lines.length; i++) {
    if (/MOTHER'?S?\s*NAME/i.test(lines[i])) {
      const inlineM = lines[i].match(/MOTHER'?S?\s*NAME\s+(.+)/i);
      if (inlineM && inlineM[1].trim().length > 1) {
        motherName = inlineM[1].trim();
      } else if (i + 1 < lines.length) {
        motherName = lines[i + 1].trim();
      }
      break;
    }
  }

  const fatherName = ''; // HSC sheets do not include father's name

  // ── Year ──────────────────────────────────────────────────
  // "FEBRUARY-24" → store as-is, or full form "MARCH-2015"
  // Also handle 2-digit year: "FEBRUARY-24" means 2024
  let year = 'Not Found';
  const yrM = text.match(/((?:JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)[-–](\d{2,4}))/i);
  if (yrM) {
    let yr = yrM[2];
    if (yr.length === 2) yr = '20' + yr; // "24" → "2024"
    year = yrM[1].split(/[-–]/)[0].toUpperCase() + '-' + yr;
  }

  // ── Subjects ─────────────────────────────────────────────
  // Full row (single OCR line):
  //   "01 ENGLISH ENG 100 075 SEVENTYFIVE"
  //   "33 SANSKRIT ENG 100 082 EIGHTYTWO"
  //   "42 POLITICAL SCIENCE ENG 100 083 EIGHTYTHREE"
  //   "48 PSYCHOLOGY ENG 100 093 NINETYTHREE"
  //
  // Pattern: \d{2,3}  SUBJECT_NAME  MEDIUM(ENG|MAR|URD|SEM|HIN)  100  MARKS  WORDWORD
  // Grade-only rows have "-" instead of 100 — they will NOT match the 100 anchor.
  //
  // OCR may also split across lines. We handle both single-line and multi-line cases.

  const MEDIUM_RE = /^(?:ENG|MAR|URD|SEM|HIN|GUJ|SAN|TAM|TEL|KAN|BEN|PUN|ORI|NEP|SIN|KAS|KON|BOD|DOG|MAI|MAN|SAN)$/;

  const subjects = [];
  const seenCodes = new Set();

  // ── Pass 1: single-line rows ──────────────────────────────
  // Regex accounts for optional medium token between subject name and "100"
  // Medium is 2-3 uppercase letters. We make it optional with (?:MEDIUM\s+)?
  const singleRe = /^(\d{2,3})\s+([A-Z][A-Z0-9\s().\/&,'%-]+?)\s+(?:[A-Z]{2,3}\s+)?100\s+(\d{2,3})\s+[A-Z]{4,}/;

  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(singleRe);
    if (!m) continue;
    const code    = m[1];
    const subject = m[2].trim().replace(/\s{2,}/g, ' ');
    const marks   = parseInt(m[3]);
    if (!seenCodes.has(code) && marks >= 0 && marks <= 100) {
      subjects.push({ name: subject, marks });
      seenCodes.add(code);
      Logger.log('  ✅ [1-line] ' + subject + ': ' + marks);
    }
  }

  // ── Pass 2: multi-line fallback ───────────────────────────
  // OCR may output:
  //   line i:   "01 ENGLISH"          or  "01 ENGLISH ENG"
  //   line i+1: "ENG"                 (medium, if split off)
  //   line i+2: "100"                 (max marks)
  //   line i+3: "075"                 (obtained)
  //   line i+4: "SEVENTYFIVE"         (words — ignored)
  //
  // Strategy: find lines starting with numeric code + subject name,
  // then scan next 6 lines for the sequence: (optional medium) → "100" → obtained
  if (subjects.length === 0) {
    Logger.log('  ⚠️ Single-line regex found 0 — trying multi-line fallback');

    const subjectStartRe = /^(\d{2,3})\s+([A-Z][A-Z0-9\s().\/&,'%-]+)$/;

    for (let i = 0; i < lines.length; i++) {
      const sm = lines[i].match(subjectStartRe);
      if (!sm) continue;
      const code    = sm[1];
      const subject = sm[2].trim().replace(/\s{2,}/g, ' ');
      if (seenCodes.has(code)) continue;

      // Scan next 1-6 lines for "100" then obtained
      let found100 = -1;
      for (let j = i + 1; j <= i + 6 && j < lines.length; j++) {
        if (lines[j] === '100') { found100 = j; break; }
        // Also handle "100" appearing mid-line with the medium: "ENG 100"
        if (/^(?:[A-Z]{2,3}\s+)?100$/.test(lines[j])) { found100 = j; break; }
      }
      if (found100 === -1) continue;

      // Next non-empty line after "100" should be the obtained marks
      if (found100 + 1 < lines.length) {
        const obtainedLine = lines[found100 + 1].trim();
        const obtainedM    = obtainedLine.match(/^(\d{2,3})$/);
        if (obtainedM) {
          const marks = parseInt(obtainedM[1]);
          if (marks >= 0 && marks <= 100) {
            subjects.push({ name: subject, marks });
            seenCodes.add(code);
            Logger.log('  ✅ [multi-line] ' + subject + ': ' + marks);
            i = found100 + 1; // skip ahead
          }
        }
      }
    }
  }

  // ── Percentage ────────────────────────────────────────────
  // Printed as:  "टक्केवारी/ Percentage   83.50"
  // or:          "Percentage  83.50"
  let percentage = '0.00';
  const pctM = text.match(/(?:टक्केवारी|Percentage)\s*[\/|]?\s*(?:Percentage\s*)?(\d{1,3}\.\d{2})/i)
            || text.match(/(\d{2,3}\.\d{2})\s*(?:टक्केवारी|Percentage)/i);
  if (pctM) {
    percentage = pctM[1];
  } else if (subjects.length > 0) {
    const grand = subjects.reduce((s, x) => s + x.marks, 0);
    percentage  = ((grand / (subjects.length * 100)) * 100).toFixed(2);
  }

  const result = /\bPASS\b/i.test(text) ? 'PASS' : 'FAIL';
  Logger.log('  📊 ' + subjects.length + ' subjects | %: ' + percentage);
  return { name, rollNo, motherName, fatherName, year, subjects, percentage, result };
}


// ============================================================
// PARSER 5 — ICSE 10th
//
// OCR layout — subjects and marks appear on THE SAME LINE:
//   ENGLISH LANGUAGE        85  EIGHTY FIVE
//   LITERATURE IN ENGLISH   91  NINETY ONE
//   HINDI                   78  SEVENTY EIGHT
//
// The left column has the subject name (all caps).
// The right column has: mark(2 digits)  WORD WORD
// Both columns are on the same physical line in OCR output.
//
// Unique-ID is roll number. No printed percentage — compute average.
// ============================================================
function parseICSE10th(text) {
  Logger.log('🔍 Parsing ICSE 10th');

  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);

  // ── Student info ─────────────────────────────────────────
  // "Name  DEEPENDRA DEV SINGH"
  let name = 'Not Found';
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(/^Name\s+([A-Z][A-Z\s]+)$/i);
    if (m) { name = m[1].trim(); break; }
  }

  const rollNo     = reExtract(text, /UNIQUE\s*ID\s+(\d+)/i);
  // "Smt  VEENITA SINGH" — may have period or space after Smt
  const motherName = reExtract(text, /Smt\.?\s+([A-Z][A-Z\s]+?)(?:\n|Shri|$)/im);
  const fatherName = reExtract(text, /Shri\.?\s+([A-Z][A-Z\s]+?)(?:\n|$)/im);
  // "INDIAN SCHOOL CERTIFICATE EXAMINATION (CLASS - X) - YEAR 2022"
  const year       = reExtract(text, /[-–]\s*YEAR\s+(\d{4})/i);

  // ── Subjects ─────────────────────────────────────────────
  // Known ICSE 10th subjects — ordered longest-first so "ENGLISH LANGUAGE"
  // matches before bare "ENGLISH"
  const KNOWN = [
    'ENGLISH LANGUAGE',
    'LITERATURE IN ENGLISH',
    'HISTORY, CIVICS & GEOGRAPHY',
    'HISTORY & CIVICS',
    'COMPUTER APPLICATIONS',
    'COMPUTER SCIENCE',
    'ENVIRONMENTAL SCIENCE',
    'COMMERCIAL STUDIES',
    'PHYSICAL EDUCATION',
    'HOME SCIENCE',
    'MATHEMATICS',
    'GEOGRAPHY',
    'CHEMISTRY',
    'BIOLOGY',
    'SCIENCE',
    'PHYSICS',
    'ECONOMICS',
    'SANSKRIT',
    'MARATHI',
    'FRENCH',
    'HINDI',
    'ART',
  ];

  const subjects = [];
  const seen     = new Set();

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Each subject row looks like:
    //   "ENGLISH LANGUAGE 85 EIGHTY FIVE"    (subject  mark  WORDS)
    // or sometimes OCR puts mark on the next line:
    //   "ENGLISH LANGUAGE"
    //   "85 EIGHTY FIVE"
    //
    // Match strategy: find which known subject the line STARTS WITH,
    // then extract the 2-digit number that follows (on same line or next).

    const matched = KNOWN.find(s => line === s || line.startsWith(s + ' ') || line.startsWith(s + '\t'));
    if (!matched || seen.has(matched)) continue;

    // Try to get mark from same line (after subject name)
    const afterSubject = line.slice(matched.length).trim();
    const inlineMatch  = afterSubject.match(/^(\d{2})\s+[A-Z]+/);
    if (inlineMatch) {
      const mark = parseInt(inlineMatch[1]);
      if (mark >= 20 && mark <= 100) {
        subjects.push({ name: matched, marks: mark });
        seen.add(matched);
        Logger.log('  ✅ ' + matched + ': ' + mark);
        continue;
      }
    }

    // Fallback: mark on next 1-3 lines
    for (let j = i + 1; j <= i + 3 && j < lines.length; j++) {
      const nextMatch = lines[j].match(/^(\d{2})\s+[A-Z]+/);
      if (nextMatch) {
        const mark = parseInt(nextMatch[1]);
        if (mark >= 20 && mark <= 100) {
          subjects.push({ name: matched, marks: mark });
          seen.add(matched);
          Logger.log('  ✅ [next-line] ' + matched + ': ' + mark);
          break;
        }
      }
    }
  }

  const avg = subjects.length > 0
    ? (subjects.reduce((s, x) => s + x.marks, 0) / subjects.length).toFixed(2)
    : '0.00';

  const result = /RESULT\s*[-–]?\s*PASS/i.test(text) ? 'PASS' : 'FAIL';
  Logger.log('  📊 ' + subjects.length + ' subjects | Avg%: ' + avg);
  return { name, rollNo, motherName, fatherName, year, subjects, percentage: avg, result };
}


// ============================================================
// PARSER 6 — ISC 12th  (Indian School Certificate, Class XII)
//
// OCR layout — subject and mark are on THE SAME LINE:
//   ENGLISH    92  NINE  TWO
//   HINDI      97  NINE  SEVEN
//   PHYSICS    91  NINE  ONE
//   CHEMISTRY  93  NINE  THREE
//   BIOLOGY    97  NINE  SEVEN
//
// The layout section is:
//   SUBJECTS                    Percentage Marks
//   External Examination
//   ENGLISH   92 NINE TWO
//   ...
//   Internal Assessment         Grade
//   SUPW & COMMUNITY SERVICE    A
//
// No printed percentage — compute average of external subjects.
// ============================================================
function parseISC12th(text) {
  Logger.log('🔍 Parsing ISC 12th');

  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);

  // ── Student info ─────────────────────────────────────────
  let name = 'Not Found';
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(/^Name\s+([A-Z][A-Z\s]+)$/i);
    if (m) { name = m[1].trim(); break; }
  }

  const rollNo     = reExtract(text, /UNIQUE\s*ID\s+(\d+)/i);
  const motherName = reExtract(text, /Smt\.?\s+([A-Z][A-Z\s]+?)(?:\n|Shri|$)/im);
  const fatherName = reExtract(text, /Shri\.?\s+([A-Z][A-Z\s]+?)(?:\n|$)/im);
  // "INDIAN SCHOOL CERTIFICATE EXAMINATION (CLASS - XII) - YEAR 2020"
  const year       = reExtract(text, /[-–]\s*YEAR\s+(\d{4})/i);

  // ── Subjects ─────────────────────────────────────────────
  // Known ISC 12th subjects — longest first to avoid partial matches
  const KNOWN = [
    'COMPUTER SCIENCE',
    'POLITICAL SCIENCE',
    'ENVIRONMENTAL SCIENCE',
    'PHYSICAL EDUCATION',
    'HOME SCIENCE',
    'BIOTECHNOLOGY',
    'MATHEMATICS',
    'CHEMISTRY',
    'GEOGRAPHY',
    'SOCIOLOGY',
    'PSYCHOLOGY',
    'ECONOMICS',
    'COMMERCE',
    'ACCOUNTS',
    'HISTORY',
    'PHYSICS',
    'BIOLOGY',
    'ENGLISH',
    'SANSKRIT',
    'MARATHI',
    'FRENCH',
    'HINDI',
    'ART',
  ];

  const subjects  = [];
  const seen      = new Set();
  let   inExternal = false;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    if (/External\s*Examination/i.test(line)) { inExternal = true;  continue; }
    if (/Internal\s*Assessment/i.test(line))  { inExternal = false; break;   }
    if (!inExternal) continue;

    // Each subject row: "ENGLISH 92 NINE TWO"
    // Find which known subject this line starts with
    const matched = KNOWN.find(s => line === s || line.startsWith(s + ' ') || line.startsWith(s + '\t'));
    if (!matched || seen.has(matched)) continue;

    // Extract mark from same line (after subject name)
    const afterSubject = line.slice(matched.length).trim();
    const inlineMatch  = afterSubject.match(/^(\d{2})\s+[A-Z]+/);
    if (inlineMatch) {
      const mark = parseInt(inlineMatch[1]);
      if (mark >= 20 && mark <= 100) {
        subjects.push({ name: matched, marks: mark });
        seen.add(matched);
        Logger.log('  ✅ ' + matched + ': ' + mark);
        continue;
      }
    }

    // Fallback: mark on next 1-3 lines (OCR occasionally wraps)
    for (let j = i + 1; j <= i + 3 && j < lines.length; j++) {
      const nextMatch = lines[j].match(/^(\d{2})\s+[A-Z]+/);
      if (nextMatch) {
        const mark = parseInt(nextMatch[1]);
        if (mark >= 20 && mark <= 100) {
          subjects.push({ name: matched, marks: mark });
          seen.add(matched);
          Logger.log('  ✅ [next-line] ' + matched + ': ' + mark);
          break;
        }
      }
    }
  }

  const avg = subjects.length > 0
    ? (subjects.reduce((s, x) => s + x.marks, 0) / subjects.length).toFixed(2)
    : '0.00';

  const result = /RESULT\s*[-–]?\s*PASS/i.test(text) ? 'PASS' : 'FAIL';
  Logger.log('  📊 ' + subjects.length + ' subjects | Avg%: ' + avg);
  return { name, rollNo, motherName, fatherName, year, subjects, percentage: avg, result };
}


// ============================================================
// BRANCH / FOLDER PROCESSING
// ============================================================
function processBranch(branchFolder, sheetTab) {
  const subfolders = branchFolder.getFolders();
  while (subfolders.hasNext()) {
    const folder     = subfolders.next();
    const folderName = folder.getName().toLowerCase();
    if      (folderName.includes('10th') || folderName.includes('10')) processMarksheetSubfolder('10th', folder, sheetTab);
    else if (folderName.includes('12th') || folderName.includes('12')) processMarksheetSubfolder('12th', folder, sheetTab);
  }
}

function processMarksheetSubfolder(type, folder, sheet) {
  const files     = folder.getFiles();
  const dataRange = sheet.getDataRange().getValues();
  const headers   = dataRange[0];
  const rollIndex = headers.indexOf('College Roll No.');

  if (rollIndex === -1) {
    Logger.log("❌ 'College Roll No.' column not found in " + sheet.getName());
    return;
  }

  while (files.hasNext()) {
    const file       = files.next();
    const rollNumber = file.getName().split('.')[0].trim();

    Logger.log('\n' + '='.repeat(60));
    Logger.log('📋 Processing: ' + rollNumber + ' (' + type + ')');
    Logger.log('='.repeat(60));

    const text = parseMarksheetFromFile(file);
    if (!text) {
      Logger.log('❌ OCR returned no text — skipping');
      continue;
    }

    const parsedData = parseMarksheet(text, type);
    if (!parsedData) {
      Logger.log('❌ Parser returned null — skipping');
      continue;
    }

    // Find the matching row by roll number
    let matchedRow = -1;
    for (let i = 1; i < dataRange.length; i++) {
      if (String(dataRange[i][rollIndex]).trim() === rollNumber) {
        matchedRow = i + 1;
        break;
      }
    }

    if (matchedRow === -1) {
      Logger.log('❌ No row found for roll: ' + rollNumber);
      continue;
    }

    updateMarksheetData(sheet, headers, matchedRow, parsedData, detectBoard(text), type);
    Logger.log('✅ Data written to row ' + matchedRow);
  }
}


// ============================================================
// SHEET WRITE HELPERS
// ============================================================
function updateMarksheetData(sheet, headers, row, data, boardName, prefix) {
  const basicFields = [
    prefix + ' Board Name',
    prefix + ' Board Roll No.',
    prefix + " Mother's Name",
    prefix + " Father's Name",
    prefix + ' Academic Year',
    'Name as per ' + prefix + ' marksheet',
    prefix + ' Percentage',
    prefix + ' Result'
  ];
  const basicValues = [
    boardName,
    data.rollNo     || '',
    data.motherName || '',
    data.fatherName || '',
    data.year       || '',
    data.name       || '',
    data.percentage || '',
    data.result     || ''
  ];
  updateSheetRow(sheet, headers, basicFields, basicValues, row);

  const subjects   = data.subjects || [];
  const maxSubjects = Math.min(subjects.length, 6);
  for (let i = 0; i < maxSubjects; i++) {
    updateSheetRow(
      sheet, headers,
      [prefix + ' Subject ' + (i + 1), prefix + ' Subject ' + (i + 1) + ' Marks'],
      [subjects[i].name || '', subjects[i].marks || 0],
      row
    );
  }
}

function updateSheetRow(sheet, headers, fieldNames, values, row) {
  fieldNames.forEach((field, i) => {
    const col = headers.indexOf(field);
    if (col !== -1) sheet.getRange(row, col + 1).setValue(values[i]);
  });
}


// ============================================================
// FORMAT & CLEAR UTILITIES
// ============================================================
function smartFormatSheet() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(sheet => {
    const range = sheet.getDataRange();
    range.setWrap(true);
    range.setVerticalAlignment('middle');
    const lastCol = sheet.getLastColumn();
    for (let i = 1; i <= lastCol; i++) {
      sheet.autoResizeColumn(i);
      if (sheet.getColumnWidth(i) > 300) sheet.setColumnWidth(i, 300);
    }
    sheet.setFrozenRows(1);
  });
  SpreadsheetApp.flush();
}

function clearAllSheetDataExceptHeaders() {
  const mainFolder  = DriveApp.getFolderById(MAIN_FOLDER_ID);
  const yearFolders = mainFolder.getFolders();
  while (yearFolders.hasNext()) {
    const yearFolder = yearFolders.next();
    const files = yearFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
    let parsedSheetFile = null;
    while (files.hasNext()) {
      const file = files.next();
      if (!file.getName().toLowerCase().includes('master_list')) {
        parsedSheetFile = file; break;
      }
    }
    if (!parsedSheetFile) continue;
    const spreadsheet = SpreadsheetApp.open(parsedSheetFile);
    spreadsheet.getSheets().forEach(sheet => {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    });
  }
}
