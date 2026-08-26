/**
 * IT WIZARDS � Intramurals Registration System
 * Dynamic Multi-Tab Router & Live Registration Counter API
 *
 * ---------------------------------------------------------------
 *  ROUTING & COUNTING FEATURES:
 *  1. Automatic Gender Split for:
 *     - "Mr. & Ms Intramurals" -> "Mr. Intramurals" / "Ms. Intramurals"
 *     - "Volleyball"           -> "Volleyball - Male" / "Volleyball - Female"
 *     - "Chess"                -> "Chess - Male" / "Chess - Female"
 *  2. All other events -> Single dedicated tab per event.
 *  3. doGet API returns real-time registration counts for all events.
 * ---------------------------------------------------------------
 */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const eventName = (data.event || 'General').trim();
    const gender = (data.gender || '').trim();
    const timestamp = new Date().toLocaleString('en-PH', { timeZone: 'Asia/Manila' });

    // Determine target tab name
    const tabName = getTargetTabName(eventName, gender);

    // 1. Log to Specific Event/Gender Tab
    logToSheet(ss, tabName, [
      timestamp,
      data.name || '',
      gender || '',
      data.yearCourse || '',
      data.section || '',
      data.contact || '',
      data.email || ''
    ], ['Timestamp', 'Full Name', 'Gender', 'Year & Course', 'Section', 'Contact Number', 'Email Address']);

    // 2. Also log to Master Log tab for overall oversight
    logToSheet(ss, 'Master Log', [
      timestamp,
      eventName,
      data.name || '',
      gender || '',
      data.yearCourse || '',
      data.section || '',
      data.contact || '',
      data.email || ''
    ], ['Timestamp', 'Event', 'Full Name', 'Gender', 'Year & Course', 'Section', 'Contact Number', 'Email Address']);

    return ContentService.createTextOutput(JSON.stringify({ result: 'success', tab: tabName }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --- GET API: Returns Live Counts for Every Event -----------------------------
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const counts = {};

    sheets.forEach(sheet => {
      const name = sheet.getName();
      if (name !== 'Master Log') {
        const lastRow = sheet.getLastRow();
        counts[name] = Math.max(0, lastRow - 1);
      }
    });

    return ContentService.createTextOutput(JSON.stringify({ result: 'success', counts: counts }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --- Helper: Determine Tab Name ----------------------------------------------
function getTargetTabName(event, gender) {
  const normEvent = (event || 'General').trim();
  const normGender = (gender || '').trim();
  const lowerEvent = normEvent.toLowerCase();
  const lowerGender = normGender.toLowerCase();

  // 1. Mr. & Ms Intramurals (Split)
  if (lowerEvent.indexOf('mr.') !== -1 && lowerEvent.indexOf('ms.') !== -1) {
    if (lowerGender === 'male') return 'Mr. Intramurals';
    if (lowerGender === 'female') return 'Ms. Intramurals';
    return 'Mr. & Ms Intramurals';
  }

  // 2. Volleyball (Split)
  if (lowerEvent === 'volleyball') {
    if (lowerGender === 'male') return 'Volleyball - Male';
    if (lowerGender === 'female') return 'Volleyball - Female';
    return 'Volleyball';
  }

  // 3. Chess (Split)
  if (lowerEvent === 'chess') {
    if (lowerGender === 'male') return 'Chess - Male';
    if (lowerGender === 'female') return 'Chess - Female';
    return 'Chess';
  }

  // 4. Table Tennis (Split)
  if (lowerEvent === 'table tennis') {
    if (lowerGender === 'male') return 'Table Tennis - Male';
    if (lowerGender === 'female') return 'Table Tennis - Female';
    return 'Table Tennis';
  }

  // 5. Badminton (Split)
  if (lowerEvent === 'badminton') {
    if (lowerGender === 'male') return 'Badminton - Male';
    if (lowerGender === 'female') return 'Badminton - Female';
    return 'Badminton';
  }

  // 6. Pickleball (Split)
  if (lowerEvent === 'pickleball') {
    if (lowerGender === 'male') return 'Pickleball - Male';
    if (lowerGender === 'female') return 'Pickleball - Female';
    return 'Pickleball';
  }

  // 7. ALL OTHER EVENTS -> Single Tab per Event
  let cleanName = normEvent.replace(/[:\\/?*\[\]]/g, '').trim();
  if (cleanName.length > 30) {
    cleanName = cleanName.substring(0, 30).trim();
  }

  return cleanName || 'Registrations';
}

// --- Helper: Append row & auto-create tab with headers -----------------------
function logToSheet(ss, sheetName, rowData, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#000000');
    headerRange.setFontColor('#CCFF00');
  }
  sheet.appendRow(rowData);
}
