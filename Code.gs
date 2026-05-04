// ============================================================
// CODE.GS — Main Router & Core Data Operations
// ============================================================

const SHEET_ID = '1MxVhInv31dg10ZaQRidOv6gUStExG3wf-r8SpsysZ0c'; // ← Paste your Sheet ID
const CHAT_WEBHOOK_URL = 'YOUR_GOOGLE_CHAT_WEBHOOK_URL_HERE'; // ← Paste webhook URL

// ── Entry Point ──────────────────────────────────────────────
function doGet(e) {
  const page = e.parameter.page || 'l0';
  let template;

  if (page === 'l0') {
    template = HtmlService.createTemplateFromFile('l0-form');
  } else if (page === 'crx') {
    template = HtmlService.createTemplateFromFile('crx-dashboard');
  } else if (page === 'analytics') {
    template = HtmlService.createTemplateFromFile('analytics-dashboard');
  } else {
    return HtmlService.createHtmlOutput('<h2>Page not found</h2>');
  }

  return template.evaluate()
    .setTitle('CRX Doubt Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Include HTML partials
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Config Helpers ────────────────────────────────────────────
function getConfig() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();

  const crxMembers = [];
  const queueNames = [];
  const violations = [];

  // Skip header row (row 0), read from row 1 onwards
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() !== '') 
      crxMembers.push(data[i][0].toString().trim());   // Column A

    if (data[i][1] && data[i][1].toString().trim() !== '') 
      queueNames.push(data[i][1].toString().trim());   // Column B ← WAS [2], now [1]

    if (data[i][2] && data[i][2].toString().trim() !== '') 
      violations.push(data[i][2].toString().trim());   // Column C ← WAS [4], now [2]
  }

  Logger.log('CRX Members: ' + crxMembers.length);
  Logger.log('Queue Names: ' + queueNames.length);
  Logger.log('Violations: ' + violations.length);

  return { crxMembers, queueNames, violations };
}

// ── SUBMIT DOUBT (called from L0 form) ───────────────────────
function submitDoubt(formData) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const doubtsSheet = ss.getSheetByName('Doubts');

    const doubtId = 'D-' + new Date().getTime();
    const submittedAt = new Date().toISOString();
    const violationsStr = formData.violations.join(', ');

    doubtsSheet.appendRow([
      doubtId,
      formData.extensionPickupTime,
      formData.extensionId,
      violationsStr,
      formData.queueName,
      formData.doubtDetails,
      formData.l0Name,
      formData.l0Email,
      submittedAt,
      'Open',
      '',  // Assigned_To
      ''   // Assigned_At
    ]);

    // Send Google Chat notification
    sendChatNotification(doubtId, formData);

    return { success: true, doubtId: doubtId };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── GET ALL DOUBTS (for CRX dashboard) ───────────────────────
function getDoubts() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Doubts');

    if (!sheet) {
      Logger.log('Doubts sheet not found. Available sheets: ' + 
        ss.getSheets().map(s => s.getName()).join(', '));
      return [];
    }

    const data = sheet.getDataRange().getValues();
    Logger.log('Doubts rows: ' + data.length);

    if (data.length <= 1) return [];

    const headers = data[0];
    
    return data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        const key = h.toString().trim(); // trim spaces from header names
        let val = row[i];
        
        // Convert Date objects to ISO string so they serialize correctly
        if (val instanceof Date) {
          val = val.toISOString();
        } else if (val === null || val === undefined) {
          val = '';
        } else {
          val = val.toString();
        }
        
        obj[key] = val;
      });
      return obj;
    });

  } catch (err) {
    Logger.log('getDoubts ERROR: ' + err.message);
    return []; // never return null
  }
}

// ── ASSIGN DOUBT ──────────────────────────────────────────────
function assignDoubt(doubtId, memberName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Doubts');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === doubtId) {
        sheet.getRange(i + 1, 11).setValue(memberName);        // Assigned_To (col K)
        sheet.getRange(i + 1, 12).setValue(new Date().toISOString()); // Assigned_At (col L)
        sheet.getRange(i + 1, 10).setValue('Assigned');        // Status (col J)
        return { success: true };
      }
    }
    return { success: false, error: 'Doubt not found' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── RESOLVE DOUBT ─────────────────────────────────────────────
function resolveDoubt(resolveData) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const doubtsSheet = ss.getSheetByName('Doubts');
    const resolvedSheet = ss.getSheetByName('Resolved');
    const data = doubtsSheet.getDataRange().getValues();

    let doubtRow = null;
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === resolveData.doubtId) {
        doubtRow = data[i];
        rowIndex = i + 1;
        break;
      }
    }

    if (!doubtRow) return { success: false, error: 'Doubt not found' };

    const resolvedAt = new Date().toISOString();

    // Log to Resolved sheet
    resolvedSheet.appendRow([
      doubtRow[0],   // Doubt_ID
      doubtRow[1],   // Extension_Pickup_Time
      doubtRow[2],   // Extension_ID
      doubtRow[3],   // Violations
      doubtRow[4],   // Queue_Name
      doubtRow[5],   // Doubt_Details
      doubtRow[6],   // L0_Name
      doubtRow[7],   // L0_Email
      doubtRow[8],   // Submitted_At
      doubtRow[10],  // Assigned_To
      resolveData.clarification,
      resolveData.finalVerdict,
      resolveData.resolvedBy,
      resolvedAt
    ]);

    // Update status in Doubts sheet → "Resolved"
    doubtsSheet.getRange(rowIndex, 10).setValue('Resolved');

    // Send email to L0
    sendResolutionEmail({
      l0Email: doubtRow[7],
      l0Name: doubtRow[6],
      extensionId: doubtRow[2],
      violations: doubtRow[3],
      doubtDetails: doubtRow[5],
      clarification: resolveData.clarification,
      finalVerdict: resolveData.finalVerdict,
      resolvedBy: resolveData.resolvedBy,
      doubtId: doubtRow[0]
    });

    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── ANALYTICS DATA ────────────────────────────────────────────
function getAnalyticsData(filters) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const resolvedSheet = ss.getSheetByName('Resolved');
    const doubtsSheet   = ss.getSheetByName('Doubts');

    const resolvedData = resolvedSheet.getDataRange().getValues();
    const doubtsData   = doubtsSheet.getDataRange().getValues();

    // ── Summary counts from Doubts sheet ──
    const allDoubts  = doubtsData.length > 1 ? doubtsData.slice(1) : [];
    const totalOpen     = allDoubts.filter(r => r[9].toString().trim() === 'Open').length;
    const totalAssigned = allDoubts.filter(r => r[9].toString().trim() === 'Assigned').length;
    const totalResolved = resolvedData.length > 1 ? resolvedData.length - 1 : 0;
    const total         = allDoubts.length;

    // No resolved rows yet
    if (resolvedData.length <= 1) {
      return {
        summary: { total, totalOpen, totalAssigned, totalResolved: 0 },
        byViolation: {}, byCRXMember: {}, trend: []
      };
    }

    const resolvedRows = resolvedData.slice(1);

    // ── Parse filter dates ──
    const dateFrom = filters && filters.dateFrom ? new Date(filters.dateFrom) : null;
    const dateTo   = filters && filters.dateTo
      ? new Date(filters.dateTo + 'T23:59:59') : null;
    const memberFilter    = filters && filters.crxMember !== 'All' ? filters.crxMember : null;
    const violationFilter = filters && filters.violation !== 'All'
      ? filters.violation.toLowerCase().trim() : null;

    Logger.log('Filters — dateFrom: ' + dateFrom + ' | dateTo: ' + dateTo +
               ' | member: ' + memberFilter + ' | violation: ' + violationFilter);

    // ── Filter resolved rows ──
    const filtered = resolvedRows.filter(r => {
      // Col N (index 13) = Resolved_At
      let resolvedAt = r[13];
      if (resolvedAt instanceof Date) {
        resolvedAt = resolvedAt; // already a Date object
      } else {
        resolvedAt = new Date(resolvedAt);
      }

      if (isNaN(resolvedAt)) {
        Logger.log('Invalid date in row: ' + JSON.stringify(r[13]));
        return false;
      }

      if (dateFrom && resolvedAt < dateFrom) return false;
      if (dateTo   && resolvedAt > dateTo)   return false;

      // Col M (index 12) = Resolved_By
      if (memberFilter) {
        const member = r[12].toString().trim();
        if (member !== memberFilter) return false;
      }

      // Col D (index 3) = Violations
      if (violationFilter) {
        const viols = r[3].toString().toLowerCase();
        if (!viols.includes(violationFilter)) return false;
      }

      return true;
    });

    Logger.log('Filtered resolved rows: ' + filtered.length);

    // ── By Violation ──
    const byViolation = {};
    filtered.forEach(r => {
      const viols = r[3].toString().split(',');
      viols.forEach(v => {
        v = v.trim();
        if (v) byViolation[v] = (byViolation[v] || 0) + 1;
      });
    });

    // ── By CRX Member ──
    const byCRXMember = {};
    filtered.forEach(r => {
      const member = r[12].toString().trim() || 'Unassigned';
      byCRXMember[member] = (byCRXMember[member] || 0) + 1;
    });

    // ── Daily Trend ──
    const byDate = {};
    filtered.forEach(r => {
      let d = r[13];
      if (d instanceof Date) {
        d = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        d = d.toString().split('T')[0];
      }
      byDate[d] = (byDate[d] || 0) + 1;
    });

    const trend = Object.entries(byDate)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([date, count]) => ({ date, count }));

    Logger.log('byViolation: ' + JSON.stringify(byViolation));
    Logger.log('byCRXMember: ' + JSON.stringify(byCRXMember));
    Logger.log('trend: ' + JSON.stringify(trend));

    return {
      summary: { total, totalOpen, totalAssigned, totalResolved },
      byViolation,
      byCRXMember,
      trend
    };

  } catch (err) {
    Logger.log('getAnalyticsData ERROR: ' + err.message);
    return {
      summary: { total: 0, totalOpen: 0, totalAssigned: 0, totalResolved: 0 },
      byViolation: {}, byCRXMember: {}, trend: []
    };
  }
}
function testConfig() {
  const config = getConfig();
  Logger.log(JSON.stringify(config, null, 2));
}

function debugDoubts() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  // List ALL sheet names in your spreadsheet
  const allSheets = ss.getSheets().map(s => s.getName());
  Logger.log('All sheets found: ' + JSON.stringify(allSheets));

  // Try to open Doubts sheet
  const sheet = ss.getSheetByName('Doubts');
  Logger.log('Doubts sheet found: ' + (sheet !== null));

  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  Logger.log('Total rows including header: ' + data.length);
  Logger.log('Headers: ' + JSON.stringify(data[0]));
  
  if (data.length > 1) {
    Logger.log('First data row: ' + JSON.stringify(data[1]));
  }
}
