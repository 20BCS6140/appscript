// ============================================================
// Code.GS — Google Chat Webhook Notifications
// ============================================================

const SHEET_ID = '1MxVhInv31dg10ZaQRidOv6gUStExG3wf-r8SpsysZ0c';
const CHAT_WEBHOOK_URL = 'YOUR_WEBHOOK_URL_HERE';

function onSheetChange(e) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ss.getSheetByName('Meta').getRange('A1').setValue(new Date().toISOString());
}

function getLastUpdated() {
  return SpreadsheetApp.openById(SHEET_ID)
    .getSheetByName('Meta').getRange('A1').getValue().toString();
}

function updateMeta() {
  try {
    SpreadsheetApp.openById(SHEET_ID)
      .getSheetByName('Meta').getRange('A1')
      .setValue(new Date().toISOString());
  } catch(e) { Logger.log('Meta update failed: ' + e.message); }
}

// ── Static arrays (no longer from Config sheet) ──────────────
const QUEUE_NAMES    = ['RE', 'NBR', 'HQ'];
const VIOLATIONS_LIST = ['malware', 'spam', 'misleading', 'SPP', 'RHC'];

// ── Entry Point ───────────────────────────────────────────────
function doGet(e) {
  const page = e.parameter.page || 'l0';

  // ── Access control — CRX and Analytics restricted to CRX members only ──
  if (page === 'crx' || page === 'analytics') {
    const auth = checkCRXAccess();
    if (!auth.allowed) {
      return HtmlService.createHtmlOutput(`
        <!DOCTYPE html><html><head>
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <link href="https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;600&display=swap" rel="stylesheet">
          <style>
            * { box-sizing: border-box; margin: 0; padding: 0; }
            body { font-family: 'Google Sans', Arial, sans-serif; background: #f0f4f9;
                   display: flex; align-items: center; justify-content: center;
                   min-height: 100vh; }
            .card { background: white; border-radius: 16px; padding: 48px 40px;
                    text-align: center; box-shadow: 0 2px 16px rgba(0,0,0,0.1);
                    max-width: 420px; width: 90%; }
            .icon { font-size: 56px; margin-bottom: 20px; }
            h2 { color: #d93025; font-size: 22px; margin-bottom: 12px; }
            .msg { color: #5f6368; font-size: 14px; line-height: 1.6; margin-bottom: 8px; }
            .email-badge { display: inline-block; background: #f1f3f4; color: #3c4043;
                           padding: 4px 12px; border-radius: 20px; font-size: 12px;
                           margin: 12px 0 24px; }
            .btn { display: inline-block; background: #1a73e8; color: white;
                   padding: 11px 24px; border-radius: 8px; text-decoration: none;
                   font-size: 14px; font-weight: 500; }
            .btn:hover { background: #1557b0; }
          </style>
        </head><body>
          <div class="card">
            <div class="icon">🔒</div>
            <h2>Access Denied</h2>
            <p class="msg">This page is restricted to <strong>CRX team members</strong> only.</p>
            <p class="msg">If you believe this is a mistake, contact your supervisor.</p>
            <div class="email-badge">Signed in as: ${auth.email || 'unknown'}</div><br>
            <a class="btn" href="${ScriptApp.getService().getUrl()}?page=l0">Go to L0 Form →</a>
          </div>
        </body></html>
      `).setTitle('Access Denied — CRX Doubt Management');
    }
  }

  let template;
  if (page === 'l0')             template = HtmlService.createTemplateFromFile('l0-form');
  else if (page === 'crx')       template = HtmlService.createTemplateFromFile('crx-dashboard');
  else if (page === 'analytics') template = HtmlService.createTemplateFromFile('analytics-dashboard');
  else return HtmlService.createHtmlOutput('<h2>Page not found</h2>');

  return template.evaluate()
    .setTitle('CRX Doubt Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// ── Helper: parse "Xh Ym" → minutes ──────────────────────────
function parseDurationToMinutes(str) {
  if (!str) return 0;
  const s = str.toString();
  const h = s.match(/(\d+)h/); const m = s.match(/(\d+)m/);
  return (h ? parseInt(h[1]) : 0) * 60 + (m ? parseInt(m[1]) : 0);
}



// ── CRX Access Check ──────────────────────────────────────────
function checkCRXAccess() {
  try {
    const userEmail = Session.getActiveUser().getEmail() || '';
    // Config sheet stores LDAPs (e.g. "ajaygajula"), so extract local part of email
    const userLdap  = userEmail.split('@')[0].toLowerCase();

    const ss     = SpreadsheetApp.openById(SHEET_ID);
    const config = ss.getSheetByName('Config');
    const data   = config.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const member = data[i][0] ? data[i][0].toString().trim().toLowerCase() : '';
      if (!member) continue;
      // Match either the LDAP part or the full email
      if (member === userLdap || member === userEmail.toLowerCase()) {
        return { allowed: true, email: userEmail };
      }
    }
    return { allowed: false, email: userEmail };
  } catch (err) {
    Logger.log('checkCRXAccess error: ' + err.message);
    return { allowed: false, email: 'unknown' };
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Config — only CRX members now; queues/violations are static ──
function getConfig() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getDataRange().getValues();
  const crxMembers = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() !== '')
      crxMembers.push(data[i][0].toString().trim());
  }
  return { crxMembers };
}

function getUserEmail() {
  try { return Session.getActiveUser().getEmail() || ''; }
  catch (e) { return ''; }
}

// function getUserLdap() {
//   try {

//   }
//     const L0_Email = Session.getActiveUser().getEmail() || '';
//     // Config sheet stores LDAPs (e.g. "ajaygajula"), so extract local part of email
//     return L0Email.split('@')[0].toLowerCase();
// }

// ── POD + Supervisor lookup ───────────────────────────────────
// Assumes your POD tab is named 'POD' with headers: L0_Email | POD_No | Supervisor
function getPodInfo(email) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const podSheet = ss.getSheetByName('POD'); // ← update tab name if different
    if (!podSheet) return { pod: '', supervisor: '' };
    const data = podSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toLowerCase() === email.trim().toLowerCase()) {
        return {
          pod:        data[i][1].toString().trim(),
          supervisor: data[i][2].toString().trim()
        };
      }
    }
    return { pod: '', supervisor: '' };
  } catch (e) {
    return { pod: '', supervisor: '' };
  }
}

// ── SUBMIT DOUBT ──────────────────────────────────────────────
// Doubts columns (0-indexed):
// 0:Doubt_ID 1:Submitted_At 2:Extension_Pickup_Time 3:L0_Name 4:L0_Email
// 5:Extension_ID 6:revision_id 7:task_id 8:Item's Functionality 9:Queue_Name
// 10:Category 11:Sub Category 12:Stage of query 13:DCR Docs Link
// 14:L0 Approach opinion 15:Doubt_Details 16:Violations 17:Status
// 18:Assigned_To 19:Assigned_At 20:POD 21:Supervisor
function submitDoubt(formData) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const doubtsSheet = ss.getSheetByName('Doubts');
    const doubtId    = 'D-' + Date.now();
    const submittedAt = new Date().toISOString();
    const violationsStr = Array.isArray(formData.violations)
      ? formData.violations.join(', ') : formData.violations;

    const podInfo = getPodInfo(formData.l0Email);

    doubtsSheet.appendRow([
      doubtId,                       // 0
      submittedAt,                   // 1
      formData.extensionPickupTime,  // 2
      formData.l0Name,               // 3
      formData.l0Email,              // 4
      formData.extensionId,          // 5
      formData.revisionId,           // 6
      formData.taskId,               // 7
      formData.itemsFunctionality,   // 8
      formData.queueName,            // 9
      formData.category,             // 10
      formData.subCategory,          // 11
      formData.stageOfQuery,         // 12
      formData.docLink,              // 13
      formData.l0ApproachOpinion,    // 14 ← NEW
      formData.doubtDetails,         // 15
      violationsStr,                 // 16
      'Open',                        // 17
      '',                            // 18 Assigned_To
      '',                            // 19 Assigned_At
      podInfo.pod,                   // 20 ← POD auto-filled
      podInfo.supervisor             // 21 ← Supervisor auto-filled
    ]);

    sendChatNotification(doubtId, formData);
    updateMeta();
    return { success: true, doubtId };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── GET ALL DOUBTS ────────────────────────────────────────────
// function getDoubts() {
//   try {
//     const ss    = SpreadsheetApp.openById(SHEET_ID);
//     const sheet = ss.getSheetByName('Doubts');
//     if (!sheet) return [];
//     const data = sheet.getDataRange().getValues();
//     if (data.length <= 1) return [];
//     const headers = data[0];
//     return data.slice(1).map(row => {
//       const obj = {};
//       headers.forEach((h, i) => {
//         const key = h.toString().trim();
//         let val = row[i];
//         if (val instanceof Date) val = val.toISOString();
//         else if (val === null || val === undefined) val = '';
//         else val = val.toString();
//         obj[key] = val;
//       });
//       return obj;
//     });
//   } catch (err) {
//     Logger.log('getDoubts ERROR: ' + err.message);
//     return [];
//   }
// }


// ── MODIFIED getDoubts — last 30 days only ────────────────────
function getDoubts() {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Doubts');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const headers = data[0];

    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 30); // 30-day window

    return data.slice(1)
      .filter(row => {
        const d = row[1] instanceof Date ? row[1] : new Date(row[1]);
        return !isNaN(d) && d >= cutoff;
      })
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          const key = h.toString().trim();
          let val = row[i];
          if (val instanceof Date) val = val.toISOString();
          else if (val === null || val === undefined) val = '';
          else val = val.toString();
          obj[key] = val;
        });
        return obj;
      });
  } catch (err) {
    Logger.log('getDoubts ERROR: ' + err.message);
    return [];
  }
}

// ── ASSIGN DOUBT ──────────────────────────────────────────────
// Status=col18, Assigned_To=col19, Assigned_At=col20 (1-indexed)
function assignDoubt(doubtId, memberName) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Doubts');
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === doubtId.toString()) {
        sheet.getRange(i + 1, 19).setValue(memberName);
        sheet.getRange(i + 1, 20).setValue(new Date().toISOString());
        sheet.getRange(i + 1, 18).setValue('Assigned');
        updateMeta();
        return { success: true };
      }
    }
    return { success: false, error: 'Doubt not found' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── RESOLVE DOUBT ─────────────────────────────────────────────
// Resolved columns (0-indexed):
// 0:Doubt_ID 1:Submitted_At 2:Extension_Pickup_Time 3:L0_Name 4:L0_Email
// 5:revision_id 6:task_id 7:Item's Functionality 8:Queue_Name 9:Category
// 10:Sub Category 11:Stage of query 12:DCR Docs Link 13:L0 Approach opinion
// 14:Doubt_Details 15:Violations 16:Status 17:Assigned_To 18:Assigned_At
// 19:Type of consult 20:Escalated due to lack of tools 21:Escalation validity
// 22:Was the query clear from L0 23:Is this a new scenario for DCR
// 24:L0 Area of improvement 25:L0 Approach validation 26:Provided Verdict
// 27:L2 Confirmation 28:Clarification Details 29:L2 Additional Comments
// 30:L2 LDAP 31:Resolved_At 32:POD 33:Supervisor 34:Consult AHT 35:Wait Time
function resolveDoubt(resolveData) {
  try {
    const ss           = SpreadsheetApp.openById(SHEET_ID);
    const doubtsSheet  = ss.getSheetByName('Doubts');
    const resolvedSheet= ss.getSheetByName('Resolved');
    const data = doubtsSheet.getDataRange().getValues();

    let doubtRow = null, rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === resolveData.doubtId.toString()) {
        doubtRow = data[i]; rowIndex = i + 1; break;
      }
    }
    if (!doubtRow) return { success: false, error: 'Doubt not found' };

    const resolvedAt   = new Date().toISOString();
    const assignedAt   = doubtRow[19] ? new Date(doubtRow[19]) : null;
    const submittedAt  = doubtRow[1]  ? new Date(doubtRow[1])  : null;
    const resolvedDate = new Date(resolvedAt);

    const consultAHT = (assignedAt && !isNaN(assignedAt))
      ? formatDuration(resolvedDate - assignedAt) : '';
    const waitTime   = (submittedAt && assignedAt && !isNaN(submittedAt) && !isNaN(assignedAt))
      ? formatDuration(assignedAt - submittedAt) : '';

    const l2ConfStr = Array.isArray(resolveData.l2Confirmation)
      ? resolveData.l2Confirmation.join(', ')
      : (resolveData.l2Confirmation || '');

    resolvedSheet.appendRow([
      doubtRow[0],   // Doubt_ID
      doubtRow[1],   // Submitted_At
      doubtRow[2],   // Extension_Pickup_Time
      doubtRow[3],   // L0_Name
      doubtRow[4],   // L0_Email
      doubtRow[5],   // Extension ID 
      doubtRow[6],   // revision_id   (skips Extension_ID at [5])
      doubtRow[7],   // task_id
      doubtRow[8],   // Item's Functionality
      doubtRow[9],   // Queue_Name
      doubtRow[10],  // Category
      doubtRow[11],  // Sub Category
      doubtRow[12],  // Stage of query
      doubtRow[13],  // DCR Docs Link
      doubtRow[14],  // L0 Approach opinion
      doubtRow[15],  // Doubt_Details
      doubtRow[16],  // Violations
      'Resolved',    // Status
      doubtRow[18],  // Assigned_To
      doubtRow[19],  // Assigned_At
      resolveData.typeOfConsult,
      resolveData.escalatedDueToTools,
      resolveData.escalationValidity,
      resolveData.queryClarity,
      resolveData.newScenario,
      resolveData.l0AreaOfImprovement,
      resolveData.l0ApproachValidation,
      resolveData.providedVerdict,
      l2ConfStr,
      resolveData.clarification,
      resolveData.l2AdditionalComments,
      resolveData.resolvedBy,  // L2 LDAP
      resolvedAt,
      doubtRow[20],  // POD
      doubtRow[21],  // Supervisor
      consultAHT,
      waitTime
    ]);

    // Mark resolved in Doubts sheet — Status is col 18 (1-indexed)
    doubtsSheet.getRange(rowIndex, 18).setValue('Resolved');

    sendResolutionEmail({
      l0Email:      doubtRow[4],
      l0Name:       doubtRow[3],
      extensionId:  doubtRow[5],
      violations:   doubtRow[16],
      doubtDetails: doubtRow[15],
      clarification: resolveData.clarification,
      finalVerdict:  resolveData.providedVerdict,
      resolvedBy:    resolveData.resolvedBy,
      doubtId:       doubtRow[0]
    });

    updateMeta();
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── GET A SINGLE RESOLVED DOUBT (for edit overlay) ───────────

function getResolvedDoubt(doubtId) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Resolved');
    const data  = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === doubtId.toString()) {
        const obj = {};
        headers.forEach((h, idx) => {
          let val = data[i][idx];
          if (val instanceof Date) val = val.toISOString();
          else if (val === null || val === undefined) val = '';
          else val = val.toString();
          obj[h.toString().trim()] = val;
        });
        return obj;
      }
    }
    return null;
  } catch (err) {
    Logger.log('getResolvedDoubt ERROR: ' + err.message);
    return null;
  }
}

// ── UPDATE AN EXISTING RESOLVED ROW ─────────────────────────
// ── MODIFIED updateResolvedDoubt — optional email resend ──────
function updateResolvedDoubt(updateData) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Resolved');
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === updateData.doubtId.toString()) {
        const row = i + 1;
        const l2Str = Array.isArray(updateData.l2Confirmation)
          ? updateData.l2Confirmation.join(', ')
          : (updateData.l2Confirmation || '');
        sheet.getRange(row, 21).setValue(updateData.typeOfConsult);
        sheet.getRange(row, 22).setValue(updateData.escalatedDueToTools);
        sheet.getRange(row, 23).setValue(updateData.escalationValidity);
        sheet.getRange(row, 24).setValue(updateData.queryClarity);
        sheet.getRange(row, 25).setValue(updateData.newScenario);
        sheet.getRange(row, 26).setValue(updateData.l0AreaOfImprovement);
        sheet.getRange(row, 27).setValue(updateData.l0ApproachValidation);
        sheet.getRange(row, 28).setValue(updateData.providedVerdict);
        sheet.getRange(row, 29).setValue(l2Str);
        sheet.getRange(row, 30).setValue(updateData.clarification);
        sheet.getRange(row, 31).setValue(updateData.l2AdditionalComments);
        sheet.getRange(row, 32).setValue(updateData.resolvedBy);

        // Optionally re-send email to L0
        if (updateData.resendEmail) {
          sendResolutionEmail({
            l0Email:      data[i][4],
            l0Name:       data[i][3],
            extensionId:  data[i][5],
            violations:   data[i][16],
            doubtDetails: data[i][15],
            clarification: updateData.clarification,
            finalVerdict:  updateData.providedVerdict,
            resolvedBy:    updateData.resolvedBy,
            doubtId:       data[i][0]
          });
        }
        updateMeta();
        return { success: true };
      }
    }
    return { success: false, error: 'Resolved doubt not found' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}


// ── GET TEAM STATUS ──────────────────────────────────────────
function getTeamStatus() {
  try {
    const ss      = SpreadsheetApp.openById(SHEET_ID);
    const today   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    // All CRX members from Config
    const cfgData = ss.getSheetByName('Config').getDataRange().getValues();
    const crxMembers = [];
    for (let i = 1; i < cfgData.length; i++) {
      if (cfgData[i][0] && cfgData[i][0].toString().trim())
        crxMembers.push(cfgData[i][0].toString().trim());
    }

    // Today's presence from Presence sheet
    const presSheet = ss.getSheetByName('Presence');
    const presData  = presSheet ? presSheet.getDataRange().getValues() : [];
    const todayPres = {};
    for (let i = 1; i < presData.length; i++) {
      let rd = presData[i][0];
      rd = rd instanceof Date
        ? Utilities.formatDate(rd, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : rd.toString().split('T')[0];
      if (rd === today) {
        todayPres[presData[i][1].toString().trim()] = {
          status:     presData[i][2].toString().trim().toLowerCase(),
          customNote: presData[i][3] ? presData[i][3].toString().trim() : ''
        };
      }
    }

    // Active doubt assignments (Status === 'Assigned')
    const dData   = ss.getSheetByName('Doubts').getDataRange().getValues();
    const active  = {};
    for (let i = 1; i < dData.length; i++) {
      if (dData[i][17].toString().trim() === 'Assigned') {
        const assignee = dData[i][18].toString().trim();
        if (assignee) {
          if (!active[assignee]) active[assignee] = [];
          active[assignee].push(dData[i][0].toString());
        }
      }
    }

    const team = crxMembers.map(member => {
      const p  = todayPres[member] || { status: 'absent', customNote: '' };
      const ad = active[member]    || [];
      const isPresent       = p.status === 'present';
      const isBusyWithDoubt = ad.length > 0;
      const isBusyOther     = isPresent && p.customNote !== '';
      return {
        member,
        isPresent,
        customNote:     p.customNote,
        activeDoubts:   ad,
        isBusyWithDoubt,
        isBusyOther,
        isAvailable: isPresent && !isBusyWithDoubt && !isBusyOther
      };
    });

    return { team, today };
  } catch (err) {
    Logger.log('getTeamStatus ERROR: ' + err.message);
    return { team: [], today: '' };
  }
}


// ── MARK PRESENCE ─────────────────────────────────────────────
function markPresence(memberName, status, customNote) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('Presence');
    if (!sheet) return { success: false, error: 'Presence sheet not found. Create it first.' };

    const today    = new Date();
    const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const data     = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      let rd = data[i][0];
      rd = rd instanceof Date
        ? Utilities.formatDate(rd, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : rd.toString().split('T')[0];
      if (rd === todayStr && data[i][1].toString().trim() === memberName.trim()) {
        sheet.getRange(i + 1, 3).setValue(status);
        sheet.getRange(i + 1, 4).setValue(customNote || '');
        updateMeta();
        return { success: true };
      }
    }
    sheet.appendRow([today, memberName, status, customNote || '']);
    updateMeta();
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function formatDuration(ms) {
  if (!ms || ms < 0) return '0h 0m';
  const totalMins = Math.floor(ms / 60000);
  return Math.floor(totalMins / 60) + 'h ' + (totalMins % 60) + 'm';
}

// ── ANALYTICS DATA ────────────────────────────────────────────
// Doubts: Status=index17, Violations=index16
// Resolved: L2 LDAP=index30, Resolved_At=index31
// ── MODIFIED getAnalyticsData ─────────────────────────────────
// Uses L2 Confirmation (col 28) for violations
// Adds memberStats + violationCounts tables
function getAnalyticsData(filters) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const resolvedSheet = ss.getSheetByName('Resolved');
    const doubtsSheet   = ss.getSheetByName('Doubts');
    const resolvedData  = resolvedSheet.getDataRange().getValues();
    const doubtsData    = doubtsSheet.getDataRange().getValues();
    const allDoubts     = doubtsData.length > 1 ? doubtsData.slice(1) : [];

    const totalOpen     = allDoubts.filter(r => r[17].toString().trim() === 'Open').length;
    const totalAssigned = allDoubts.filter(r => r[17].toString().trim() === 'Assigned').length;
    const totalResolved = resolvedData.length > 1 ? resolvedData.length - 1 : 0;
    const total         = allDoubts.length;

    const empty = { summary: { total, totalOpen, totalAssigned, totalResolved: 0 },
                    byViolation:{}, byCRXMember:{}, trend:[], memberStats:[], violationCounts:[] };

    if (resolvedData.length <= 1) return empty;

    const resolvedRows = resolvedData.slice(1);
    const dateFrom = filters?.dateFrom ? new Date(filters.dateFrom) : null;
    const dateTo   = filters?.dateTo   ? new Date(filters.dateTo + 'T23:59:59') : null;
    const memberFilter    = filters?.crxMember !== 'All' ? filters.crxMember : null;
    const violationFilter = filters?.violation !== 'All' ? filters.violation?.toLowerCase().trim() : null;

    const filtered = resolvedRows.filter(r => {
      let d = r[32]; // Resolved_At
      if (!(d instanceof Date)) d = new Date(d);
      if (isNaN(d)) return false;
      if (dateFrom && d < dateFrom) return false;
      if (dateTo   && d > dateTo)   return false;
      if (memberFilter && r[31].toString().trim() !== memberFilter) return false;
      if (violationFilter) {
        const l2 = (r[28] || '').toString().toLowerCase();
        if (!l2.includes(violationFilter)) return false;
      }
      return true;
    });

    // ── Violations from L2 Confirmation (col 28) ──
    const byViolation = {};
    filtered.forEach(r => {
      (r[28] || '').toString().split(',').forEach(v => {
        v = v.trim(); if (v) byViolation[v] = (byViolation[v] || 0) + 1;
      });
    });

    // ── By CRX Member ──
    const byCRXMember = {};
    filtered.forEach(r => {
      const m = r[31].toString().trim() || 'Unassigned';
      byCRXMember[m] = (byCRXMember[m] || 0) + 1;
    });

    // ── Daily Trend ──
    const byDate = {};
    filtered.forEach(r => {
      let d = r[32];
      d = d instanceof Date
        ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : d.toString().split('T')[0];
      if (d) byDate[d] = (byDate[d] || 0) + 1;
    });
    const trend = Object.entries(byDate)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([date, count]) => ({ date, count }));

    // ── Member Stats Table ──
    const msMap = {};
    filtered.forEach(r => {
      const m  = r[31].toString().trim() || 'Unassigned';
      const ct = r[20].toString().trim().toLowerCase();
      const esc = r[21].toString().trim().toLowerCase() === 'yes';
      if (!msMap[m]) msMap[m] = { total:0, easy:0, medium:0, complex:0, escalated:0, aht:0, wait:0 };
      msMap[m].total++;
      if (ct === 'easy') msMap[m].easy++;
      else if (ct === 'medium') msMap[m].medium++;
      else if (ct === 'complex') msMap[m].complex++;
      if (esc) msMap[m].escalated++;
      msMap[m].aht  += parseDurationToMinutes(r[35] ? r[35].toString() : '');
      msMap[m].wait += parseDurationToMinutes(r[36] ? r[36].toString() : '');
    });
    const memberStats = Object.entries(msMap)
      .sort((a, b) => b[1].total - a[1].total)
      .map(([member, s]) => ({
        member, total: s.total, easy: s.easy, medium: s.medium,
        complex: s.complex, escalated: s.escalated,
        avgAHT:  s.total > 0 ? formatDuration((s.aht  / s.total) * 60000) : '—',
        avgWait: s.total > 0 ? formatDuration((s.wait / s.total) * 60000) : '—'
      }));

    // ── Violation Counts Table ──
    const vcMap = {};
    filtered.forEach(r => {
      (r[28] || '').toString().split(',').forEach(v => {
        v = v.trim(); if (v) vcMap[v] = (vcMap[v] || 0) + 1;
      });
    });
    const violationCounts = Object.entries(vcMap)
      .sort((a, b) => b[1] - a[1])
      .map(([violation, count]) => ({ violation, count }));

    return { summary: { total, totalOpen, totalAssigned, totalResolved },
             byViolation, byCRXMember, trend, memberStats, violationCounts };
  } catch (err) {
    Logger.log('getAnalyticsData ERROR: ' + err.message);
    return { summary:{total:0,totalOpen:0,totalAssigned:0,totalResolved:0},
             byViolation:{}, byCRXMember:{}, trend:[], memberStats:[], violationCounts:[] };
  }
}

// // ============================================================
// // CODE.GS — Main Router & Core Data Operations
// // ============================================================

// const SHEET_ID = '1MxVhInv31dg10ZaQRidOv6gUStExG3wf-r8SpsysZ0c'; // ← Paste your Sheet ID
// const CHAT_WEBHOOK_URL = 'YOUR_GOOGLE_CHAT_WEBHOOK_URL_HERE'; // ← Paste webhook URL

// // ── Entry Point ──────────────────────────────────────────────
// function doGet(e) {
//   const page = e.parameter.page || 'l0';
//   let template;

//   if (page === 'l0') {
//     template = HtmlService.createTemplateFromFile('l0-form');
//   } else if (page === 'crx') {
//     template = HtmlService.createTemplateFromFile('crx-dashboard');
//   } else if (page === 'analytics') {
//     template = HtmlService.createTemplateFromFile('analytics-dashboard');
//   } else {
//     return HtmlService.createHtmlOutput('<h2>Page not found</h2>');
//   }

//   return template.evaluate()
//     .setTitle('CRX Doubt Management')
//     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
//     .addMetaTag('viewport', 'width=device-width, initial-scale=1');
// }

// // Include HTML partials
// function include(filename) {
//   return HtmlService.createHtmlOutputFromFile(filename).getContent();
// }

// // ── Config Helpers ────────────────────────────────────────────
// function getConfig() {
//   const ss = SpreadsheetApp.openById(SHEET_ID);
//   const configSheet = ss.getSheetByName('Config');
//   const data = configSheet.getDataRange().getValues();

//   const crxMembers = [];
//   const queueNames = [];
//   const violations = [];

//   // Skip header row (row 0), read from row 1 onwards
//   for (let i = 1; i < data.length; i++) {
//     if (data[i][0] && data[i][0].toString().trim() !== '') 
//       crxMembers.push(data[i][0].toString().trim());   // Column A

//     if (data[i][1] && data[i][1].toString().trim() !== '') 
//       queueNames.push(data[i][1].toString().trim());   // Column B ← WAS [2], now [1]

//     if (data[i][2] && data[i][2].toString().trim() !== '') 
//       violations.push(data[i][2].toString().trim());   // Column C ← WAS [4], now [2]
//   }

//   Logger.log('CRX Members: ' + crxMembers.length);
//   Logger.log('Queue Names: ' + queueNames.length);
//   Logger.log('Violations: ' + violations.length);

//   return { crxMembers, queueNames, violations };
// }

// //get Email Automatically
// function getUserEmail() {
//   try {
//     return Session.getActiveUser().getEmail() || "";
//   } catch (e) {
//     return ""; // Fallback if browser permission scopes restrict it
//   }
// }

// // ── SUBMIT DOUBT (called from L0 form) ───────────────────────
// function submitDoubt(formData) {
//   try {
//     const ss = SpreadsheetApp.openById(SHEET_ID);
//     const doubtsSheet = ss.getSheetByName('Doubts');

//     const doubtId = 'D-'+Date.now();
//     const submittedAt = new Date().toLocaleString();;
//     const violationsStr = formData.violations.join(', ');

//     doubtsSheet.appendRow([
//       doubtId,
//       submittedAt,
//       formData.extensionPickupTime,
//       formData.l0Name,
//       formData.l0Email,
//       formData.extensionId,
//       formData.revisionId,            // col G - revision_id
//       formData.taskId,                // col H - task_id
//       formData.itemsFunctionality,    // col I - Item's Functionality
//       formData.queueName,             // col J - Queue Name
//       formData.category,              // col K - Category
//       formData.subCategory,           // col L - Sub Category
//       formData.stageOfQuery,          // col M - Stage of query
//       formData.docLink, 
//       formData.doubtDetails,
//       violationsStr,
//       'Open',
//       '',  // Assigned_To
//       ''   // Assigned_At
//     ]);

//     // Send Google Chat notification
//     sendChatNotification(doubtId, formData,doubtId);

//     return { success: true, doubtId: doubtId };
//   } catch (err) {
//     return { success: false, error: err.message };
//   }
// }

// // ── GET ALL DOUBTS (for CRX dashboard) ───────────────────────
// function getDoubts() {
//   try {
//     const ss = SpreadsheetApp.openById(SHEET_ID);
//     const sheet = ss.getSheetByName('Doubts');

//     if (!sheet) {
//       Logger.log('Doubts sheet not found. Available sheets: ' + 
//         ss.getSheets().map(s => s.getName()).join(', '));
//       return [];
//     }

//     const data = sheet.getDataRange().getValues();
//     Logger.log('Doubts rows: ' + data.length);

//     if (data.length <= 1) return [];

//     const headers = data[0];
    
//     return data.slice(1).map(row => {
//       const obj = {};
//       headers.forEach((h, i) => {
//         const key = h.toString().trim(); // trim spaces from header names
//         let val = row[i];
        
//         // Convert Date objects to ISO string so they serialize correctly
//         if (val instanceof Date) {
//           val = val.toISOString();
//         } else if (val === null || val === undefined) {
//           val = '';
//         } else {
//           val = val.toString();
//         }
        
//         obj[key] = val;
//       });
//       return obj;
//     });

//   } catch (err) {
//     Logger.log('getDoubts ERROR: ' + err.message);
//     return []; // never return null
//   }
// }

// // ── ASSIGN DOUBT ──────────────────────────────────────────────
// function assignDoubt(doubtId, memberName) {
//   try {
//     const ss = SpreadsheetApp.openById(SHEET_ID);
//     const sheet = ss.getSheetByName('Doubts');
//     const data = sheet.getDataRange().getValues();

//     for (let i = 1; i < data.length; i++) {
//       if (data[i][0] === doubtId) {
//         sheet.getRange(i + 1, 18).setValue(memberName);        // Assigned_To (col K)
//         sheet.getRange(i + 1, 19).setValue(new Date().toLocaleString()); // Assigned_At (col L)
//         sheet.getRange(i + 1, 17).setValue('Assigned');        // Status (col J)
//         return { success: true };
//       }
//     }
//     return { success: false, error: 'Doubt not found' };
//   } catch (err) {
//     return { success: false, error: err.message };
//   }
// }

// // ── RESOLVE DOUBT ─────────────────────────────────────────────
// function resolveDoubt(resolveData) {
//   try {
//     const ss = SpreadsheetApp.openById(SHEET_ID);
//     const doubtsSheet = ss.getSheetByName('Doubts');
//     const resolvedSheet = ss.getSheetByName('Resolved');
//     const data = doubtsSheet.getDataRange().getValues();

//     let doubtRow = null;
//     let rowIndex = -1;

//     for (let i = 1; i < data.length; i++) {
//       if (data[i][0] === resolveData.doubtId) {
//         doubtRow = data[i];
//         rowIndex = i + 1;
//         break;
//       }
//     }

//     if (!doubtRow) return { success: false, error: 'Doubt not found' };

//     const resolvedAt = new Date().toLocaleString();

//     // Log to Resolved sheet
//     resolvedSheet.appendRow([
//       doubtRow[0],   
//       doubtRow[1],   // 
//       doubtRow[2],   // 
//       doubtRow[3],   // 
//       doubtRow[4],   // 
//       doubtRow[5],   // 
//       doubtRow[6],   // 
//       doubtRow[7],   // 
//       doubtRow[8],   // 
//       doubtRow[9],  //
//       doubtRow[10],
//       doubtRow[11],
//       doubtRow[12],
//       doubtRow[13],
//       doubtRow[14],
//       doubtRow[15],
//       doubtRow[16],
//       doubtRow[17],
//       doubtRow[18],
//       resolveData.clarification,
//       resolveData.finalVerdict,
//       resolveData.resolvedBy,
//       resolvedAt
//     ]);

//     // Update status in Doubts sheet → "Resolved"
//     doubtsSheet.getRange(rowIndex, 17).setValue('Resolved');

//     // Send email to L0
//     sendResolutionEmail({
//       l0Name:       doubtRow[3],       // L0_Name ✓
//       l0Email:      doubtRow[4],       // L0_Email ✓
//       extensionId:  doubtRow[5],       // Extension_ID ✓
//       // doubtDetails: doubtRow[14],
//       clarification: resolveData.clarification,
//       finalVerdict: resolveData.finalVerdict,
//       resolvedBy: resolveData.resolvedBy,
//       doubtId: doubtRow[0]
//     });

//     return { success: true };
//   } catch (err) {
//     return { success: false, error: err.message };
//   }
// }

// // ── ANALYTICS DATA ────────────────────────────────────────────
// // ── ANALYTICS DATA ────────────────────────────────────────────
// function getAnalyticsData(filters) {
//   try {
//     const ss = SpreadsheetApp.openById(SHEET_ID);
//     const resolvedSheet = ss.getSheetByName('Resolved');
//     const doubtsSheet   = ss.getSheetByName('Doubts');

//     const resolvedData = resolvedSheet.getDataRange().getValues();
//     const doubtsData   = doubtsSheet.getDataRange().getValues();

//     const allDoubts  = doubtsData.length > 1 ? doubtsData.slice(1) : [];
//     const totalOpen      = allDoubts.filter(r => r[16].toString().trim() === 'Open').length;
//     const totalAssigned = allDoubts.filter(r => r[16].toString().trim() === 'Assigned').length;
//     const totalResolved = resolvedData.length > 1 ? resolvedData.length - 1 : 0;
//     const total         = allDoubts.length;

//     // ── FIX: Map Doubt_ID -> Violations from Doubts sheet (index 7) ──
//     const doubtIdToViolations = {};
//     allDoubts.forEach(r => {
//       const dId = r[0].toString().trim();
//       if (dId) {
//         doubtIdToViolations[dId] = r[15] ? r[15].toString().trim() : '';
//       }
//     });

//     if (resolvedData.length <= 1) {
//       return {
//         summary: { total, totalOpen, totalAssigned, totalResolved: 0 },
//         byViolation: {}, byCRXMember: {}, trend: []
//       };
//     }

//     const resolvedRows = resolvedData.slice(1);

//     const dateFrom = filters && filters.dateFrom ? new Date(filters.dateFrom) : null;
//     const dateTo   = filters && filters.dateTo ? new Date(filters.dateTo + 'T23:59:59') : null;
//     const memberFilter    = filters && filters.crxMember !== 'All' ? filters.crxMember : null;
//     const violationFilter = filters && filters.violation !== 'All' ? filters.violation.toLowerCase().trim() : null;

//     // ── Filter resolved rows ──
//     const filtered = resolvedRows.filter(r => {
//       // Index 13 = Assigned_At in Resolved tab
//       let metricDate = r[22];
//       if (!(metricDate instanceof Date)) {
//         metricDate = new Date(metricDate);
//       }

//       if (isNaN(metricDate)) return false;
//       if (dateFrom && metricDate < dateFrom) return false;
//       if (dateTo   && metricDate > dateTo)   return false;

//       // Index 12 = L2 LDAP (Resolved By)
//       if (memberFilter) {
//         const member = r[21].toString().trim();
//         if (member !== memberFilter) return false;
//       }

//       // FIX: Use cross-reference lookup map for the filters
//       if (violationFilter) {
//         const doubtId = r[0].toString().trim();
//         const viols = (doubtIdToViolations[doubtId] || '').toLowerCase();
//         if (!viols.includes(violationFilter)) return false;
//       }

//       return true;
//     });

//     // ── FIX: Populate metrics using actual mapped Violations ──
//     const byViolation = {};
//     filtered.forEach(r => {
//       const doubtId = r[0].toString().trim();
//       const violationsStr = doubtIdToViolations[doubtId] || '';
//       if (violationsStr) {
//         const viols = violationsStr.split(',');
//         viols.forEach(v => {
//           v = v.trim();
//           if (v) byViolation[v] = (byViolation[v] || 0) + 1;
//         });
//       }
//     });

//     // ── By CRX Member ──
//     const byCRXMember = {};
//     filtered.forEach(r => {
//       const member = r[21].toString().trim() || 'Unassigned';
//       byCRXMember[member] = (byCRXMember[member] || 0) + 1;
//     });

//     // ── Daily Trend ──
//     const byDate = {};
//     filtered.forEach(r => {
//       let d = r[22];
//       if (d instanceof Date) {
//         d = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
//       } else {
//         d = d.toString().split('T')[0];
//       }
//       if (d) byDate[d] = (byDate[d] || 0) + 1;
//     });

//     const trend = Object.entries(byDate)
//       .sort((a, b) => a[0].localeCompare(b[0]))
//       .map(([date, count]) => ({ date, count }));

//     return {
//       summary: { total, totalOpen, totalAssigned, totalResolved },
//       byViolation,
//       byCRXMember,
//       trend
//     };

//   } catch (err) {
//     Logger.log('getAnalyticsData ERROR: ' + err.message);
//     return {
//       summary: { total: 0, totalOpen: 0, totalAssigned: 0, totalResolved: 0 },
//       byViolation: {}, byCRXMember: {}, trend: []
//     };
//   }
// }
// function testConfig() {
//   const config = getConfig();
//   Logger.log(JSON.stringify(config, null, 2));
// }

// function debugDoubts() {
//   const ss = SpreadsheetApp.openById(SHEET_ID);
  
//   // List ALL sheet names in your spreadsheet
//   const allSheets = ss.getSheets().map(s => s.getName());
//   Logger.log('All sheets found: ' + JSON.stringify(allSheets));

//   // Try to open Doubts sheet
//   const sheet = ss.getSheetByName('Doubts');
//   Logger.log('Doubts sheet found: ' + (sheet !== null));

//   if (!sheet) return;

//   const data = sheet.getDataRange().getValues();
//   Logger.log('Total rows including header: ' + data.length);
//   Logger.log('Headers: ' + JSON.stringify(data[0]));
  
//   if (data.length > 1) {
//     Logger.log('First data row: ' + JSON.stringify(data[1]));
//   }
// }
