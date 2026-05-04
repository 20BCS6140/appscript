// ============================================================
// EMAILSERVICE.GS — Gmail Auto-Email to L0 on Resolution
// ============================================================

function sendResolutionEmail(data) {
  try {
    const subject = `[CRX Resolution] Extension ${data.extensionId} — Doubt ${data.doubtId}`;

    const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Google Sans', Arial, sans-serif; color: #202124; margin: 0; padding: 0; }
    .container { max-width: 600px; margin: 0 auto; }
    .header { background: #1a73e8; padding: 24px 32px; }
    .header h1 { color: white; margin: 0; font-size: 20px; }
    .header p { color: #e8f0fe; margin: 4px 0 0; font-size: 13px; }
    .body { padding: 24px 32px; background: #fff; }
    .field { margin-bottom: 20px; }
    .field label { font-size: 11px; font-weight: 600; color: #5f6368; text-transform: uppercase; letter-spacing: 0.5px; display: block; margin-bottom: 4px; }
    .field .value { font-size: 14px; color: #202124; background: #f8f9fa; padding: 10px 14px; border-radius: 6px; border-left: 3px solid #1a73e8; }
    .verdict { background: #e6f4ea; border-left: 3px solid #34a853 !important; }
    .footer { background: #f8f9fa; padding: 16px 32px; border-top: 1px solid #e8eaed; }
    .footer p { font-size: 12px; color: #80868b; margin: 0; }
    .badge { display: inline-block; background: #e8f0fe; color: #1a73e8; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 500; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>✅ Your Doubt Has Been Resolved</h1>
      <p>Chrome Extensions Review (CRX) Team Response</p>
    </div>
    <div class="body">
      <p>Hi <strong>${data.l0Name}</strong>,</p>
      <p>Your doubt has been reviewed and resolved by the CRX team. Below are the details:</p>

      <div class="field">
        <label>Doubt ID</label>
        <div class="value"><span class="badge">${data.doubtId}</span></div>
      </div>

      <div class="field">
        <label>Extension ID</label>
        <div class="value">${data.extensionId}</div>
      </div>

      <div class="field">
        <label>Violations Flagged</label>
        <div class="value">${data.violations}</div>
      </div>

      <div class="field">
        <label>Your Original Doubt</label>
        <div class="value">${data.doubtDetails}</div>
      </div>

      <div class="field">
        <label>CRX Clarification</label>
        <div class="value">${data.clarification}</div>
      </div>

      <div class="field">
        <label>Final Verdict</label>
        <div class="value verdict"><strong>${data.finalVerdict}</strong></div>
      </div>

      <div class="field">
        <label>Resolved By</label>
        <div class="value">👤 ${data.resolvedBy}</div>
      </div>
    </div>
    <div class="footer">
      <p>This is an automated message from the CRX Doubt Management System. Please do not reply to this email.</p>
      <p style="margin-top:8px;">Chrome Extensions Review Team · Google</p>
    </div>
  </div>
</body>
</html>`;

    GmailApp.sendEmail(data.l0Email, subject, 
      `Your doubt (${data.doubtId}) for Extension ${data.extensionId} has been resolved by ${data.resolvedBy}. Final verdict: ${data.finalVerdict}`,
      { htmlBody: htmlBody, name: 'CRX Team' }
    );

  } catch (err) {
    Logger.log('Email send error: ' + err.message);
  }
}
