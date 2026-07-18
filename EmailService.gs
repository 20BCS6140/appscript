// ============================================================
// EMAILSERVICE.GS — Gmail Auto-Email to L0 on Resolution
// ============================================================

function sendResolutionEmail(data) {
  try {
    const subject = `[CRX Resolution] Extension ${data.extensionId} — Doubt ${data.doubtId}`;


  const supervisorEmail = data.supervisor ? data.supervisor + '@gmail.com': '';
  const htmlBody = `<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body>
<div style="font-family: 'Segoe UI', Roboto, 'Helvetica Neue', sans-serif; color: #2d2d2d;">
  <p style="font-size: 16px;">Hi <strong>${data.l0Name}</strong>,</p>
  <p style="font-size: 15px;">📬 Please find the final response regarding your query below:</p>
  <table style="border-collapse: collapse; width: 100%; font-size: 14px;">
    <thead>
      <tr style="background-color: #2980b9; color: #ecf0f1;">
        <th style="padding: 12px; border: 1px solid #ddd; text-align: left;">📌 Field</th>
        <th style="padding: 12px; border: 1px solid #ddd; text-align: left;">📄 Details</th>
      </tr>
    </thead>
    <tbody>
      ${createRow('🔑 Reviewer LDAP',                        data.l0Name,             0, 'color: #3498db; font-weight: bold;')}
      ${createRow('📍 POD',                                  data.pod,                1, 'color: #27ae60; font-weight: bold;')}
      ${createRow('👤 Supervisor',                           data.supervisor,         2, 'color: #8e44ad; font-weight: bold;')}
      ${createRow('🆔 Extension ID',                         data.extensionId,        3, '')}
      ${createRow('✅ Final Verdict',                        data.finalVerdict,       4, 'color: green; font-weight: bold;')}
      ${createRow('📝 CRX Team Justification',               data.clarification,      5, '')}
      ${createRow('📈 L0 Area of Improvement',               data.l0Improvement,      6, '')}
      ${createRow('💬 Additional Comments/Recommendations',  data.additionalComments, 7, '')}
      ${createRow('🔍 Type of Consult (Easy/Medium/Complex)',data.typeOfConsult,       8, getTypeOfConsultStyle(data.typeOfConsult))}
      ${createRow('🛠️ Approach Validation',                  data.approachValidation, 9, getApproachValidationStyle(data.approachValidation))}
      ${createRow('👤 CRX Validator',                        data.resolvedBy,        10, 'color: #4B0082; font-weight: bold;')}
    </tbody>
  </table>
  <p style="margin-top: 20px; font-size: 14px;">Regards,<br><strong>CWS ACN CRX Team</strong></p>
</div>
</body>
</html>`;

    const emailOptions = {
      name: 'CRX Team',
      bcc: 'shmodem@google.com, abhilashnukala@google.com',
      htmlBody: htmlBody
    };

    if(supervisorEmail) {
      emailOptions.cc = supervisorEmail;
    }

    GmailApp.sendEmail(data.l0Email,subject,'',emailOptions);

  } catch (err) {
    Logger.log('Email send error: ' + err.message);
  }
}

function createRow(label, value, index, overrideStyle = '') {
  const bgColor = index % 2 === 0 ? '#f4f6f8' : '#e1e8f0';
  const colorStyle = overrideStyle || '';
  
  return ` 
    <tr style="background-color: ${bgColor};">
      <td style="padding: 12px; border: 1px solid #ddd; font-weight: 600;">${label}</td>
      <td style="padding: 12px; border: 1px solid #ddd; ${colorStyle}">${value !== undefined ? value : ''}</td>
    </tr>
  `;
}

function getTypeOfConsultStyle(value) {
  if (!value || typeof value !== 'string') return '';
  const val = value.toLowerCase().trim();
  if (val === 'easy') return 'color: green; font-weight: bold;';
  if (val === 'medium') return 'color: orange; font-weight: bold;';
  if (val === 'complex') return 'color: red; font-weight: bold;';
  return '';
}

function getApproachValidationStyle(value) {
  if (!value || typeof value !== 'string') return '';
  const val = value.toLowerCase().trim();
  if (val === 'valid') return 'color: green; font-weight: bold;';
  if (val === 'invalid') return 'color: red; font-weight: bold;';
  return '';
}
