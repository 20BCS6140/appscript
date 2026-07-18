// ============================================================
// CHATSERVICE.GS — Google Chat Webhook Notifications
// ============================================================

// ============================================================
// CHATSERVICE.GS — Google Chat Webhook Notifications
// ============================================================

function sendChatNotification(doubtId, formData,timestamp) {
  try {
    const violationsStr = formData.violations.join(', ');

    const emailOrLdap = formData.l0Name || '';           // Column B - Email Address or LDAP
  const extensionId = formData.extensionId || '';           // Column E - Extension ID
  const doubtDetails = formData.doubtDetails || '';         // Column L - Doubt Details
  const timestamp = new Date().toLocaleString();
    
    // const message = {
    //   cards: [{
    //     header: {
    //       title: '🔔 New Doubt Submitted',
    //       subtitle: `Doubt ID: ${doubtId}`,
    //       imageUrl: 'https://fonts.gstatic.com/s/i/googlematerialicons/extension/v6/googblue-24dp/1x/gm_extension_googblue_24dp.png'
    //     },
    //     sections: [
    //       {
    //         widgets: [
    //           {
    //             keyValue: {
    //               topLabel: 'Extension ID',
    //               content: formData.extensionId,
    //               icon: 'BOOKMARK'
    //             }
    //           },
    //           {
    //             keyValue: {
    //               topLabel: 'Submitted By',
    //               content: `${formData.l0Name} (${formData.l0Email})`,
    //               icon: 'PERSON'
    //             }
    //           },
    //           {
    //             keyValue: {
    //               topLabel: 'Violations',
    //               content: violationsStr,
    //               icon: 'DESCRIPTION'
    //             }
    //           },
              
    //           {
    //             keyValue: {
    //               topLabel: 'Doubt Details',
    //               content: formData.doubtDetails,
    //               icon: 'DESCRIPTION'
    //             }
    //           }
    //         ]
    //       }
    //     ]
    //   }]
    // };

    const message = `
    📩 New Form Submission - ${timestamp}

    📧 Email Address or LDAP: ${emailOrLdap}
    💻 Extension ID: ${extensionId}
    📝 Doubt Details: ${doubtDetails}
  `;

    // const options = {
    //   method: 'POST',
    //   contentType: 'application/json',
    //   payload: JSON.stringify(message)
    // };

    // UrlFetchApp.fetch(CHAT_WEBHOOK_URL, options);
    MailApp.sendEmail({
    to: "demo",  // Send to the CRX-form-notifier email address
    subject: `📩 New Form Submission - ${timestamp}`,
    body: message  // Plain text message body
  });
  } catch (err) {
    Logger.log('Chat notification error: ' + err.message);
  }
}
