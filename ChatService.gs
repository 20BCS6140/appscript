// ============================================================
// CHATSERVICE.GS — Google Chat Webhook Notifications
// ============================================================

function sendChatNotification(doubtId, formData) {
  try {
    const violationsStr = formData.violations.join(', ');
    
    const message = {
      cards: [{
        header: {
          title: '🔔 New Doubt Submitted',
          subtitle: `Doubt ID: ${doubtId}`,
          imageUrl: 'https://fonts.gstatic.com/s/i/googlematerialicons/extension/v6/googblue-24dp/1x/gm_extension_googblue_24dp.png'
        },
        sections: [
          {
            widgets: [
              {
                keyValue: {
                  topLabel: 'Extension ID',
                  content: formData.extensionId,
                  icon: 'BOOKMARK'
                }
              },
              {
                keyValue: {
                  topLabel: 'Submitted By',
                  content: `${formData.l0Name} (${formData.l0Email})`,
                  icon: 'PERSON'
                }
              },
              {
                keyValue: {
                  topLabel: 'Queue',
                  content: formData.queueName,
                  icon: 'MULTIPLE_PEOPLE'
                }
              },
              {
                keyValue: {
                  topLabel: 'Violations',
                  content: violationsStr,
                  icon: 'DESCRIPTION'
                }
              },
              {
                keyValue: {
                  topLabel: 'Pickup Time',
                  content: formData.extensionPickupTime,
                  icon: 'CLOCK'
                }
              },
              {
                keyValue: {
                  topLabel: 'Doubt Details',
                  content: formData.doubtDetails,
                  icon: 'DESCRIPTION'
                }
              }
            ]
          },
          {
            widgets: [{
              buttons: [{
                textButton: {
                  text: 'Open CRX Dashboard →',
                  onClick: {
                    openLink: {
                      url: ScriptApp.getService().getUrl() + '?page=crx'
                    }
                  }
                }
              }]
            }]
          }
        ]
      }]
    };

    const options = {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(message)
    };

    UrlFetchApp.fetch(CHAT_WEBHOOK_URL, options);
  } catch (err) {
    Logger.log('Chat notification error: ' + err.message);
  }
}
