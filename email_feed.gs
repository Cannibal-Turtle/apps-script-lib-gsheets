/**
 * PayPal → Full RSS 2.0 feed via string building (showing all messages, then marking them read).
 * Wraps title and description in CDATA so apostrophes and other symbols appear correctly.
 * Deploy as Web App (Execute as “Me”, access “Anyone”) and point MonitoRSS at the /exec URL.
 */

const LABEL_NAME  = 'PayPal';   // Your Gmail label for PayPal emails
const MAX_THREADS = 50;         // How many conversation threads to fetch

/**
 * Simple XML‑escape for any non‑CDATA fields
 */
function escapeXml(str) {
  return (str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function doGet() {
  // 1. Fetch up to MAX_THREADS threads labeled “PayPal”
  const label   = GmailApp.getUserLabelByName(LABEL_NAME);
  const threads = label ? label.getThreads(0, MAX_THREADS) : [];

  // 2. Build RSS envelope
  let xml = '<?xml version="1.0" encoding="UTF-8"?>\n'
          + '<rss version="2.0">\n'
          + '  <channel>\n'
          + '    <title>PayPal Notifications</title>\n'
          + '    <link>https://mail.google.com/mail/u/0/</link>\n'
          + '    <description>All PayPal‑labeled emails</description>\n'
          + '    <lastBuildDate>' + (new Date()).toUTCString() + '</lastBuildDate>\n';

  // 3. Iterate each message in each thread
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const subject = msg.getSubject();
      // take first few lines of the plain‑text body
      const body = msg.getPlainBody().trim().split('\n').slice(0, 5).join(' ');
      const date = msg.getDate().toUTCString();

      xml += '    <item>\n'
           + '      <title><![CDATA['       + subject + ']]></title>\n'
           + '      <description><![CDATA[' + body    + ']]></description>\n'
           + '      <pubDate>'               + escapeXml(date) + '</pubDate>\n'
           + '    </item>\n';

      // 4. Mark each email read so it won’t reappear next time
      msg.markRead();
    });
  });

  // 5. Close out channel and rss
  xml += '  </channel>\n'
      + '</rss>';

  // 6. Serve as RSS feed
  return ContentService
    .createTextOutput(xml)
    .setMimeType(ContentService.MimeType.RSS);
}
