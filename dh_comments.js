/**
 * DH Comments → Full RSS 2.0 feed via string building (showing both read & unread).
 * Deploy as Web App and call its URL to fetch your public RSS.
 */

const LABEL_NAME = 'DH Comments';  // your Gmail label

/**
 * Escape XML special chars in text
 */
function escapeXml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function doGet() {
  // fetch the label and its most recent 20 threads
  const label   = GmailApp.getUserLabelByName(LABEL_NAME);
  const threads = label ? label.getThreads(0, 20) : [];

  // XML header + <rss> open with namespaces
  let xml = '<?xml version="1.0" encoding="UTF-8"?>\n' +
    '<rss version="2.0"\n' +
    '     xmlns:content="http://purl.org/rss/1.0/modules/content/"\n' +
    '     xmlns:dc="http://purl.org/dc/elements/1.1/"\n' +
    '     xmlns:atom="http://www.w3.org/2005/Atom">\n' +
    '  <channel>\n' +
    '    <title>DH Comments Feed</title>\n' +
    '    <link>https://mail.google.com/mail/u/0/</link>\n' +
    '    <description>New Dragonholic comment notifications</description>\n' +
    '    <lastBuildDate>' + (new Date()).toUTCString() + '</lastBuildDate>\n';

  threads.forEach(thread => {
    // get every message in the thread, regardless of read status
    const msgs = thread.getMessages();
    msgs.forEach(msg => {
      const subject = msg.getSubject();
      const body    = msg.getPlainBody();
      const date    = msg.getDate();

      // parse title & chapter from subject
      const subjMatch = subject.match(
        /\[Dragonholic\] Please moderate: "(.*)" - Chapter (\d+)/
      );
      const storyTitle = subjMatch ? subjMatch[1] : subject;
      const chapterNum = subjMatch ? subjMatch[2] : '';

      // parse author, comment, URLs from plain‑text
      const authMatch    = body.match(/Author:\s*([^\s]+)/);
      const commMatch    = body.match(/Comment:\s*([^\r\n]+)/);
      const approveMatch = body.match(
        /https:\/\/dragonholic\.com\/wp-admin\/comment\.php\?action=approve&c=\d+#wpbody-content/
      );
      const trashMatch   = body.match(
        /https:\/\/dragonholic\.com\/wp-admin\/comment\.php\?action=trash&c=\d+#wpbody-content/
      );

      const authorText  = authMatch    ? authMatch[1].trim() : '';
      const commentText = commMatch    ? commMatch[1].trim()  : '';
      const approveUrl  = approveMatch ? approveMatch[0]      : '';
      const trashUrl    = trashMatch   ? trashMatch[0]        : '';

      xml += '    <item>\n' +
        '      <title>'   + escapeXml(storyTitle) + '</title>\n' +
        '      <chapter>Chapter ' + escapeXml(chapterNum) + '</chapter>\n' +
        '      <dc:creator><![CDATA[' + authorText + ']]></dc:creator>\n' +
        '      <description><![CDATA[' + commentText + ']]></description>\n' +
        '      <approve_url>' + escapeXml(approveUrl) + '</approve_url>\n' +
        '      <trash_url>'   + escapeXml(trashUrl)   + '</trash_url>\n' +
        '      <pubDate>'     + date.toUTCString()    + '</pubDate>\n' +
        '    </item>\n';
    });
  });

  xml += '  </channel>\n</rss>';

  return ContentService
    .createTextOutput(xml)
    .setMimeType(ContentService.MimeType.RSS);
}
