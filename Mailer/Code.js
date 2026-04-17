/**
 * Configuration Constants
 * Update these values if your data moves to different columns or you want a different email.
 */
const LINKS_COL = 'A';
const EMAILS_COL = 'B';
const EMAIL_SUBJECT = 'Files Shared with You';
const EMAIL_GREETING = 'Hello,<br><br>Please find the requested files attached as PDFs. You can also access the live versions here:';
const EMAIL_SIGN_OFF = '<br>Best regards.';

/**
 * Creates a custom menu when the spreadsheet is opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Tools 💌')
      .addItem('Send All Files to All Emails', 'sendAllFilesToAllEmails')
      .addToUi();
}

/**
 * Collects all files from the links column and sends them to all emails in the emails column.
 */
function sendAllFilesToAllEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    ui.alert('Sheet has no data rows.');
    return;
  }

  const linksData = sheet.getRange(`${LINKS_COL}2:${LINKS_COL}${lastRow}`).getValues();
  const emailsData = sheet.getRange(`${EMAILS_COL}2:${EMAILS_COL}${lastRow}`).getValues();
  const myEmail = Session.getActiveUser().getEmail();

  // Check daily quota before doing any work
  if (MailApp.getRemainingDailyQuota() < 1) {
    ui.alert('Cannot send email: daily Gmail quota exhausted.');
    return;
  }

  const attachments = [];
  const fileLinksHtml = [];
  const recipientEmails = [];
  const failedUrls = [];

  // 1. Collect all valid emails
  for (const row of emailsData) {
    const emailAddress = String(row[0]).trim();
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emailAddress)) {
      recipientEmails.push(emailAddress);
    }
  }

  // 2. Collect all valid files and convert to PDF
  for (const row of linksData) {
    const fileUrl = row[0];
    if (!fileUrl || typeof fileUrl !== 'string') continue;

    const fileId = extractFileId_(fileUrl);
    if (!fileId) continue;

    try {
      const file = DriveApp.getFileById(fileId);
      attachments.push(file.getAs(MimeType.PDF));
      fileLinksHtml.push(`<li><a href="${fileUrl}">${file.getName()}</a></li>`);
    } catch (e) {
      Logger.log(`Error fetching file for URL ${fileUrl}: ${e.message}`);
      failedUrls.push(fileUrl);
    }
  }

  // 3. Verification checks before sending
  if (attachments.length === 0) {
    ui.alert(`No valid files found in column ${LINKS_COL}.`);
    return;
  }
  if (recipientEmails.length === 0) {
    ui.alert(`No valid emails found in column ${EMAILS_COL}.`);
    return;
  }

  // 4. Construct the email body
  const htmlBody = `${EMAIL_GREETING}<ul>${fileLinksHtml.join('')}</ul>${EMAIL_SIGN_OFF}`;

  // 5. Send the single consolidated email
  try {
    MailApp.sendEmail({
      to: myEmail,
      bcc: recipientEmails.join(','),
      subject: EMAIL_SUBJECT,
      htmlBody: htmlBody,
      attachments: attachments
    });

    const failedNote = failedUrls.length > 0
      ? `\n\nSkipped ${failedUrls.length} file(s) that could not be converted:\n${failedUrls.join('\n')}`
      : '';
    ui.alert(`Success! Sent ${attachments.length} file(s) to ${recipientEmails.length} recipient(s).${failedNote}`);

  } catch (e) {
    ui.alert(`Failed to send email: ${e.message}`);
  }
}

/**
 * Helper function to extract the Google Drive File ID from a standard URL.
 * Handles both /d/<id>/ and id=<id> URL formats.
 */
function extractFileId_(url) {
  const match = url.match(/(?:\/d\/|[?&]id=)([a-zA-Z0-9_-]{25,})/);
  return match ? match[1] : null;
}
