/***************************************
 * ENR Records Intake
 *
 * Looks for inbox emails labeled ENR_RECORDS,
 * uploads non-prodeo_ attachments to Drive,
 * logs received records in the sheet,
 * adds email body as a note on the subject cell,
 * and alerts Google Chat when files are uploaded.
 ***************************************/

const ENR_RECORDS_CONFIG = {
  gmailLabelName: 'ENR_RECORDS',
  processedLabelName: 'ENR_RECORDS_PROCESSED',

  destinationFolderId: '1lCJhLvAzT4p-GyMzbd9DU6bT8BQ7Z36n',
  destinationSheetName: 'Records Received',

  chatWebhookUrl: 'https://chat.googleapis.com/v1/spaces/AAQA1FeFPy0/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=FnTOpsbGr8x3yDTzxbXNiwuEwAbph_9ik8bSg0yZrQk',

  // Sheet columns:
  // A Date
  // B Sender Name
  // C Sender Email
  // D Subject
  // E Links to uploaded files
  startColumn: 1,
  numColumns: 5,

  skipAttachmentPrefix: 'prodeo_'
};


/**
 * Main function.
 * Run manually or attach to a time-based trigger.
 */
function processEnrRecordsEmails() {
  const config = ENR_RECORDS_CONFIG;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(config.destinationSheetName);
  if (!sheet) {
    throw new Error(`Sheet not found: ${config.destinationSheetName}`);
  }

  const folder = DriveApp.getFolderById(config.destinationFolderId);

  const sourceLabel = GmailApp.getUserLabelByName(config.gmailLabelName);
  if (!sourceLabel) {
    throw new Error(`Gmail label not found: ${config.gmailLabelName}`);
  }

  let processedLabel = GmailApp.getUserLabelByName(config.processedLabelName);
  if (!processedLabel) {
    processedLabel = GmailApp.createLabel(config.processedLabelName);
  }

  const query = `in:inbox label:${config.gmailLabelName} -label:${config.processedLabelName}`;
  const threads = GmailApp.search(query);

  Logger.log(`Found ${threads.length} ENR_RECORDS thread(s) to review.`);

  threads.forEach(thread => {
    const messages = thread.getMessages();

    messages.forEach(message => {
      const attachments = message.getAttachments({
        includeInlineImages: false,
        includeAttachments: true
      });

      const uploadedFiles = [];

      attachments.forEach(attachment => {
        const originalFilename = attachment.getName();

        if (shouldSkipAttachment_(originalFilename, config.skipAttachmentPrefix)) {
          Logger.log(`Skipping request attachment: ${originalFilename}`);
          return;
        }

        const messageDate = message.getDate();
        const renamedFilename = buildDatedFilename_(messageDate, originalFilename);

        const uploadedFile = folder.createFile(attachment.copyBlob());
        uploadedFile.setName(renamedFilename);

        uploadedFiles.push({
          name: renamedFilename,
          url: uploadedFile.getUrl()
        });

        Logger.log(`Uploaded file: ${renamedFilename}`);
      });

      // Only log and alert emails where at least one qualifying attachment was uploaded.
      if (uploadedFiles.length > 0) {
        const sender = parseSender_(message.getFrom());
        const subject = message.getSubject();
        const body = getCleanEmailBody_(message);

        const row = findFirstEmptyRowInColumnA_(sheet);

        const rowValues = [[
          message.getDate(),
          sender.name,
          sender.email,
          subject,
          uploadedFiles.map(file => file.name).join('\n')
        ]];

        sheet
          .getRange(row, config.startColumn, 1, config.numColumns)
          .setValues(rowValues);

        sheet
          .getRange(row, 1)
          .setNumberFormat('MM/dd/yyyy');

        // Add the email body as a note on the subject cell in column D.
        sheet
          .getRange(row, 4)
          .setNote(body);

        setLinksCell_(sheet.getRange(row, 5), uploadedFiles);

        sendRecordsReceivedChatAlert_({
          webhookUrl: config.chatWebhookUrl,
          senderName: sender.name,
          senderEmail: sender.email,
          subject: subject,
          body: body,
          files: uploadedFiles
        });

        Logger.log(`Logged row ${row} for email: ${subject}`);
      }
    });

    thread.addLabel(processedLabel);
  });

  Logger.log('ENR_RECORDS processing complete.');
}


/**
 * Skips attachments that start with prodeo_, ignoring case.
 */
function shouldSkipAttachment_(filename, prefix) {
  if (!filename) return true;

  return filename
    .toLowerCase()
    .startsWith(prefix.toLowerCase());
}


/**
 * Renames files as YYYYMMDD_filename.extension.
 * Example:
 * enrollment.pdf -> 20260617_enrollment.pdf
 */
function buildDatedFilename_(date, originalFilename) {
  const datePrefix = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyyMMdd'
  );

  return `${datePrefix}_${originalFilename}`;
}


/**
 * Parses Gmail sender strings like:
 * "Jane Doe <jane@example.com>"
 * or:
 * jane@example.com
 */
function parseSender_(fromText) {
  const match = fromText.match(/^(.*?)\s*<([^>]+)>$/);

  if (match) {
    return {
      name: cleanSenderName_(match[1]),
      email: match[2].trim()
    };
  }

  return {
    name: '',
    email: fromText.trim()
  };
}


/**
 * Removes surrounding quotes from sender names.
 */
function cleanSenderName_(name) {
  return name
    .trim()
    .replace(/^"|"$/g, '');
}


/**
 * Gets only the most recent message from the email body.
 * Attempts to remove quoted replies / forwarded thread history.
 */
function getCleanEmailBody_(message) {
  const maxBodyLength = 12000;

  let body = message.getPlainBody() || '';

  body = body
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .trim();

  body = keepMostRecentEmailMessageOnly_(body);

  body = body
    .replace(/\n{4,}/g, '\n\n\n')
    .trim();

  if (body.length > maxBodyLength) {
    body = body.substring(0, maxBodyLength) + '\n\n[Message body truncated]';
  }

  return body;
}


/**
 * Gets only the most recent visible message from the email body.
 * Preserves hyperlink display text, but removes the underlying URLs.
 */
function getCleanEmailBody_(message) {
  const maxBodyLength = 12000;

  let body = message.getBody() || '';

  body = htmlEmailToPlainTextWithoutHrefUrls_(body);

  body = body
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .trim();

  body = keepMostRecentEmailMessageOnly_(body);

  body = removeBareUrlsFromEmailBody_(body);

  body = body
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{4,}/g, '\n\n\n')
    .trim();

  if (body.length > maxBodyLength) {
    body = body.substring(0, maxBodyLength) + '\n\n[Message body truncated]';
  }

  return body;
}


/**
 * Converts HTML email to plain text while keeping the visible text of links
 * and discarding the actual href URLs.
 *
 * Example:
 * <a href="https://long-tracking-url">4141 University Ave NE</a>
 * becomes:
 * 4141 University Ave NE
 */
function htmlEmailToPlainTextWithoutHrefUrls_(html) {
  if (!html) return '';

  let text = html;

  // Remove script/style blocks.
  text = text.replace(/<script[\s\S]*?<\/script>/gi, '');
  text = text.replace(/<style[\s\S]*?<\/style>/gi, '');

  // Keep only the display text inside links.
  text = text.replace(/<a\b[^>]*>([\s\S]*?)<\/a>/gi, '$1');

  // Convert common block/line-break tags to newlines.
  text = text.replace(/<br\s*\/?>/gi, '\n');
  text = text.replace(/<\/p>/gi, '\n');
  text = text.replace(/<\/div>/gi, '\n');
  text = text.replace(/<\/tr>/gi, '\n');
  text = text.replace(/<\/li>/gi, '\n');

  // Add spacing for table cells.
  text = text.replace(/<\/td>/gi, ' ');

  // Remove remaining HTML tags.
  text = text.replace(/<[^>]+>/g, '');

  // Decode common HTML entities.
  text = decodeHtmlEntities_(text);

  return text;
}


/**
 * Removes remaining bare URLs that may survive from plain-text signatures
 * or auto-generated email footers.
 *
 * This intentionally removes http/https URLs only.
 * It does NOT remove things like www.prodeoacademy.org, because that may be
 * the visible display text the sender intended.
 */
function removeBareUrlsFromEmailBody_(body) {
  if (!body) return '';

  return body
    .split('\n')
    .map(line => {
      return line
        .replace(/https?:\/\/\S+/gi, '')
        .replace(/[ \t]{2,}/g, ' ')
        .trimEnd();
    })
    .filter(line => line.trim() !== '')
    .join('\n');
}


/**
 * Decodes common HTML entities.
 */
function decodeHtmlEntities_(text) {
  if (!text) return '';

  return text
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&apos;/gi, "'");
}

/**
 * Finds the first empty row in column A.
 * Preserves headers and existing records.
 */
function findFirstEmptyRowInColumnA_(sheet) {
  const maxRows = sheet.getMaxRows();
  const values = sheet.getRange(1, 1, maxRows, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) {
      return i + 1;
    }
  }

  return maxRows + 1;
}


/**
 * Writes one or more clickable Drive links into the links cell.
 * Multiple uploaded files are separated by line breaks.
 */
function setLinksCell_(cell, files) {
  if (!files || files.length === 0) {
    cell.clearContent();
    return;
  }

  const text = files.map(file => file.name).join('\n');
  const richTextBuilder = SpreadsheetApp.newRichTextValue().setText(text);

  let position = 0;

  files.forEach(file => {
    const start = position;
    const end = start + file.name.length;

    richTextBuilder.setLinkUrl(start, end, file.url);

    position = end + 1; // Account for newline.
  });

  cell.setRichTextValue(richTextBuilder.build());
  cell.setWrap(true);
}


/**
 * Sends a Google Chat alert when records are received.
 */
function sendRecordsReceivedChatAlert_(options) {
  const recordsFolderUrl = 'https://drive.google.com/drive/folders/1lCJhLvAzT4p-GyMzbd9DU6bT8BQ7Z36n';

  const filesText = options.files
    .map(file => `• <${file.url}|${file.name}>`)
    .join('\n');

  const messageText =
    `*Records Received*\n` +
    `From: ${options.senderName || '[No sender name]'} - ${options.senderEmail || '[No sender email]'}\n\n` +

    `*Files:*\n` +
    `${filesText || '[No files uploaded]'}\n\n` +

    `_All received files can be found in the <${recordsFolderUrl}|records folder>._\n\n` +

    `*Message:*\n` +
    `Subject: ${options.subject || '[No subject]'}\n\n` +
    `${options.body || '[No email body]'}`;

  const payload = {
    text: messageText
  };

  const response = UrlFetchApp.fetch(options.webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();

  if (responseCode < 200 || responseCode >= 300) {
    Logger.log(`Google Chat alert failed. Response code: ${responseCode}`);
    Logger.log(response.getContentText());
  } else {
    Logger.log('Google Chat alert sent.');
  }
}