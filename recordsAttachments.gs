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

  chatWebhookPropertyName: 'ENR_RECORDS_CHAT_WEBHOOK_URL',

  // Sheet columns:
  // A Date
  // B Sender Name
  // C Sender Email
  // D Subject
  // E Links to uploaded files
  startColumn: 1,
  numColumns: 5,

  skipAttachmentPrefix: 'prodeo_',

  pdfCoApiKeyPropertyName: 'PDFCO_API_KEY',

  maxPdfSizeMb: 9,
  oversizedOriginalPrefix: 'ORIGINAL_OVERSIZED_',
  pdfSplitMaxDepth: 8,
  pdfSplitMaxFilesPerRun: 2,
};


/**
 * Main function.
 * Run manually or attach to a time-based trigger.
 */
function processEnrRecordsEmails() {
  if (!shouldRunEnrRecordsNow_()) {
    Logger.log('Skipping ENR_RECORDS processing because current time is outside weekday 7am-7pm window.');
    return;
  }

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

        const uploadedFile = folder.createFile(
          attachment.copyBlob().setName(renamedFilename)
        );

        uploadedFile.setName(renamedFilename);

        Logger.log(`Uploaded file: ${renamedFilename}`);

        const finalFiles = prepareUploadedRecordFileForAlertAndLog_({
          file: uploadedFile,
          folder: folder,
          config: config
        });

uploadedFiles.push(...finalFiles);
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
          webhookUrl: getRequiredScriptProperty_(config.chatWebhookPropertyName),
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
 * Gets a required script property.
 * Throws a clear error if it has not been set.
 */
function getRequiredScriptProperty_(propertyName) {
  const value = PropertiesService
    .getScriptProperties()
    .getProperty(propertyName);

  if (!value) {
    throw new Error(
      `Missing required script property: ${propertyName}. ` +
      `Add it in Apps Script Project Settings > Script Properties.`
    );
  }

  return value;
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
 * Trims an email body down to the newest visible message.
 *
 * Handles common Gmail / Outlook reply separators like:
 * On Tue, Jun 16, 2026 at 10:23 AM Jane Doe <jane@example.com> wrote:
 * From: Jane Doe <jane@example.com>
 * -----Original Message-----
 * ---------- Forwarded message ---------
 * ________________________________
 */
function keepMostRecentEmailMessageOnly_(body) {
  if (!body) return '';

  const splitPatterns = [
    // Gmail reply header:
    // On Tue, Jun 16, 2026 at 10:23 AM Jane Doe <jane@example.com> wrote:
    /\nOn .+?wrote:\s*\n/i,

    // Outlook / Microsoft style original message separators.
    /\n-{2,}\s*Original Message\s*-{2,}\s*\n/i,
    /\n_{5,}\s*\n/,

    // Forwarded message separators.
    /\n-{2,}\s*Forwarded message\s*-{2,}\s*\n/i,
    /\nBegin forwarded message:\s*\n/i,

    // Common reply metadata blocks.
    /\nFrom:\s.+\nSent:\s.+\nTo:\s.+\nSubject:\s.+/i,
    /\nFrom:\s.+\nDate:\s.+\nSubject:\s.+\nTo:\s.+/i,

    // Simpler fallback for messages where quoted content starts with From:
    /\nFrom:\s.+<.+@.+>\s*\n/i
  ];

  let earliestSplitIndex = -1;

  splitPatterns.forEach(pattern => {
    const match = body.match(pattern);

    if (match && typeof match.index === 'number') {
      if (earliestSplitIndex === -1 || match.index < earliestSplitIndex) {
        earliestSplitIndex = match.index;
      }
    }
  });

  if (earliestSplitIndex > -1) {
    body = body.substring(0, earliestSplitIndex);
  }

  // Remove quoted lines that survived the split.
  body = body
    .split('\n')
    .filter(line => !line.trim().startsWith('>'))
    .join('\n');

  // Remove common mobile/client signatures if desired.
  // Keep this conservative so we don't accidentally remove useful sender text.
  body = body.replace(/\nSent from my iPhone\s*$/i, '');
  body = body.replace(/\nSent from my iPad\s*$/i, '');
  body = body.replace(/\nGet Outlook for .+$/i, '');

  return body.trim();
}

/**
 * Returns true only Monday-Friday from 7:00 AM through 6:59 PM.
 */
function shouldRunEnrRecordsNow_() {
  const now = new Date();
  const timezone = Session.getScriptTimeZone();

  const dayOfWeek = Number(
    Utilities.formatDate(now, timezone, 'u')
  );
  // ISO day of week:
  // 1 = Monday
  // 2 = Tuesday
  // ...
  // 7 = Sunday

  const hour = Number(
    Utilities.formatDate(now, timezone, 'H')
  );
  // 0-23

  const isWeekday = dayOfWeek >= 1 && dayOfWeek <= 5;
  const isBusinessHours = hour >= 7 && hour < 19;

  return isWeekday && isBusinessHours;
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
    .filter(file => !file.name.startsWith(ENR_RECORDS_CONFIG.oversizedOriginalPrefix))
    .map(file => {
      const warning = file.stillOversized ? ' — STILL OVERSIZED' : '';
      return `• <${file.url}|${file.name}>${warning}`;
    })
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

/**
 * Creates one every-2-hours trigger.
 *
 * The function itself decides whether to continue or skip
 * based on weekday/business-hour rules.
 *
 * Run this manually one time.
 */
function createEnrRecordsEveryTwoHoursTrigger() {
  deleteEnrRecordsTriggers_();

  ScriptApp.newTrigger('processEnrRecordsEmails')
    .timeBased()
    .everyHours(2)
    .create();

  Logger.log('Created ENR Records every-2-hours trigger.');
}


/**
 * Deletes existing triggers for processEnrRecordsEmails.
 */
function deleteEnrRecordsTriggers_() {
  const functionName = 'processEnrRecordsEmails';
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log(`Deleted existing triggers for ${functionName}.`);
}


/**
 * Takes a newly uploaded record file and returns the files that should be
 * logged to the sheet and posted to Google Chat.
 *
 * If the file is an oversized PDF, it is split first and the original is
 * renamed ORIGINAL_OVERSIZED_...
 *
 * If the file is not oversized, it is returned as-is.
 */
function prepareUploadedRecordFileForAlertAndLog_(options) {
  const file = options.file;
  const folder = options.folder;
  const config = options.config;

  const maxBytes = mbToBytes_(config.maxPdfSizeMb || 10);

  if (!isOversizedPdfSplitCandidate_(file, config, maxBytes)) {
    return [{
      name: file.getName(),
      url: file.getUrl()
    }];
  }

  Logger.log(`New upload is oversized PDF. Splitting before logging/chat alert: ${file.getName()}`);

  const apiKey = getRequiredScriptProperty_(config.pdfCoApiKeyPropertyName);

  const splitResult = splitSingleOversizedPdfFile_(
    file,
    folder,
    apiKey,
    config,
    maxBytes
  );

  return splitResult.createdFiles
    .filter(createdFile => {
      return !createdFile.name.startsWith(config.oversizedOriginalPrefix);
    })
    .map(createdFile => {
      return {
        name: createdFile.name,
        url: createdFile.url,
        sizeMb: createdFile.sizeMb,
        stillOversized: createdFile.stillOversized === true
      };
    });
}


/***************************************
 * Oversized PDF Splitter using PDF.co
 ***************************************/

const PDFCO_BASE_URL = 'https://api.pdf.co/v1';


/**
 * Companion function for ENR records process.
 *
 * Scans the records folder for PDF files >= 10 MB,
 * splits them using PDF.co, saves parts back to Drive,
 * and renames the original file to avoid reprocessing.
 */
function splitOversizedRecordPdfsWithPdfCo() {
  const config = ENR_RECORDS_CONFIG;

  const folder = DriveApp.getFolderById(config.destinationFolderId);
  const apiKey = getRequiredScriptProperty_(config.pdfCoApiKeyPropertyName);

  const maxBytes = mbToBytes_(config.maxPdfSizeMb || 10);
  const maxFilesPerRun = config.pdfSplitMaxFilesPerRun || 2;

  const files = folder.getFiles();
  const splitResults = [];
  const failedResults = [];

  let processedCount = 0;

  while (files.hasNext()) {
    const file = files.next();

    if (processedCount >= maxFilesPerRun) {
      Logger.log(`Stopping after ${maxFilesPerRun} oversized PDF(s) this run.`);
      break;
    }

    if (!isOversizedPdfSplitCandidate_(file, config, maxBytes)) {
      continue;
    }

    try {
      Logger.log(`Splitting oversized PDF: ${file.getName()} (${bytesToMb_(file.getSize())} MB)`);

      const result = splitSingleOversizedPdfFile_(file, folder, apiKey, config, maxBytes);

      splitResults.push(result);
      processedCount++;

      Logger.log(`Finished splitting ${file.getName()}. Created ${result.createdFiles.length} part file(s).`);
    } catch (error) {
      Logger.log(`Failed to split ${file.getName()}: ${error.message}`);

      failedResults.push({
        name: file.getName(),
        url: file.getUrl(),
        error: error.message
      });
    }
  }

  if (splitResults.length > 0 || failedResults.length > 0) {
    sendPdfSplitChatAlertIfConfigured_({
      splitResults: splitResults,
      failedResults: failedResults
    });
  }

  Logger.log('Oversized PDF split check complete.');
}


/**
 * Optional convenience function.
 *
 * Add this near the end of processEnrRecordsEmails()
 * after the normal email processing is complete.
 */
function splitOversizedRecordPdfsAfterEmailProcessing_() {
  splitOversizedRecordPdfsWithPdfCo();
}


/**
 * Determines whether this Drive file should be split.
 */
function isOversizedPdfSplitCandidate_(file, config, maxBytes) {
  const name = file.getName();

  if (!name) return false;

  if (name.startsWith(config.oversizedOriginalPrefix)) return false;
  if (/_part_\d+\.pdf$/i.test(name)) return false;

  const isPdf =
    file.getMimeType() === MimeType.PDF ||
    name.toLowerCase().endsWith('.pdf');

  if (!isPdf) return false;

  return file.getSize() >= maxBytes;
}


/**
 * Splits one oversized PDF file and saves the resulting parts.
 */
function splitSingleOversizedPdfFile_(file, folder, apiKey, config, maxBytes) {
  const originalName = file.getName();
  const originalBlob = file.getBlob().setName(originalName);

  const splitParts = splitPdfBlobRecursively_({
    blob: originalBlob,
    apiKey: apiKey,
    maxBytes: maxBytes,
    maxDepth: config.pdfSplitMaxDepth || 8,
    depth: 0
  });

  const createdFiles = [];

  splitParts.forEach((part, index) => {
    const partNumber = index + 1;
    const partName = buildPdfPartFilename_(originalName, partNumber);

    const createdFile = folder.createFile(
      part.blob
        .copyBlob()
        .setName(partName)
    );

    createdFiles.push({
      name: createdFile.getName(),
      url: createdFile.getUrl(),
      sizeMb: bytesToMb_(createdFile.getSize()),
      stillOversized: part.stillOversized === true
    });

    Logger.log(`Created split PDF part: ${createdFile.getName()} (${bytesToMb_(createdFile.getSize())} MB)`);
  });

  file.setName(config.oversizedOriginalPrefix + originalName);

  return {
    originalName: originalName,
    originalUrl: file.getUrl(),
    createdFiles: createdFiles
  };
}


/**
 * Recursively splits a PDF blob in half until each part is below maxBytes.
 *
 * If a single-page PDF is still oversized, it cannot be split further by pages.
 * The script will return it with stillOversized = true.
 */
function splitPdfBlobRecursively_(options) {
  const blob = options.blob;
  const apiKey = options.apiKey;
  const maxBytes = options.maxBytes;
  const maxDepth = options.maxDepth;
  const depth = options.depth || 0;

  const blobSize = getBlobSizeBytes_(blob);

  if (blobSize < maxBytes) {
    return [{
      blob: blob,
      stillOversized: false
    }];
  }

  if (depth >= maxDepth) {
    Logger.log(`Maximum split depth reached for ${blob.getName()}.`);

    return [{
      blob: blob,
      stillOversized: true
    }];
  }

  const uploaded = pdfCoUploadBlob_(apiKey, blob);
  const info = pdfCoGetPdfInfo_(apiKey, uploaded.url);

  const pageCount = Number(info.info && info.info.PageCount);

  if (!pageCount || pageCount < 1) {
    throw new Error(`Unable to determine page count for ${blob.getName()}.`);
  }

  if (pageCount <= 1) {
    Logger.log(`Single-page PDF is still oversized and cannot be split further by page: ${blob.getName()}`);

    return [{
      blob: blob,
      stillOversized: true
    }];
  }

  const firstEndPage = Math.floor(pageCount / 2);
  const pages = [
    buildPdfPageRange_(1, firstEndPage, pageCount),
    buildPdfPageRange_(firstEndPage + 1, pageCount, pageCount)
  ].join(',');

  Logger.log(`Splitting ${blob.getName()} at depth ${depth}. Pages: ${pages}`);

  const splitResponse = pdfCoSplitPdf_(apiKey, uploaded.url, pages, blob.getName());

  if (!splitResponse.urls || splitResponse.urls.length < 2) {
    throw new Error(`PDF.co did not return expected split URLs for ${blob.getName()}.`);
  }

  let results = [];

  splitResponse.urls.forEach((url, index) => {
    const downloadedBlob = downloadPdfBlobFromUrl_(
      url,
      `${stripPdfExtension_(blob.getName())}_temp_${index + 1}.pdf`
    );

    const downloadedSize = getBlobSizeBytes_(downloadedBlob);

    if (downloadedSize >= maxBytes) {
      const nestedResults = splitPdfBlobRecursively_({
        blob: downloadedBlob,
        apiKey: apiKey,
        maxBytes: maxBytes,
        maxDepth: maxDepth,
        depth: depth + 1
      });

      results = results.concat(nestedResults);
    } else {
      results.push({
        blob: downloadedBlob,
        stillOversized: false
      });
    }
  });

  return results;
}


/**
 * Uploads a Drive blob to PDF.co temporary storage.
 *
 * PDF.co's presigned upload flow returns:
 * - presignedUrl: where the script PUTs the local file
 * - url: the temporary file URL to pass into later PDF.co endpoints
 */
function pdfCoUploadBlob_(apiKey, blob) {
  const filename = blob.getName() || 'source.pdf';

  const presignedEndpoint =
    `${PDFCO_BASE_URL}/file/upload/get-presigned-url` +
    `?contenttype=application/pdf` +
    `&name=${encodeURIComponent(filename)}`;

  const presignedResponse = UrlFetchApp.fetch(presignedEndpoint, {
    method: 'get',
    headers: {
      'x-api-key': apiKey
    },
    muteHttpExceptions: true
  });

  const presignedJson = parsePdfCoJsonResponse_(presignedResponse, 'PDF.co presigned upload URL');

  const uploadResponse = UrlFetchApp.fetch(presignedJson.presignedUrl, {
    method: 'put',
    contentType: 'application/pdf',
    payload: blob.getBytes(),
    muteHttpExceptions: true
  });

  const uploadCode = uploadResponse.getResponseCode();

  if (uploadCode < 200 || uploadCode >= 300) {
    throw new Error(
      `PDF.co file upload failed. HTTP ${uploadCode}: ${uploadResponse.getContentText()}`
    );
  }

  return {
    url: presignedJson.url,
    presignedUrl: presignedJson.presignedUrl
  };
}


/**
 * Gets PDF metadata from PDF.co, including PageCount.
 */
function pdfCoGetPdfInfo_(apiKey, uploadedFileUrl) {
  return pdfCoPostJson_(
    apiKey,
    '/pdf/info',
    {
      url: uploadedFileUrl,
      async: false
    },
    'PDF.co PDF Info'
  );
}


/**
 * Splits a PDF by page ranges using PDF.co.
 *
 * The pages parameter is 1-based.
 * Example: "1-5,6-" returns two files.
 */
function pdfCoSplitPdf_(apiKey, uploadedFileUrl, pages, originalName) {
  return pdfCoPostJson_(
    apiKey,
    '/pdf/split',
    {
      url: uploadedFileUrl,
      pages: pages,
      inline: false,
      async: false,
      name: stripPdfExtension_(originalName) + '_split.pdf'
    },
    'PDF.co Split PDF'
  );
}


/**
 * Generic PDF.co JSON POST helper.
 */
function pdfCoPostJson_(apiKey, path, payload, operationName) {
  const response = UrlFetchApp.fetch(PDFCO_BASE_URL + path, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  return parsePdfCoJsonResponse_(response, operationName);
}


/**
 * Parses and validates PDF.co JSON responses.
 */
function parsePdfCoJsonResponse_(response, operationName) {
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  let json;

  try {
    json = JSON.parse(responseText);
  } catch (error) {
    throw new Error(`${operationName} returned non-JSON response. HTTP ${responseCode}: ${responseText}`);
  }

  if (responseCode < 200 || responseCode >= 300) {
    throw new Error(`${operationName} failed. HTTP ${responseCode}: ${responseText}`);
  }

  if (json.error === true) {
    throw new Error(`${operationName} failed: ${json.message || responseText}`);
  }

  return json;
}


/**
 * Downloads a split PDF part from a PDF.co temporary URL.
 */
function downloadPdfBlobFromUrl_(url, filename) {
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();

  if (responseCode < 200 || responseCode >= 300) {
    throw new Error(`Failed to download split PDF part. HTTP ${responseCode}: ${response.getContentText()}`);
  }

  return response
    .getBlob()
    .setContentType(MimeType.PDF)
    .setName(filename);
}


/**
 * Builds filenames like:
 * original_part_1.pdf
 * original_part_2.pdf
 */
function buildPdfPartFilename_(originalName, partNumber) {
  const baseName = stripPdfExtension_(originalName);

  return `${baseName}_part_${partNumber}.pdf`;
}


/**
 * Builds PDF.co 1-based page ranges.
 */
function buildPdfPageRange_(startPage, endPage, totalPages) {
  if (startPage === endPage) {
    return String(startPage);
  }

  if (endPage === totalPages) {
    return `${startPage}-`;
  }

  return `${startPage}-${endPage}`;
}


/**
 * Removes .pdf from the end of a filename.
 */
function stripPdfExtension_(filename) {
  return String(filename || 'document').replace(/\.pdf$/i, '');
}


/**
 * Returns blob size in bytes.
 */
function getBlobSizeBytes_(blob) {
  return blob.getBytes().length;
}


/**
 * Converts MB to bytes.
 */
function mbToBytes_(mb) {
  return mb * 1024 * 1024;
}


/**
 * Converts bytes to MB, rounded to one decimal.
 */
function bytesToMb_(bytes) {
  return Math.round((bytes / 1024 / 1024) * 10) / 10;
}


/**
 * Sends a Google Chat alert after oversized PDFs are split.
 * Uses your existing ENR_RECORDS_CHAT_WEBHOOK_URL script property.
 */
function sendPdfSplitChatAlertIfConfigured_(options) {
  const config = ENR_RECORDS_CONFIG;

  let webhookUrl;

  try {
    webhookUrl = getRequiredScriptProperty_(config.chatWebhookPropertyName);
  } catch (error) {
    Logger.log(`Skipping PDF split chat alert: ${error.message}`);
    return;
  }

  const recordsFolderUrl = `https://drive.google.com/drive/folders/${config.destinationFolderId}`;

  const successfulText = options.splitResults.map(result => {
    const createdFilesText = result.createdFiles.map(file => {
      const sizeText = `${file.sizeMb} MB`;
      const warning = file.stillOversized ? ' — STILL OVERSIZED' : '';

      return `• <${file.url}|${file.name}> — ${sizeText}${warning}`;
    }).join('\n');

    return `*Original:* <${result.originalUrl}|${result.originalName}>\n${createdFilesText}`;
  }).join('\n\n');

  const failedText = options.failedResults.map(file => {
    return `• <${file.url}|${file.name}> — ${file.error}`;
  }).join('\n');

  let messageText =
    `*Oversized PDFs Split*\n\n`;

  if (successfulText) {
    messageText +=
      `*Created Files:*\n` +
      `${successfulText}\n\n`;
  }

  if (failedText) {
    messageText +=
      `*Failed Files:*\n` +
      `${failedText}\n\n`;
  }

  messageText +=
    `_All received files can be found in the <${recordsFolderUrl}|records folder>._`;

  const response = UrlFetchApp.fetch(webhookUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      text: messageText
    }),
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();

  if (responseCode < 200 || responseCode >= 300) {
    Logger.log(`Google Chat PDF split alert failed. Response code: ${responseCode}`);
    Logger.log(response.getContentText());
  } else {
    Logger.log('Google Chat PDF split alert sent.');
  }
}