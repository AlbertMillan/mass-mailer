/**
 * SheetMailer - Mass Email Application for Google Sheets
 * Main application logic
 */

// ==================== MENU & SIDEBAR ====================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SheetMailer')
    .addItem('Open Sidebar', 'showSidebar')
    .addItem('Send Test Email', 'showTestEmailDialog')
    .addSeparator()
    .addItem('Refresh Tracking Data', 'refreshTrackingData')
    .addItem('View Campaign Stats', 'showStatsDialog')
    .addSeparator()
    .addItem('Settings', 'showSettingsDialog')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('SheetMailer')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ==================== CONFIGURATION ====================

const CONFIG = {
  TRACKING_COLUMNS: ['Status', 'Sent At', 'Opens', 'Clicks', 'Last Opened'],
  STATUS: {
    PENDING: 'PENDING',
    SENT: 'SENT',
    OPEN: 'OPEN',
    FAILED: 'FAILED',
    SCHEDULED: 'SCHEDULED',
    INVALID: 'INVALID'
  },
  STATUS_COLORS: {
    PENDING: '#E0E0E0',   // Light Gray
    SENT: '#A8D4F0',      // Light Blue
    OPEN: '#B7E1CD',      // Light Green
    FAILED: '#F4CCCC',    // Light Red
    SCHEDULED: '#FFE599', // Light Yellow
    INVALID: '#D9D2E9'    // Light Purple
  },
  THROTTLE_DELAY: 1000, // 1 second between emails
  EMAIL_COLUMN_NAMES: ['email', 'e-mail', 'email address', 'emailaddress', 'mail'],
  MERGE_TAG_REGEX: /\{\{([^}]+)\}\}/g
};

/**
 * Set status value and background color for a cell
 */
function setStatusWithColor(sheet, row, col, status) {
  const cell = sheet.getRange(row, col);
  cell.setValue(status);
  const color = CONFIG.STATUS_COLORS[status];
  if (color) {
    cell.setBackground(color);
  }
}

// ==================== DATA RETRIEVAL ====================

/**
 * Get all Gmail drafts for template selection
 */
function getGmailDrafts() {
  const drafts = GmailApp.getDrafts();
  return drafts.map(draft => {
    const message = draft.getMessage();
    return {
      id: draft.getId(),
      subject: message.getSubject() || '(No Subject)',
      snippet: message.getPlainBody().substring(0, 100) + '...'
    };
  });
}

/**
 * Get Gmail aliases (send-as addresses)
 */
function getGmailAliases() {
  const aliases = [];

  // Primary email
  const primaryEmail = Session.getActiveUser().getEmail();
  aliases.push({
    email: primaryEmail,
    name: 'Primary',
    isPrimary: true
  });

  // Try to get aliases via Gmail API (requires advanced service)
  try {
    if (typeof Gmail !== 'undefined') {
      const sendAs = Gmail.Users.Settings.SendAs.list('me');
      if (sendAs.sendAs) {
        sendAs.sendAs.forEach(alias => {
          if (alias.sendAsEmail !== primaryEmail && alias.verificationStatus === 'accepted') {
            aliases.push({
              email: alias.sendAsEmail,
              name: alias.displayName || alias.sendAsEmail,
              isPrimary: false
            });
          }
        });
      }
    }
  } catch (e) {
    console.log('Gmail API not available for aliases: ' + e.message);
  }

  return aliases;
}

/**
 * Get sheet headers and detect email column
 */
function getSheetInfo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Find email column
  let emailColumnIndex = -1;
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i].toString().toLowerCase().trim();
    if (CONFIG.EMAIL_COLUMN_NAMES.includes(header)) {
      emailColumnIndex = i;
      break;
    }
  }

  // Count rows with data
  const dataRowCount = Math.max(0, sheet.getLastRow() - 1);

  return {
    headers: headers,
    emailColumnIndex: emailColumnIndex,
    emailColumnName: emailColumnIndex >= 0 ? headers[emailColumnIndex] : null,
    rowCount: dataRowCount,
    sheetName: sheet.getName()
  };
}

/**
 * Get remaining Gmail quota
 */
function getRemainingQuota() {
  return MailApp.getRemainingDailyQuota();
}

/**
 * Get saved settings for current sheet
 */
function getSavedSettings() {
  const docProps = PropertiesService.getDocumentProperties();
  const settings = docProps.getProperty('sheetmailer_settings');
  return settings ? JSON.parse(settings) : null;
}

/**
 * Save settings for current sheet
 */
function saveSettings(settings) {
  const docProps = PropertiesService.getDocumentProperties();
  docProps.setProperty('sheetmailer_settings', JSON.stringify(settings));
}

// ==================== VALIDATION ====================

/**
 * Validate email format
 */
function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email.trim());
}

/**
 * Pre-send validation: check for duplicates, invalid emails, etc.
 */
function validateRecipients(options) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const emailColIndex = options.emailColumnIndex;
  const filterColIndex = options.filterColumnIndex;
  const filterValue = options.filterValue;

  const results = {
    valid: [],
    invalid: [],
    duplicates: [],
    filtered: 0,
    total: 0
  };

  const seenEmails = new Set();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const email = row[emailColIndex] ? row[emailColIndex].toString().trim() : '';

    // Check filter
    if (filterColIndex >= 0 && filterValue) {
      const cellValue = row[filterColIndex] ? row[filterColIndex].toString().trim().toLowerCase() : '';
      if (cellValue !== filterValue.toLowerCase()) {
        results.filtered++;
        continue;
      }
    }

    results.total++;

    // Validate email
    if (!isValidEmail(email)) {
      results.invalid.push({ row: i + 1, email: email || '(empty)' });
      continue;
    }

    // Check duplicates
    const emailLower = email.toLowerCase();
    if (seenEmails.has(emailLower)) {
      results.duplicates.push({ row: i + 1, email: email });
      continue;
    }

    seenEmails.add(emailLower);
    results.valid.push({ row: i + 1, email: email, rowData: row });
  }

  return results;
}

// ==================== EMAIL SENDING ====================

/**
 * Get draft content by ID
 */
function getDraftContent(draftId) {
  const draft = GmailApp.getDraft(draftId);
  const message = draft.getMessage();

  return {
    subject: message.getSubject(),
    htmlBody: message.getBody(),
    plainBody: message.getPlainBody(),
    attachments: message.getAttachments()
  };
}

/**
 * Replace merge tags in text
 */
function replaceMergeTags(text, rowData, headers) {
  if (!text) return text;

  return text.replace(CONFIG.MERGE_TAG_REGEX, (match, tagName) => {
    const columnIndex = headers.findIndex(h =>
      h.toString().toLowerCase().trim() === tagName.toLowerCase().trim()
    );
    if (columnIndex >= 0 && rowData[columnIndex] !== undefined) {
      return rowData[columnIndex].toString();
    }
    return match; // Keep original if not found
  });
}

/**
 * Add tracking pixel to HTML body
 */
function addTrackingPixel(htmlBody, trackingId) {
  const webAppUrl = getTrackingWebAppUrl();
  if (!webAppUrl) return htmlBody;

  const pixelUrl = `${webAppUrl}?action=open&id=${trackingId}&t=${Date.now()}`;
  const pixel = `<img src="${pixelUrl}" width="1" height="1" style="display:none" alt="" />`;

  // Add before closing body tag or at end
  if (htmlBody.includes('</body>')) {
    return htmlBody.replace('</body>', pixel + '</body>');
  }
  return htmlBody + pixel;
}

/**
 * Replace links with tracking redirects
 */
function addClickTracking(htmlBody, trackingId) {
  const webAppUrl = getTrackingWebAppUrl();
  if (!webAppUrl) return htmlBody;

  // Match href="..." in anchor tags
  const linkRegex = /<a\s+([^>]*?)href=["']([^"']+)["']([^>]*)>/gi;

  return htmlBody.replace(linkRegex, (match, before, url, after) => {
    // Don't track mailto: or tel: links
    if (url.startsWith('mailto:') || url.startsWith('tel:')) {
      return match;
    }

    const encodedUrl = encodeURIComponent(url);
    const trackingUrl = `${webAppUrl}?action=click&id=${trackingId}&url=${encodedUrl}&t=${Date.now()}`;
    return `<a ${before}href="${trackingUrl}"${after}>`;
  });
}

/**
 * Get tracking web app URL from document properties
 */
function getTrackingWebAppUrl() {
  const docProps = PropertiesService.getDocumentProperties();
  return docProps.getProperty('tracking_web_app_url');
}

/**
 * Set tracking web app URL
 */
function setTrackingWebAppUrl(url) {
  const docProps = PropertiesService.getDocumentProperties();
  docProps.setProperty('tracking_web_app_url', url);
}

/**
 * Generate unique tracking ID
 */
function generateTrackingId() {
  return Utilities.getUuid();
}

/**
 * Get attachment from Google Drive
 */
function getAttachmentFromDrive(fileIdOrUrl, rowData, headers) {
  try {
    let fileId = fileIdOrUrl;

    // Extract file ID from URL if needed
    if (fileIdOrUrl.includes('drive.google.com')) {
      const match = fileIdOrUrl.match(/[-\w]{25,}/);
      if (match) fileId = match[0];
    }

    // Replace merge tags in file ID
    fileId = replaceMergeTags(fileId, rowData, headers);

    const file = DriveApp.getFileById(fileId);
    return file.getBlob();
  } catch (e) {
    console.error('Failed to get attachment: ' + e.message);
    return null;
  }
}

/**
 * Send a single email
 */
function sendSingleEmail(options) {
  const { to, subject, htmlBody, plainBody, fromAlias, attachments, trackingId } = options;

  let finalHtmlBody = htmlBody;

  // Add tracking if enabled
  if (trackingId) {
    finalHtmlBody = addTrackingPixel(finalHtmlBody, trackingId);
    finalHtmlBody = addClickTracking(finalHtmlBody, trackingId);
  }

  const emailOptions = {
    htmlBody: finalHtmlBody,
    name: fromAlias ? fromAlias.name : undefined
  };

  // Add attachments if any
  if (attachments && attachments.length > 0) {
    emailOptions.attachments = attachments;
  }

  // Send from alias if specified
  if (fromAlias && !fromAlias.isPrimary) {
    emailOptions.from = fromAlias.email;
  }

  GmailApp.sendEmail(to, subject, plainBody, emailOptions);
}

/**
 * Main send campaign function
 */
function sendCampaign(options) {
  const {
    draftId,
    aliasEmail,
    filterColumnIndex,
    filterValue,
    attachmentColumnIndex,
    isTest,
    testEmail
  } = options;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetInfo = getSheetInfo();
  const headers = sheetInfo.headers;

  // Get draft content
  const draft = getDraftContent(draftId);

  // Get alias info
  const aliases = getGmailAliases();
  const selectedAlias = aliases.find(a => a.email === aliasEmail) || aliases[0];

  // Ensure tracking columns exist
  ensureTrackingColumns(sheet, headers);

  // Get updated headers after adding tracking columns
  const updatedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusColIndex = updatedHeaders.findIndex(h => h === 'Status');
  const sentAtColIndex = updatedHeaders.findIndex(h => h === 'Sent At');
  const opensColIndex = updatedHeaders.findIndex(h => h === 'Opens');
  const clicksColIndex = updatedHeaders.findIndex(h => h === 'Clicks');
  const trackingIdColIndex = ensureTrackingIdColumn(sheet, updatedHeaders);

  // Validate recipients
  const validation = validateRecipients({
    emailColumnIndex: sheetInfo.emailColumnIndex,
    filterColumnIndex: filterColumnIndex,
    filterValue: filterValue
  });

  // For test email, just send one
  if (isTest && testEmail) {
    const testRowData = validation.valid.length > 0 ? validation.valid[0].rowData : headers.map(() => 'Test');
    const subject = replaceMergeTags(draft.subject, testRowData, headers);
    const htmlBody = replaceMergeTags(draft.htmlBody, testRowData, headers);
    const plainBody = replaceMergeTags(draft.plainBody, testRowData, headers);

    sendSingleEmail({
      to: testEmail,
      subject: subject,
      htmlBody: htmlBody,
      plainBody: plainBody,
      fromAlias: selectedAlias,
      attachments: draft.attachments,
      trackingId: null // No tracking for test
    });

    return { success: true, message: 'Test email sent to ' + testEmail };
  }

  // Check quota
  const quota = getRemainingQuota();
  const recipientCount = validation.valid.length;

  if (recipientCount === 0) {
    return {
      success: false,
      message: 'No valid recipients found',
      validation: validation
    };
  }

  // Send emails
  const results = {
    sent: 0,
    failed: 0,
    scheduled: 0,
    errors: []
  };

  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const sheetId = sheet.getSheetId();

  for (let i = 0; i < validation.valid.length; i++) {
    const recipient = validation.valid[i];
    const rowIndex = recipient.row;
    const rowData = recipient.rowData;

    // Check if we've exceeded quota
    if (results.sent >= quota) {
      // Schedule remaining for later
      scheduleRemainingEmails({
        spreadsheetId: spreadsheetId,
        sheetId: sheetId,
        startRow: rowIndex,
        draftId: draftId,
        aliasEmail: aliasEmail,
        filterColumnIndex: filterColumnIndex,
        filterValue: filterValue,
        attachmentColumnIndex: attachmentColumnIndex
      });
      results.scheduled = validation.valid.length - i;
      break;
    }

    try {
      // Generate tracking ID
      const trackingId = generateTrackingId();

      // Replace merge tags
      const subject = replaceMergeTags(draft.subject, rowData, headers);
      const htmlBody = replaceMergeTags(draft.htmlBody, rowData, headers);
      const plainBody = replaceMergeTags(draft.plainBody, rowData, headers);

      // Get personalized attachment if specified
      let attachments = [...draft.attachments];
      if (attachmentColumnIndex >= 0 && rowData[attachmentColumnIndex]) {
        const driveAttachment = getAttachmentFromDrive(
          rowData[attachmentColumnIndex].toString(),
          rowData,
          headers
        );
        if (driveAttachment) {
          attachments.push(driveAttachment);
        }
      }

      // Send email
      sendSingleEmail({
        to: recipient.email,
        subject: subject,
        htmlBody: htmlBody,
        plainBody: plainBody,
        fromAlias: selectedAlias,
        attachments: attachments,
        trackingId: trackingId
      });

      // Update sheet
      setStatusWithColor(sheet, rowIndex, statusColIndex + 1, CONFIG.STATUS.SENT);
      sheet.getRange(rowIndex, sentAtColIndex + 1).setValue(new Date());
      sheet.getRange(rowIndex, opensColIndex + 1).setValue(0);
      sheet.getRange(rowIndex, clicksColIndex + 1).setValue(0);
      sheet.getRange(rowIndex, trackingIdColIndex + 1).setValue(trackingId);

      // Store tracking data
      storeTrackingData(trackingId, {
        spreadsheetId: spreadsheetId,
        sheetId: sheetId,
        rowIndex: rowIndex,
        email: recipient.email,
        sentAt: new Date().toISOString()
      });

      results.sent++;

      // Throttle
      if (i < validation.valid.length - 1) {
        Utilities.sleep(CONFIG.THROTTLE_DELAY);
      }

    } catch (e) {
      results.failed++;
      results.errors.push({ row: rowIndex, email: recipient.email, error: e.message });
      setStatusWithColor(sheet, rowIndex, statusColIndex + 1, CONFIG.STATUS.FAILED);
    }
  }

  // Mark invalid rows
  validation.invalid.forEach(inv => {
    setStatusWithColor(sheet, inv.row, statusColIndex + 1, CONFIG.STATUS.INVALID);
  });

  return {
    success: true,
    results: results,
    validation: validation
  };
}

/**
 * Ensure tracking columns exist in sheet
 */
function ensureTrackingColumns(sheet, headers) {
  const existingHeaders = headers.map(h => h.toString().toLowerCase());

  CONFIG.TRACKING_COLUMNS.forEach(colName => {
    if (!existingHeaders.includes(colName.toLowerCase())) {
      const newColIndex = sheet.getLastColumn() + 1;
      sheet.getRange(1, newColIndex).setValue(colName);
    }
  });
}

/**
 * Ensure tracking ID column exists (hidden)
 */
function ensureTrackingIdColumn(sheet, headers) {
  const trackingIdColName = '_TrackingId';
  let colIndex = headers.findIndex(h => h === trackingIdColName);

  if (colIndex < 0) {
    colIndex = sheet.getLastColumn();
    sheet.getRange(1, colIndex + 1).setValue(trackingIdColName);
  }

  return colIndex;
}

// ==================== TRACKING DATA STORAGE ====================

/**
 * Store tracking data in script properties (accessible from web app context)
 */
function storeTrackingData(trackingId, data) {
  const scriptProps = PropertiesService.getScriptProperties();
  const key = 'track_' + trackingId;
  scriptProps.setProperty(key, JSON.stringify(data));
}

/**
 * Get tracking data by ID
 */
function getTrackingData(trackingId) {
  const scriptProps = PropertiesService.getScriptProperties();
  const key = 'track_' + trackingId;
  const data = scriptProps.getProperty(key);
  return data ? JSON.parse(data) : null;
}

/**
 * Record open event
 */
function recordOpen(trackingId) {
  console.log('recordOpen called with trackingId: ' + trackingId);

  const data = getTrackingData(trackingId);
  console.log('Tracking data: ' + JSON.stringify(data));

  if (!data) {
    console.log('No tracking data found for trackingId: ' + trackingId);
    return;
  }

  try {
    const ss = SpreadsheetApp.openById(data.spreadsheetId);
    const sheets = ss.getSheets();
    const sheet = sheets.find(s => s.getSheetId() === data.sheetId);

    if (sheet) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const statusColIndex = headers.findIndex(h => h === 'Status');
      const opensColIndex = headers.findIndex(h => h === 'Opens');
      const lastOpenedColIndex = headers.findIndex(h => h === 'Last Opened');

      if (opensColIndex >= 0) {
        const currentOpens = sheet.getRange(data.rowIndex, opensColIndex + 1).getValue() || 0;
        sheet.getRange(data.rowIndex, opensColIndex + 1).setValue(currentOpens + 1);
      }

      if (lastOpenedColIndex >= 0) {
        sheet.getRange(data.rowIndex, lastOpenedColIndex + 1).setValue(new Date());
      }

      // Update status to OPEN on first open
      if (statusColIndex >= 0) {
        const currentStatus = sheet.getRange(data.rowIndex, statusColIndex + 1).getValue();
        if (currentStatus === CONFIG.STATUS.SENT) {
          setStatusWithColor(sheet, data.rowIndex, statusColIndex + 1, CONFIG.STATUS.OPEN);
        }
      }
    }
  } catch (e) {
    console.error('Failed to record open: ' + e.message);
  }
}

/**
 * Record click event
 */
function recordClick(trackingId) {
  const data = getTrackingData(trackingId);
  if (!data) return;

  try {
    const ss = SpreadsheetApp.openById(data.spreadsheetId);
    const sheets = ss.getSheets();
    const sheet = sheets.find(s => s.getSheetId() === data.sheetId);

    if (sheet) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const clicksColIndex = headers.findIndex(h => h === 'Clicks');

      if (clicksColIndex >= 0) {
        const currentClicks = sheet.getRange(data.rowIndex, clicksColIndex + 1).getValue() || 0;
        sheet.getRange(data.rowIndex, clicksColIndex + 1).setValue(currentClicks + 1);
      }
    }
  } catch (e) {
    console.error('Failed to record click: ' + e.message);
  }
}

// ==================== SCHEDULING ====================

/**
 * Schedule remaining emails for later
 */
function scheduleRemainingEmails(params) {
  const docProps = PropertiesService.getDocumentProperties();
  docProps.setProperty('scheduled_campaign', JSON.stringify(params));

  // Create trigger for next day
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(9, 0, 0, 0); // 9 AM next day

  ScriptApp.newTrigger('resumeScheduledCampaign')
    .timeBased()
    .at(tomorrow)
    .create();
}

/**
 * Resume a scheduled campaign
 */
function resumeScheduledCampaign() {
  const docProps = PropertiesService.getDocumentProperties();
  const params = docProps.getProperty('scheduled_campaign');

  if (!params) return;

  const campaignParams = JSON.parse(params);
  docProps.deleteProperty('scheduled_campaign');

  // Delete the trigger
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'resumeScheduledCampaign') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Resume sending
  sendCampaign(campaignParams);
}

/**
 * Schedule a campaign for later
 */
function scheduleCampaign(options, scheduledTime) {
  const docProps = PropertiesService.getDocumentProperties();
  docProps.setProperty('scheduled_campaign', JSON.stringify(options));

  const triggerTime = new Date(scheduledTime);

  ScriptApp.newTrigger('executeScheduledCampaign')
    .timeBased()
    .at(triggerTime)
    .create();

  return { success: true, scheduledFor: triggerTime.toISOString() };
}

/**
 * Execute scheduled campaign
 */
function executeScheduledCampaign() {
  const docProps = PropertiesService.getDocumentProperties();
  const params = docProps.getProperty('scheduled_campaign');

  if (!params) return;

  const campaignParams = JSON.parse(params);
  docProps.deleteProperty('scheduled_campaign');

  // Delete the trigger
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'executeScheduledCampaign') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  sendCampaign(campaignParams);
}

// ==================== STATISTICS ====================

/**
 * Get campaign statistics
 */
function getCampaignStats() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const statusColIndex = headers.findIndex(h => h === 'Status');
  const opensColIndex = headers.findIndex(h => h === 'Opens');
  const clicksColIndex = headers.findIndex(h => h === 'Clicks');

  const stats = {
    total: data.length - 1,
    sent: 0,
    open: 0,
    failed: 0,
    pending: 0,
    invalid: 0,
    totalOpens: 0,
    totalClicks: 0,
    uniqueOpens: 0,
    uniqueClicks: 0
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[statusColIndex];
    const opens = row[opensColIndex] || 0;
    const clicks = row[clicksColIndex] || 0;

    switch (status) {
      case CONFIG.STATUS.SENT:
        stats.sent++;
        break;
      case CONFIG.STATUS.OPEN:
        stats.open++;
        break;
      case CONFIG.STATUS.FAILED:
        stats.failed++;
        break;
      case CONFIG.STATUS.INVALID:
        stats.invalid++;
        break;
      default:
        stats.pending++;
    }

    stats.totalOpens += opens;
    stats.totalClicks += clicks;
    if (opens > 0) stats.uniqueOpens++;
    if (clicks > 0) stats.uniqueClicks++;
  }

  // Calculate rates (sent + open = total delivered)
  const totalDelivered = stats.sent + stats.open;
  if (totalDelivered > 0) {
    stats.openRate = ((stats.uniqueOpens / totalDelivered) * 100).toFixed(1) + '%';
    stats.clickRate = ((stats.uniqueClicks / totalDelivered) * 100).toFixed(1) + '%';
  } else {
    stats.openRate = '0%';
    stats.clickRate = '0%';
  }

  return stats;
}

/**
 * Refresh tracking data manually
 */
function refreshTrackingData() {
  // This would typically fetch from an external source
  // For now, data is updated in real-time via web app
  SpreadsheetApp.getActiveSpreadsheet().toast('Tracking data is up to date', 'SheetMailer', 3);
}

// ==================== DIALOGS ====================

function showTestEmailDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      input { width: 100%; padding: 8px; margin: 10px 0; box-sizing: border-box; }
      button { padding: 10px 20px; background: #4285f4; color: white; border: none; cursor: pointer; }
    </style>
    <p>Enter email address for test:</p>
    <input type="email" id="testEmail" placeholder="test@example.com" />
    <button onclick="sendTest()">Send Test</button>
    <script>
      function sendTest() {
        const email = document.getElementById('testEmail').value;
        google.script.run.withSuccessHandler(() => {
          alert('Test email sent!');
          google.script.host.close();
        }).sendTestFromDialog(email);
      }
    </script>
  `).setWidth(400).setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(html, 'Send Test Email');
}

function showStatsDialog() {
  const stats = getCampaignStats();
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      .stat { margin: 10px 0; }
      .label { font-weight: bold; }
    </style>
    <div class="stat"><span class="label">Total Recipients:</span> ${stats.total}</div>
    <div class="stat"><span class="label">Sent:</span> ${stats.sent}</div>
    <div class="stat"><span class="label">Opened:</span> ${stats.open}</div>
    <div class="stat"><span class="label">Failed:</span> ${stats.failed}</div>
    <div class="stat"><span class="label">Open Rate:</span> ${stats.openRate} (${stats.uniqueOpens} unique opens)</div>
    <div class="stat"><span class="label">Click Rate:</span> ${stats.clickRate} (${stats.uniqueClicks} unique clicks)</div>
  `).setWidth(400).setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(html, 'Campaign Statistics');
}

function showSettingsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Settings')
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'SheetMailer Settings');
}
