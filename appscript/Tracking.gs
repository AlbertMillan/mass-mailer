/**
 * SheetMailer - Tracking Web App
 * Handles open tracking pixels and click tracking redirects
 *
 * DEPLOYMENT INSTRUCTIONS:
 * 1. Deploy > New deployment
 * 2. Select "Web app"
 * 3. Execute as: "Me"
 * 4. Who has access: "Anyone"
 * 5. Copy the web app URL and set it using setTrackingWebAppUrl(url)
 */

/**
 * Handle GET requests for tracking
 * @param {Object} e - Event object with query parameters
 */
function doGet(e) {
  console.log('doGet called with params: ' + JSON.stringify(e.parameter));

  const action = e.parameter.action;
  const trackingId = e.parameter.id;
  const targetUrl = e.parameter.url;

  if (!trackingId) {
    console.log('No trackingId provided');
    return createPixelResponse();
  }

  try {
    if (action === 'open') {
      console.log('Processing open action for: ' + trackingId);
      // Record email open
      recordOpen(trackingId);
      return createPixelResponse();
    } else if (action === 'click' && targetUrl) {
      // Record click and redirect
      const decodedUrl = decodeURIComponent(targetUrl);
      recordClick(trackingId, decodedUrl);
      return HtmlService.createHtmlOutput(
        `<script>window.location.href="${decodedUrl}";</script>`
      );
    }
  } catch (error) {
    console.error('Tracking error: ' + error.message);
  }

  // Default: return tracking pixel
  return createPixelResponse();
}

/**
 * Create a 1x1 transparent GIF response
 */
function createPixelResponse() {
  // 1x1 transparent GIF in base64
  const pixel = Utilities.base64Decode(
    'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'
  );

  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.GIF);
}

/**
 * Alternative: Return actual image bytes
 * Note: Apps Script doesn't fully support binary responses,
 * so we use a redirect to a data URI as a workaround
 */
function createImageResponse() {
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta http-equiv="refresh" content="0;url=data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7">
      </head>
    </html>
  `;
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handle POST requests (for future use)
 */
function doPost(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Test the tracking setup
 */
function testTracking() {
  const url = getTrackingWebAppUrl();
  if (!url) {
    Logger.log('Tracking URL not set. Please deploy the web app and call setTrackingWebAppUrl(url)');
    return;
  }

  Logger.log('Tracking URL: ' + url);
  Logger.log('Open tracking URL: ' + url + '?action=open&id=test123');
  Logger.log('Click tracking URL: ' + url + '?action=click&id=test123&url=' + encodeURIComponent('https://example.com'));
}

/**
 * Batch update tracking data from external source (if needed)
 */
function batchUpdateTracking(updates) {
  updates.forEach(update => {
    if (update.type === 'open') {
      recordOpen(update.trackingId);
    } else if (update.type === 'click') {
      recordClick(update.trackingId, update.url);
    }
  });
}

/**
 * Clean up old tracking data (run periodically)
 */
function cleanupOldTrackingData() {
  const scriptProps = PropertiesService.getScriptProperties();
  const allProps = scriptProps.getProperties();
  const now = new Date();
  const maxAge = 90 * 24 * 60 * 60 * 1000; // 90 days

  let deleted = 0;

  Object.keys(allProps).forEach(key => {
    if (key.startsWith('track_')) {
      try {
        const data = JSON.parse(allProps[key]);
        const sentAt = new Date(data.sentAt);
        if (now - sentAt > maxAge) {
          scriptProps.deleteProperty(key);
          deleted++;
        }
      } catch (e) {
        // Invalid data, delete it
        scriptProps.deleteProperty(key);
        deleted++;
      }
    }
  });

  Logger.log('Cleaned up ' + deleted + ' old tracking records');
  return deleted;
}

/**
 * Create a time-based trigger to clean up old data weekly
 */
function createCleanupTrigger() {
  // Delete existing cleanup triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'cleanupOldTrackingData') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new weekly trigger
  ScriptApp.newTrigger('cleanupOldTrackingData')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(3)
    .create();

  Logger.log('Cleanup trigger created for every Sunday at 3 AM');
}

/**
 * Get tracking statistics summary
 */
function getTrackingStats() {
  const scriptProps = PropertiesService.getScriptProperties();
  const allProps = scriptProps.getProperties();

  let totalTracked = 0;
  const spreadsheets = new Set();

  Object.keys(allProps).forEach(key => {
    if (key.startsWith('track_')) {
      totalTracked++;
      try {
        const data = JSON.parse(allProps[key]);
        if (data.spreadsheetId) {
          spreadsheets.add(data.spreadsheetId);
        }
      } catch (e) {
        // Ignore invalid data
      }
    }
  });

  return {
    totalTrackedEmails: totalTracked,
    uniqueSpreadsheets: spreadsheets.size,
    trackingUrl: getTrackingWebAppUrl()
  };
}

/**
 * Debug: List all tracking entries
 */
function listAllTrackingEntries() {
  const scriptProps = PropertiesService.getScriptProperties();
  const allProps = scriptProps.getProperties();

  const entries = [];

  Object.keys(allProps).forEach(key => {
    if (key.startsWith('track_')) {
      try {
        const data = JSON.parse(allProps[key]);
        entries.push({
          id: key.replace('track_', ''),
          ...data
        });
      } catch (e) {
        entries.push({
          id: key.replace('track_', ''),
          error: 'Invalid data'
        });
      }
    }
  });

  Logger.log(JSON.stringify(entries, null, 2));
  return entries;
}
