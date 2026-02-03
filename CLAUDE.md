# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

SheetMailer is a Google Apps Script-based mass email application bound to Google Sheets. It provides mail merge functionality with tracking (opens, clicks), Gmail alias support, and scheduling.

## Development Commands

```bash
# Push local changes to Google Apps Script
cd appscript && clasp push

# Force push (overwrites remote)
cd appscript && clasp push --force

# Pull remote changes
cd appscript && clasp pull

# Watch mode (auto-push on save)
cd appscript && clasp push --watch

# View execution logs
cd appscript && clasp logs --watch

# Open in browser
cd appscript && clasp open

# Deploy web app (creates new version)
cd appscript && clasp deploy

# Update existing deployment
cd appscript && clasp deploy --deploymentId <ID>
```

## Architecture

### File Structure
- `appscript/` - All Google Apps Script code (managed by clasp)
  - `Code.gs` - Main application: menu, sidebar, email sending, validation, merge tags
  - `Tracking.gs` - Web app (`doGet`) for open/click tracking via pixel and redirects
  - `Sidebar.html` - Main UI sidebar
  - `Settings.html` - Settings dialog
  - `appsscript.json` - Manifest with OAuth scopes and advanced services
  - `.clasp.json` - clasp project config (contains Script ID)

### Key Patterns

**Tracking System**: Uses ScriptProperties to store tracking data keyed by UUID (`track_<uuid>`). The web app deployment handles GET requests for open pixels and click redirects, updating the bound spreadsheet.

**Status Flow**: PENDING → SENT → OPEN (on first open). FAILED/INVALID for errors.

**Merge Tags**: `{{ColumnName}}` syntax replaced via regex matching against sheet headers.

**Storage**:
- `ScriptProperties` - Tracking data (cross-request persistence for web app)
- `DocumentProperties` - Settings and scheduled campaigns (per-spreadsheet)

### Important Functions

- `sendCampaign(options)` - Main entry point for sending emails
- `doGet(e)` - Web app handler for tracking (Tracking.gs)
- `recordOpen(trackingId)` / `recordClick(trackingId, url)` - Update spreadsheet tracking columns
- `replaceMergeTags(text, rowData, headers)` - Merge tag substitution

### Gmail API

The Gmail advanced service is enabled for fetching send-as aliases. Check `typeof Gmail !== 'undefined'` before using.

## Deployment Notes

The tracking web app must be deployed with "Anyone" access for open/click tracking to work. After deploying, the URL must be saved via Settings dialog (stored in DocumentProperties as `tracking_web_app_url`).
