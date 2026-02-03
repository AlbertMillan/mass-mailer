# SheetMailer

A Google Apps Script-based mass email application for Google Sheets, similar to YAMM (Yet Another Mail Merge).

## Features

- **Email Templates**: Use Gmail drafts as email templates
- **Merge Tags**: Personalize emails with `{{ColumnName}}` syntax
- **Full Tracking**: Track opens, clicks, bounces with real-time analytics
- **Gmail Aliases**: Send from different email addresses
- **Google Drive Attachments**: Attach files from Drive, personalized per recipient
- **Scheduling**: Schedule campaigns for later
- **Quota Management**: Auto-schedule remaining emails if quota exceeded
- **Validation**: Email format validation, duplicate detection
- **Progress Tracking**: Real-time progress bar and sheet updates

## Setup Instructions

### 1. Create the Apps Script Project

1. Open Google Sheets
2. Go to **Extensions > Apps Script**
3. Delete any existing code in `Code.gs`
4. Create the following files and paste the code:
   - `Code.gs` - Main application logic
   - `Sidebar.html` - Sidebar user interface
   - `Tracking.gs` - Open/click tracking web app
   - `Settings.html` - Settings dialog
5. Copy `appsscript.json` content to your project manifest:
   - Click the gear icon (Project Settings)
   - Check "Show 'appsscript.json' manifest file"
   - Click on `appsscript.json` in the file list
   - Replace contents with the provided manifest

### 2. Enable Gmail API (for aliases)

1. In Apps Script, click **Services** (+ icon)
2. Search for **Gmail API**
3. Click **Add**

### 3. Deploy the Tracking Web App

1. Click **Deploy > New deployment**
2. Click the gear icon and select **Web app**
3. Configure:
   - Description: "SheetMailer Tracking"
   - Execute as: **Me**
   - Who has access: **Anyone**
4. Click **Deploy**
5. Copy the **Web app URL**
6. In Google Sheets, open **SheetMailer > Settings**
7. Paste the URL and click **Save**

### 4. Authorize the Script

1. Open a Google Sheet with your contact list
2. Refresh the page
3. Click **SheetMailer > Open Sidebar**
4. Follow the authorization prompts

## Usage

### Preparing Your Sheet

1. Create a Google Sheet with your recipients
2. Include an **Email** column (or similar: "email", "Email Address")
3. Add columns for personalization: FirstName, LastName, Company, etc.
4. Optional: Add a filter column (e.g., "Send" with values "Yes"/"No")
5. Optional: Add an attachment column with Google Drive file IDs/URLs

### Creating an Email Template

1. Open Gmail and compose a new email
2. Write your email with merge tags: `Hello {{FirstName}},`
3. **Save as draft** (don't send)
4. The subject line can also contain merge tags

### Sending a Campaign

1. Open your Google Sheet
2. Click **SheetMailer > Open Sidebar**
3. Select your email template (Gmail draft)
4. Choose the "Send From" alias
5. (Optional) Set up filters
6. (Optional) Select attachment column
7. Click **Validate** to check recipients
8. Click **Send Test** to preview
9. Click **Send Campaign** to send to all

### Tracking

After sending, the sheet will update with:
- **Status**: SENT, FAILED, PENDING, INVALID
- **Sent At**: Timestamp when email was sent
- **Opens**: Number of times email was opened
- **Clicks**: Number of link clicks
- **Last Opened**: Timestamp of most recent open

View aggregate stats in the **Stats** tab of the sidebar.

## Multiple Campaigns

Each sheet in your spreadsheet can serve as an independent campaign with its own recipients and tracking data.

### Running Campaigns in Parallel

- **Tracking is independent**: You can run a campaign on Sheet A, then immediately start another on Sheet B. Opens and clicks are tracked separately for each sheet, even while both campaigns are active.
- **Stats are per-sheet**: Switch to any sheet and view its campaign stats independently.

### Sequential Email Sending

While tracking runs in parallel, the email sending process is sequential:
- You must wait for one `sendCampaign()` operation to complete before starting another
- Google Apps Script is single-threaded, so simultaneous sending is not supported
- Once sending completes, you're free to start the next campaign while tracking continues for all previous ones

### Workflow Example

1. Switch to **Sheet A** → Send campaign → Emails sent
2. Switch to **Sheet B** → Send campaign → Emails sent
3. Over the following days, recipients open emails from both campaigns
4. Each sheet's tracking columns update independently

## Quotas & Limits

| Account Type | Daily Email Limit |
|--------------|-------------------|
| Gmail (free) | ~100 recipients/day |
| Google Workspace | ~1,500 recipients/day |

If your campaign exceeds the quota, SheetMailer will:
1. Send as many as possible
2. Automatically schedule the rest for the next day

## Merge Tag Syntax

Use double curly braces with the exact column header name:

```
Hello {{FirstName}},

Thank you for your interest in {{Company}}.

Best regards,
{{SenderName}}
```

## Troubleshooting

### "Email column not detected"
Ensure your sheet has a column named "Email", "email", "Email Address", or similar.

### "Failed to load drafts"
1. Check that you have at least one saved draft in Gmail
2. Re-authorize the script if needed

### Tracking not working
1. Verify the tracking web app URL is set correctly
2. Ensure the web app is deployed with "Anyone" access
3. Check that tracking pixels aren't blocked by email clients

### Emails going to spam
1. Avoid spam trigger words
2. Include an unsubscribe option (recommended)
3. Use a professional "From" name
4. Don't send too many emails too quickly

## File Structure

```
SheetMailer/
├── Code.gs           # Main application logic
├── Sidebar.html      # Sidebar UI (Google Material design)
├── Tracking.gs       # Web app for open/click tracking
├── Settings.html     # Settings dialog
├── appsscript.json   # Project manifest with OAuth scopes
└── README.md         # This file
```

## License

MIT License - Feel free to modify and use for your own projects.
