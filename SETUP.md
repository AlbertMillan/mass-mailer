# SheetMailer Setup Guide

Complete step-by-step instructions to set up SheetMailer in Google Sheets.

---

## Step 1: Create the Apps Script Project

1. Open **Google Sheets** (sheets.google.com)
2. Create a new spreadsheet or open an existing one
3. Click **Extensions > Apps Script**
4. A new Apps Script editor tab will open

---

## Step 2: Add the Code Files

### 2a. Replace Code.gs

1. In the editor, you'll see a file called `Code.gs` with some default code
2. **Delete all the existing code**
3. Copy the entire contents of `Code.gs` from this project
4. Paste it into the editor

### 2b. Create Sidebar.html

1. Click the **+** icon next to "Files" in the left sidebar
2. Select **HTML**
3. Name it `Sidebar` (it will become `Sidebar.html`)
4. Delete any default content
5. Copy and paste the contents of `Sidebar.html` from this project

### 2c. Create Tracking.gs

1. Click the **+** icon next to "Files"
2. Select **Script**
3. Name it `Tracking` (it will become `Tracking.gs`)
4. Delete any default content
5. Copy and paste the contents of `Tracking.gs` from this project

### 2d. Create Settings.html

1. Click the **+** icon next to "Files"
2. Select **HTML**
3. Name it `Settings` (it will become `Settings.html`)
4. Delete any default content
5. Copy and paste the contents of `Settings.html` from this project

---

## Step 3: Update the Manifest File (appsscript.json)

### 3a. Show the manifest file

1. Click the **gear icon** (⚙️) in the left sidebar labeled "Project Settings"
2. Check the box: **Show "appsscript.json" manifest file in editor**
3. Click the **< >** icon (Editor) in the left sidebar to go back to files

### 3b. Edit the manifest

1. Click on `appsscript.json` in the file list
2. **Delete all existing content**
3. Copy and paste the following:

```json
{
  "timeZone": "Europe/Madrid",
  "dependencies": {
    "enabledAdvancedServices": [
      {
        "userSymbol": "Gmail",
        "version": "v1",
        "serviceId": "gmail"
      }
    ]
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "ANYONE"
  },
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets.currentonly",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.settings.basic",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/script.container.ui",
    "https://mail.google.com/"
  ]
}
```

4. **Save the file** (Ctrl+S or Cmd+S)

> **Note:** Change `"timeZone"` to your timezone if needed (e.g., `"America/New_York"`)

---

## Step 4: Enable the Gmail API Service

This allows SheetMailer to access your Gmail aliases.

1. In the left sidebar, find **Services**
2. Click the **+** icon next to Services
3. In the dialog, scroll down to find **Gmail API**
4. Click on **Gmail API** to select it
5. Leave the identifier as `Gmail`
6. Click **Add**

You should now see `Gmail` listed under Services.

---

## Step 5: Deploy the Tracking Web App

This creates the URL for open/click tracking.

### 5a. Start deployment

1. Click **Deploy** in the top menu
2. Select **New deployment**

### 5b. Select deployment type

1. Click the **gear icon** (⚙️) next to "Select type"
2. Choose **Web app**

### 5c. Configure the deployment

Fill in these settings:

| Field | Value |
|-------|-------|
| Description | `SheetMailer Tracking` |
| Execute as | **Me** (your email address) |
| Who has access | **Anyone** |

> **Important:** "Anyone" must be selected for tracking to work.

### 5d. Complete deployment

1. Click **Deploy**
2. If prompted, click **Authorize access**
3. Choose your Google account
4. Click **Advanced** > **Go to [project name] (unsafe)**
5. Click **Allow**

### 5e. Copy the Web App URL

After authorization, you'll see a success screen with:

```
Web app
URL: https://script.google.com/macros/s/AKfycbx.../exec
```

**Copy this URL** (click the copy icon or select and copy)

---

## Step 6: Configure Tracking in SheetMailer

### 6a. Refresh Google Sheets

1. Go back to your Google Sheets tab
2. **Refresh the page** (F5 or Ctrl+R)
3. Wait for the sheet to reload

### 6b. Open Settings

1. You should now see a **SheetMailer** menu in the menu bar
2. Click **SheetMailer > Settings**
3. If prompted to authorize, follow the authorization steps

### 6c. Add the tracking URL

1. In the Settings dialog, find **Tracking Web App URL**
2. Paste the URL you copied in Step 5e
3. Click **Save Settings**

You should see a green message: "Tracking is active"

---

## Step 7: First Run Authorization

1. Click **SheetMailer > Open Sidebar**
2. If you see an authorization prompt:
   - Click **Continue**
   - Choose your Google account
   - Click **Advanced** > **Go to [project name] (unsafe)**
   - Click **Allow**

The sidebar should now open and load your data.

---

## Setup Complete! 🎉

Your file structure should look like this:

```
Files
├── Code.gs
├── Sidebar.html
├── Tracking.gs
├── Settings.html
└── appsscript.json

Services
└── Gmail
```

---

## Troubleshooting

### "SheetMailer menu doesn't appear"
- Refresh the Google Sheets page
- Wait a few seconds for scripts to load
- Check for errors in Apps Script (View > Execution log)

### "Authorization required" error
- Run any function manually from Apps Script (Run > onOpen)
- Complete the authorization flow

### "Specified permissions are not sufficient"
- Make sure `appsscript.json` includes all the OAuth scopes listed above
- Save the file and re-authorize

### "Gmail API not found"
- Go to Services > + > Gmail API > Add
- Make sure it shows as `Gmail` in the Services list

### "Tracking not working"
- Verify the web app URL is correct in Settings
- Make sure deployment has "Anyone" access
- Try creating a new deployment

---

## Updating the Deployment

If you make changes to the code later:

1. Click **Deploy > Manage deployments**
2. Click the **pencil icon** (edit) on your deployment
3. Change "Version" to **New version**
4. Click **Deploy**

This updates the code without changing your tracking URL.
