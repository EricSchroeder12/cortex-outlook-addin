# Cortex Bridge — Outlook Add-in

Reads calendar events and unread emails from all Outlook accounts and sends them to Cortex.

## How it works
1. Cortex starts a local HTTP server on port 5174
2. The add-in reads today's calendar + unread emails via the Outlook REST API
3. Posts the data to `http://localhost:5174/outlook-data`
4. Cortex displays it as a structured briefing on startup — no AI API calls

## Sideloading in Outlook (Windows)

### Option A — Shared Folder (easiest)
1. Create a folder on your machine, e.g. `C:\CortexAddin`
2. Copy `manifest.xml` into that folder
3. In Outlook: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**
4. Add the folder path `C:\CortexAddin` as a catalog, check "Show in Menu"
5. Restart Outlook
6. **Home → Get Add-ins → My Add-ins → Custom Add-ins** → you should see Cortex Bridge

### Option B — Direct sideload
1. Open Outlook
2. **Home → Get Add-ins** (or Insert → Get Add-ins)
3. Click **My add-ins** → **Add a custom add-in** → **Add from file**
4. Select `manifest.xml`
5. Accept the warning

## Using the add-in
- Open any email in Outlook
- Click **Cortex Bridge** in the ribbon (or from the Add-ins button)
- The task pane opens and auto-syncs if Cortex is running
- Click "Send to Cortex" to manually sync

## Requirements
- Cortex must be running (dev mode or installed)
- Outlook desktop on Windows (not Outlook Web)
- Microsoft 365 account connected in Outlook
