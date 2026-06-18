# Cortex Bridge — Outlook Add-in

Reads calendar events and unread emails from all Outlook accounts and sends them to Cortex.

## How it works
1. Cortex starts a local HTTP server on port 5174
2. The add-in reads today's calendar + unread emails via the Outlook REST API
3. Posts the data (tagged with the active `account`) to `http://localhost:5174/outlook-data`
4. Cortex **merges each run by account** into one unified, account-grouped cache
   (stale accounts >12h are dropped) and displays it as a structured briefing —
   no AI API calls

## Multi-account coverage
Office.js is scoped to the mailbox the add-in is activated in — one task pane
cannot read the other accounts. Open the add-in **once in each of your accounts**
and click "Send to Cortex"; each run POSTs its own slice and Cortex merges them.
Click **Run capability probe** to see exactly what the current mailbox exposes
(host, active account, REST reachability) before relying on it.

## Hosting (required)
The add-in pages must be served over **HTTPS** — New Outlook (and Outlook on the
web) run the task pane in Edge WebView2 and reject `file://`/`http://` sources.
These files are published to GitHub Pages at
`https://ericschroeder12.github.io/cortex-outlook-addin/`. After editing any file
here (`taskpane.html`, `taskpane.js`, `manifest.xml`, the `icon-*.png`s), **push
them to that Pages repo** so the hosted copy updates, then bump `<Version>` in
`manifest.xml` and re-sideload.

## Sideloading in New Outlook (Windows)
> The classic **Trust Center → Trusted Add-in Catalogs** (shared-folder) method
> does **not** exist in New Outlook — use the web-style flow below. New Outlook
> shares cloud-installed add-ins with Outlook on the web, so installing in either
> place shows it in both.

1. In **New Outlook**, open the **Apps** button on the ribbon (or **More apps `···` → Get Add-ins**).
2. Choose **My add-ins → Custom add-ins → Add a custom add-in → Add from URL**.
3. Paste the hosted manifest URL:
   `https://ericschroeder12.github.io/cortex-outlook-addin/manifest.xml`
   (If "Add from file" is offered, you can instead pick the local `manifest.xml` — but "Add from URL" is the reliable path for New Outlook.)
4. Accept the prompt. A **Cortex → Cortex Bridge** button appears on the ribbon
   when reading a message or appointment (provided by the `VersionOverrides` block).

**Alternative (same result):** sign in at `https://outlook.office.com`, then
**Settings → General → Manage add-ins** (or the **Get add-ins** dialog) → **My add-ins
→ Add a custom add-in → Add from URL** with the same manifest URL.

> If your org blocks custom/unverified add-ins, sideloading is disabled and an
> admin must allow it (Microsoft 365 admin center → Integrated apps / "Deploy Add-in"),
> or deploy the manifest org-wide.

## Using the add-in
- Open any email in Outlook
- Click **Cortex Bridge** in the ribbon (or from the Add-ins button)
- The task pane opens and auto-syncs if Cortex is running
- Click "Send to Cortex" to manually sync

## Requirements
- Cortex must be running (dev mode or installed) — it serves the bridge on `http://localhost:5174`
- **New Outlook (Windows)** or **Outlook on the web** (the classic desktop client is no longer the target)
- A mailbox connected to that account. Office.js reaches only the *active* mailbox,
  so open the add-in once per account to cover all of them (Cortex merges by account).

> Note on the local call: the HTTPS-hosted task pane fetches `http://localhost:5174`.
> Chromium/WebView2 treats `localhost` as trustworthy and allows this despite the
> page being HTTPS, so no mixed-content block — but it only works on the same
> machine Cortex is running on.
