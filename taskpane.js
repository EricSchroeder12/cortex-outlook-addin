/* Cortex Bridge — Outlook Add-in Task Pane */

const CORTEX_URL = 'http://localhost:5174/outlook-data';

let officeReady = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    officeReady = true;
    runProbe();            // capability probe FIRST — prove what's reachable
    checkCortexAndSync();
  }
});

// ── Capability probe ─────────────────────────────────────────────────────────
// Office.js is scoped to the mailbox the add-in was activated in. New Outlook
// removed the COM/MAPI path classic Outlook used to enumerate every local
// store, so a single task pane generally CANNOT read the other 3 accounts.
// This probe records exactly what THIS activation can reach so we can confirm
// reachability per account before relying on it. Open the add-in once in each
// of the 4 accounts and compare the output.
function runProbe() {
  const lines = [];
  const log = (k, v) => lines.push(`${k}: ${v}`);
  try {
    const mbx  = Office.context.mailbox;
    const diag = mbx?.diagnostics || {};
    const prof = mbx?.userProfile || {};
    log('host', diag.hostName || '(unknown)');
    log('hostVersion', diag.hostVersion || '(unknown)');
    log('OWAView', diag.OWAView || 'n/a (desktop)');
    log('activeAccount', prof.emailAddress || '(none)');
    log('displayName', prof.displayName || '(none)');
    log('accountType', prof.accountType || '(unknown)');
    log('timeZone', prof.timeZone || '(unknown)');
    log('restUrl', mbx?.restUrl || '(none)');
    log('ewsUrl', mbx?.ewsUrl || '(none)');

    // Can we mint a REST token for THIS mailbox? (other mailboxes are not addressable)
    mbx.getCallbackTokenAsync({ isRest: true }, (r) => {
      const restOk = r.status === Office.AsyncResultStatus.Succeeded;
      log('restToken', restOk ? 'OK (this mailbox only)' : `FAILED (${r.error?.message || r.status})`);
      lines.push('');
      lines.push('NOTE: Office.js reaches only the active mailbox. To cover all 4');
      lines.push('accounts, open this add-in once per account (each run POSTs its');
      lines.push('slice; Cortex merges by account), or use Graph + Google OAuth.');
      renderProbe(lines.join('\n'));
      console.log('[Cortex probe]\n' + lines.join('\n'));

      // Report capabilities to Cortex so the bridge diagnostics panel can show
      // a per-account verdict even before a full sync runs.
      const diagnostics = {
        host:        diag.hostName || null,
        hostVersion: diag.hostVersion || null,
        accountType: prof.accountType || null,
        timeZone:    prof.timeZone || null,
        restUrl:     mbx?.restUrl || null,
        restToken:   restOk ? 'ok' : 'failed',
        healthy:     restOk && !!mbx?.restUrl,
      };
      postProbe(prof.emailAddress || 'unknown', diagnostics);
    });
  } catch (e) {
    lines.push('probe error: ' + e.message);
    renderProbe(lines.join('\n'));
  }
}

function renderProbe(text) {
  const el = document.getElementById('probe');
  if (el) { el.textContent = text; el.style.display = 'block'; }
}

// POST probe/capability results to the Cortex bridge (best-effort).
async function postProbe(account, diagnostics) {
  try {
    await fetch('http://localhost:5174/outlook-probe', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ account, diagnostics, timestamp: new Date().toISOString() }),
      signal: AbortSignal.timeout(3000),
    });
  } catch { /* Cortex not running — probe still shows locally in the pane */ }
}

async function checkCortexAndSync() {
  const running = await isCortexRunning();
  if (running) {
    setStatus('connected', 'Cortex is running');
    document.getElementById('syncBtn').disabled = false;
    await sendToCortex();
  } else {
    setStatus('disconnected', 'Cortex is not running — open Cortex first');
    document.getElementById('syncBtn').disabled = true;
  }
}

async function isCortexRunning() {
  try {
    const res = await fetch('http://localhost:5174/ping', { method: 'GET', signal: AbortSignal.timeout(2000) });
    return res.ok;
  } catch {
    return false;
  }
}

async function sendToCortex() {
  if (!officeReady) return;

  setStatus('syncing', 'Reading Outlook data...');
  document.getElementById('syncBtn').disabled = true;

  try {
    const [calendar, emails] = await Promise.all([
      getCalendarEvents(),
      getRecentEmails()
    ]);

    const mbx = Office.context.mailbox;
    const payload = {
      // Top-level account so the Cortex receiver can merge per-account runs
      // into one unified, account-grouped briefing (path a — per-account run).
      account: mbx.userProfile?.emailAddress || 'unknown',
      timestamp: new Date().toISOString(),
      diagnostics: {
        host:        mbx.diagnostics?.hostName || null,
        hostVersion: mbx.diagnostics?.hostVersion || null,
        accountType: mbx.userProfile?.accountType || null,
        timeZone:    mbx.userProfile?.timeZone || null,
        restUrl:     mbx.restUrl || null,
        restToken:   'ok',          // we just used it to fetch calendar+mail
        healthy:     true,
      },
      calendar,
      emails
    };

    setStatus('syncing', `Sending ${calendar.length} events, ${emails.length} emails to Cortex...`);

    const res = await fetch(CORTEX_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    if (res.ok) {
      setStatus('connected', `Synced: ${calendar.length} events, ${emails.length} emails`);
      document.getElementById('lastSync').textContent = 'Last sync: ' + new Date().toLocaleTimeString();
      showSummary(calendar, emails);
    } else {
      setStatus('disconnected', 'Cortex returned an error');
    }
  } catch (err) {
    setStatus('disconnected', 'Failed to reach Cortex: ' + err.message);
  } finally {
    document.getElementById('syncBtn').disabled = false;
  }
}

async function getCalendarEvents() {
  return new Promise((resolve) => {
    try {
      // Get today's date range
      const now = new Date();
      const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
      const endOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);

      // Use REST API via callback token for calendar access
      Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          resolve([]);
          return;
        }

        try {
          const token = result.value;
          const restUrl = Office.context.mailbox.restUrl;
          const startStr = startOfDay.toISOString();
          const endStr = endOfDay.toISOString();

          const url = `${restUrl}/v2.0/me/calendarview?startDateTime=${encodeURIComponent(startStr)}&endDateTime=${encodeURIComponent(endStr)}&$select=subject,start,end,organizer,location,bodyPreview&$top=50`;

          const response = await fetch(url, {
            headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' }
          });

          if (!response.ok) { resolve([]); return; }

          const data = await response.json();
          const events = (data.value || []).map(e => ({
            subject: e.subject || '(No subject)',
            start: e.start?.dateTime || '',
            end: e.end?.dateTime || '',
            organizer: e.organizer?.emailAddress?.name || '',
            location: e.location?.displayName || '',
            account: Office.context.mailbox.userProfile?.emailAddress || ''
          }));
          resolve(events);
        } catch {
          resolve([]);
        }
      });
    } catch {
      resolve([]);
    }
  });
}

async function getRecentEmails() {
  return new Promise((resolve) => {
    try {
      Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          resolve([]);
          return;
        }

        try {
          const token = result.value;
          const restUrl = Office.context.mailbox.restUrl;

          // Get unread emails from last 24 hours
          const since = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();
          const url = `${restUrl}/v2.0/me/mailfolders/inbox/messages?$filter=isRead eq false and receivedDateTime ge ${encodeURIComponent(since)}&$select=subject,from,receivedDateTime,bodyPreview&$top=50&$orderby=receivedDateTime desc`;

          const response = await fetch(url, {
            headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' }
          });

          if (!response.ok) { resolve([]); return; }

          const data = await response.json();
          const emails = (data.value || []).map(e => ({
            subject: e.subject || '(No subject)',
            from: e.from?.emailAddress?.name || e.from?.emailAddress?.address || '',
            receivedAt: e.receivedDateTime || '',
            preview: (e.bodyPreview || '').substring(0, 120),
            account: Office.context.mailbox.userProfile?.emailAddress || ''
          }));
          resolve(emails);
        } catch {
          resolve([]);
        }
      });
    } catch {
      resolve([]);
    }
  });
}

function showSummary(calendar, emails) {
  const el = document.getElementById('summary');
  if (!calendar.length && !emails.length) { el.style.display = 'none'; return; }

  let html = '';
  if (calendar.length) {
    html += `<strong>Today (${calendar.length} events)</strong><br>`;
    calendar.forEach(e => {
      const t = e.start ? new Date(e.start).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '';
      html += `${t} &mdash; ${e.subject}<br>`;
    });
  }
  if (emails.length) {
    if (html) html += '<br>';
    html += `<strong>Unread (${emails.length})</strong><br>`;
    emails.slice(0, 10).forEach(e => {
      html += `${e.from}: ${e.subject}<br>`;
    });
    if (emails.length > 10) html += `...and ${emails.length - 10} more`;
  }

  el.innerHTML = html;
  el.style.display = 'block';
}

function setStatus(type, msg) {
  const el = document.getElementById('status');
  el.className = 'status ' + type;
  el.textContent = msg;
}
