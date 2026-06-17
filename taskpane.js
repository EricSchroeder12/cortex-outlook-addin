/* Cortex Bridge — Outlook Add-in Task Pane */

const CORTEX_URL = 'http://localhost:5174/outlook-data';

let officeReady = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    officeReady = true;
    checkCortexAndSync();
  }
});

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

    const payload = {
      timestamp: new Date().toISOString(),
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
