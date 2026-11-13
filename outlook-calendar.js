/*
 * Outlook calendar integration (Microsoft Graph + MSAL.js).
 *
 * Adds an "Outlook connection" panel to the Settings page "Calendar integration" card:
 *  - Sign in with a Microsoft 365 / Outlook account (OAuth 2.0 auth code + PKCE, popup).
 *  - Pick which Outlook calendar holds design reviews (own or shared-with-me calendars).
 *  - Pull Outlook events into the Review calendar view (shown as "Outlook" chips).
 *  - Push a review checkpoint to Outlook from the "Review event details" panel
 *    (creates the event the first time, updates it afterwards).
 *
 * Requires an Azure (Entra ID) app registration of type "Single-page application" whose
 * redirect URI matches the URL this page is served from. See README.md.
 * All configuration is kept in localStorage under `outlookCalendarConfig`.
 */
(function () {
  const MSAL_SRC = 'https://alcdn.msauth.net/browser/2.38.3/js/msal-browser.min.js';
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const SCOPES = ['User.Read', 'Calendars.ReadWrite', 'Calendars.ReadWrite.Shared'];
  const STORE = 'outlookCalendarConfig';

  const cfg = Object.assign(
    { clientId: '', tenant: 'common', calendarId: '', calendarName: '', showEvents: true, defaultDuration: 60, timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone },
    safeParse(localStorage.getItem(STORE))
  );
  let msalApp = null, account = null, outlookEvents = [], calendars = [];

  function safeParse(v) { try { return JSON.parse(v || '{}') || {}; } catch { return {}; } }
  function save() { localStorage.setItem(STORE, JSON.stringify(cfg)); }
  function esc(v) { return String(v ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
  function notify(msg) { (window.toast || console.log)(msg); }

  /* ---------- MSAL / auth ---------- */
  function loadMsal() {
    if (window.msal) return Promise.resolve();
    return new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = MSAL_SRC; s.onload = res; s.onerror = () => rej(new Error('Could not load Microsoft sign-in library'));
      document.head.appendChild(s);
    });
  }

  async function getMsal() {
    if (!cfg.clientId) throw new Error('Enter an Azure application (client) ID first');
    if (location.protocol === 'file:') throw new Error('Serve the template over http(s) to sign in (see README)');
    if (msalApp && msalApp.__clientId === cfg.clientId + cfg.tenant) return msalApp;
    await loadMsal();
    msalApp = new msal.PublicClientApplication({
      auth: { clientId: cfg.clientId, authority: `https://login.microsoftonline.com/${cfg.tenant || 'common'}`, redirectUri: location.origin + location.pathname },
      cache: { cacheLocation: 'localStorage' }
    });
    await msalApp.initialize();
    await msalApp.handleRedirectPromise().catch(() => null);
    msalApp.__clientId = cfg.clientId + cfg.tenant;
    account = msalApp.getAllAccounts()[0] || null;
    return msalApp;
  }

  async function token() {
    const app = await getMsal();
    if (!account) throw new Error('Not signed in to Outlook');
    try {
      return (await app.acquireTokenSilent({ scopes: SCOPES, account })).accessToken;
    } catch {
      return (await app.acquireTokenPopup({ scopes: SCOPES, account })).accessToken;
    }
  }

  async function graph(path, opts = {}) {
    const res = await fetch(GRAPH + path, {
      ...opts,
      headers: { Authorization: `Bearer ${await token()}`, 'Content-Type': 'application/json', Prefer: `outlook.timezone="${cfg.timeZone}"`, ...(opts.headers || {}) }
    });
    if (!res.ok) {
      const body = await res.json().catch(() => ({}));
      throw new Error(body.error?.message || `Microsoft Graph error ${res.status}`);
    }
    return res.status === 204 ? null : res.json();
  }

  async function connect() {
    const app = await getMsal();
    const r = await app.loginPopup({ scopes: SCOPES, prompt: 'select_account' });
    account = r.account;
    await loadCalendars();
    await syncEvents();
    notify(`Connected to Outlook as ${account.username}`);
  }

  async function disconnect() {
    if (msalApp && account) await msalApp.logoutPopup({ account }).catch(() => null);
    account = null; outlookEvents = []; calendars = [];
    rerenderCalendar(); renderPanel();
    notify('Disconnected from Outlook');
  }

  /* ---------- Calendars & events ---------- */
  function calPath() { return cfg.calendarId ? `/me/calendars/${encodeURIComponent(cfg.calendarId)}` : '/me/calendar'; }

  async function loadCalendars() {
    const r = await graph('/me/calendars?$select=id,name,owner,canEdit&$top=100');
    calendars = r.value || [];
    renderPanel();
  }

  async function syncEvents() {
    const start = new Date(); start.setMonth(start.getMonth() - 2);
    const end = new Date(); end.setMonth(end.getMonth() + 4);
    const q = `?startDateTime=${start.toISOString()}&endDateTime=${end.toISOString()}&$select=id,subject,start,end,webLink,organizer,location&$orderby=start/dateTime&$top=500`;
    const r = await graph(`${calPath()}/calendarView${q}`);
    outlookEvents = r.value || [];
    cfg.lastSync = new Date().toISOString(); save();
    rerenderCalendar(); renderPanel();
    return outlookEvents.length;
  }

  function reviewerEmail(handle) {
    const name = String(handle || '').replace(/^@/, '').trim().toLowerCase();
    if (!name) return '';
    const roster = window.reviewTeamRoster || [];
    const hit = roster.find(p => (p.name || '').toLowerCase() === name || (p.name || '').toLowerCase().replace(/\s+/g, '') === name.replace(/\s+/g, ''));
    return hit?.email || (/@.+\./.test(name) ? name : '');
  }

  function buildEvent(t, j) {
    const date = t.scheduledDates?.[j] || t.reviewDates?.[j];
    if (!date) throw new Error('Set a scheduled date for this review first');
    const time = t.meetingTimes?.[j] || '10:00';
    const startLocal = `${date}T${time}:00`;
    const endDate = new Date(`${startLocal}Z`); endDate.setUTCMinutes(endDate.getUTCMinutes() + (+cfg.defaultDuration || 60));
    const endLocal = endDate.toISOString().slice(0, 19);
    const lines = [
      `<b>${esc(reviewStages[j])}</b> · ${esc(t.project)} — ${esc(t.subjects?.[j] || t.task)}`,
      `Requester: ${esc(t.requesters?.[j] || t.resource || '')}`,
      `Reviewer: ${esc(t.reviewers?.[j] || 'Unassigned')}`,
      t.contentLinks?.[j] ? `Content to review: <a href="${esc(t.contentLinks[j])}">${esc(t.contentLinks[j])}</a>` : '',
      (t.previousNotesLinks?.[j] || t.notes?.[j]) ? `Previous notes: <a href="${esc(t.previousNotesLinks?.[j] || t.notes[j])}">link</a>` : '',
      t.meetingLinks?.[j] ? `Join: <a href="${esc(t.meetingLinks[j])}">${esc(t.meetingLinks[j])}</a>` : ''
    ].filter(Boolean);
    const email = reviewerEmail(t.reviewers?.[j]);
    return {
      subject: t.eventTitles?.[j] || `${reviewStages[j]} · ${t.project}`,
      body: { contentType: 'HTML', content: lines.join('<br>') },
      start: { dateTime: startLocal, timeZone: cfg.timeZone },
      end: { dateTime: endLocal, timeZone: cfg.timeZone },
      location: t.meetingLinks?.[j] ? { displayName: t.meetingLinks[j] } : undefined,
      attendees: email ? [{ emailAddress: { address: email }, type: 'required' }] : [],
      categories: ['Design Review']
    };
  }

  async function pushReview(i, j) {
    const t = reviewTasks[i];
    t.outlookEventIds ??= [];
    const body = buildEvent(t, j);
    let ev;
    if (t.outlookEventIds[j]) {
      try { ev = await graph(`/me/events/${encodeURIComponent(t.outlookEventIds[j])}`, { method: 'PATCH', body: JSON.stringify(body) }); }
      catch { t.outlookEventIds[j] = ''; }
    }
    if (!ev) ev = await graph(`${calPath()}/events`, { method: 'POST', body: JSON.stringify(body) });
    t.outlookEventIds[j] = ev.id;
    (window.persistReviewDisplayConfig || window.saveDashboardConfiguration)?.();
    await syncEvents().catch(() => null);
    return ev;
  }

  /* ---------- Review calendar overlay ---------- */
  function rerenderCalendar() { if (typeof window.renderReviewCalendar === 'function') window.renderReviewCalendar(); }

  const baseRender = window.renderReviewCalendar;
  if (typeof baseRender === 'function') {
    window.renderReviewCalendar = renderReviewCalendar = function () {
      baseRender.apply(this, arguments);
      if (!cfg.showEvents || !account || !outlookEvents.length) return;
      const linked = new Set(reviewTasks.flatMap(t => t.outlookEventIds || []).filter(Boolean));
      document.querySelectorAll('#reviewCalendarGrid .calendar-day[data-date]').forEach(cell => {
        const ds = cell.dataset.date, host = cell.querySelector('.calendar-events');
        const evs = outlookEvents.filter(e => (e.start?.dateTime || '').slice(0, 10) === ds && !linked.has(e.id));
        if (!host || !evs.length) return;
        host.insertAdjacentHTML('beforeend', evs.slice(0, 2).map(e =>
          `<a class="calendar-event outlook-event" href="${esc(e.webLink)}" target="_blank" rel="noopener" title="Outlook · ${esc(e.subject)} · ${esc(e.start.dateTime.slice(11, 16))}"><strong>Outlook · ${esc(e.start.dateTime.slice(11, 16))}</strong><span>${esc(e.subject || '(no subject)')}</span></a>`
        ).join('') + (evs.length > 2 ? `<span class="calendar-more">+${evs.length - 2} Outlook</span>` : ''));
      });
    };
  }

  const baseDetails = window.openCalendarMeetingDetails;
  if (typeof baseDetails === 'function') {
    window.openCalendarMeetingDetails = openCalendarMeetingDetails = function (i, j) {
      baseDetails.apply(this, arguments);
      const host = document.querySelector('#reviewDisplayInspectorContent .calendar-meeting-details');
      if (!host) return;
      const linked = reviewTasks[i]?.outlookEventIds?.[j];
      host.insertAdjacentHTML('beforeend', `<div class="outlook-push"><button type="button" class="btn" id="pushToOutlook">${linked ? 'Update Outlook event' : 'Add to Outlook calendar'}</button><p class="help" id="pushToOutlookStatus">${account ? (linked ? 'Linked to an Outlook event.' : `Creates the event in ${esc(cfg.calendarName || 'your default calendar')}.`) : 'Connect Outlook in Settings to enable.'}</p></div>`);
      const btn = host.querySelector('#pushToOutlook'), status = host.querySelector('#pushToOutlookStatus');
      btn.disabled = !account;
      btn.onclick = async () => {
        // Persist any edits made in this panel before pushing.
        // (Saving re-renders the panel, so look the controls up again afterwards.)
        host.querySelector('#saveMeetingDetails')?.click();
        const btn = document.querySelector('#pushToOutlook'), status = document.querySelector('#pushToOutlookStatus');
        btn.disabled = true; status.textContent = 'Sending to Outlook…';
        try { await pushReview(i, j); status.textContent = 'Saved to Outlook.'; btn.textContent = 'Update Outlook event'; notify('Review saved to Outlook'); }
        catch (e) { status.textContent = e.message; }
        finally { btn.disabled = false; }
      };
    };
  }

  /* ---------- Settings panel ---------- */
  const card = document.querySelector('#settingsCalendarIntegration');
  if (!card) return;
  card.querySelector('p').textContent = 'Connect a shared Outlook calendar to view and schedule design reviews.';
  const panel = document.createElement('div');
  panel.className = 'outlook-panel';
  card.appendChild(panel);

  const style = document.createElement('style');
  style.textContent = `.outlook-panel{margin-top:16px;padding-top:14px;border-top:1px solid var(--line);display:grid;gap:6px}.outlook-panel .field{max-width:none!important}.outlook-panel .row{display:grid;grid-template-columns:1fr 1fr;gap:8px}.outlook-panel .opt{display:flex;gap:6px;align-items:center;font-size:11px;color:#657067}.outlook-panel .opt input{accent-color:#315a43}.outlook-panel .who{font-size:11px;color:#315a43}.calendar-event.outlook-event{display:grid;text-decoration:none;border-left:3px solid #0f6cbd;background:#eef5fc;color:#20364d}.outlook-push{margin-top:10px;display:grid;gap:4px}`;
  document.head.appendChild(style);

  function renderPanel() {
    const connected = !!account;
    panel.innerHTML = `
      <div class="fieldlabel" style="margin:0">Outlook connection</div>
      ${connected ? `<div class="who">● Connected as ${esc(account.username)}</div>` : ''}
      <div class="row">
        <label><span class="fieldlabel">Application (client) ID</span><input class="field" id="outlookClientId" value="${esc(cfg.clientId)}" placeholder="00000000-0000-…" ${connected ? 'disabled' : ''}></label>
        <label><span class="fieldlabel">Tenant</span><input class="field" id="outlookTenant" value="${esc(cfg.tenant)}" placeholder="common or contoso.onmicrosoft.com" ${connected ? 'disabled' : ''}></label>
      </div>
      ${connected ? `
      <label><span class="fieldlabel">Design review calendar</span><select class="field" id="outlookCalendarSelect"><option value="">Default calendar</option>${calendars.map(c => `<option value="${esc(c.id)}" ${c.id === cfg.calendarId ? 'selected' : ''}>${esc(c.name)}${c.owner?.address && c.owner.address !== account.username ? ` (${esc(c.owner.name || c.owner.address)})` : ''}${c.canEdit === false ? ' · read-only' : ''}</option>`).join('')}</select></label>
      <div class="row">
        <label><span class="fieldlabel">Default length (min)</span><input class="field" id="outlookDuration" type="number" min="15" step="15" value="${esc(cfg.defaultDuration)}"></label>
        <label class="opt" style="align-self:end"><input type="checkbox" id="outlookShowEvents" ${cfg.showEvents ? 'checked' : ''}> Show Outlook events on Review calendar</label>
      </div>` : ''}
      <div class="settings-actions" style="margin-top:6px">
        ${connected
          ? `<button class="btn primary" id="outlookSync" type="button">↻ Sync now</button><button class="btn" id="outlookDisconnect" type="button">Disconnect</button>`
          : `<button class="btn primary" id="outlookConnect" type="button">Connect Outlook</button>`}
      </div>
      <div class="settings-source" id="outlookStatus">${connected ? `${outlookEvents.length} events loaded${cfg.lastSync ? ' · last sync ' + new Date(cfg.lastSync).toLocaleString() : ''}` : 'Not connected to Outlook'}</div>`;

    const status = panel.querySelector('#outlookStatus');
    const run = async (fn) => { try { await fn(); } catch (e) { status.textContent = e.message; } };
    panel.querySelector('#outlookClientId')?.addEventListener('change', e => { cfg.clientId = e.target.value.trim(); save(); });
    panel.querySelector('#outlookTenant')?.addEventListener('change', e => { cfg.tenant = e.target.value.trim() || 'common'; save(); });
    panel.querySelector('#outlookConnect')?.addEventListener('click', () => run(async () => { status.textContent = 'Opening Microsoft sign-in…'; await connect(); }));
    panel.querySelector('#outlookDisconnect')?.addEventListener('click', () => run(disconnect));
    panel.querySelector('#outlookSync')?.addEventListener('click', () => run(async () => { status.textContent = 'Syncing…'; const n = await syncEvents(); notify(`Synced ${n} Outlook events`); }));
    panel.querySelector('#outlookCalendarSelect')?.addEventListener('change', e => run(async () => {
      cfg.calendarId = e.target.value; cfg.calendarName = e.target.selectedOptions[0]?.textContent || ''; save(); await syncEvents();
    }));
    panel.querySelector('#outlookDuration')?.addEventListener('change', e => { cfg.defaultDuration = Math.max(15, +e.target.value || 60); save(); });
    panel.querySelector('#outlookShowEvents')?.addEventListener('change', e => { cfg.showEvents = e.target.checked; save(); rerenderCalendar(); });
  }

  renderPanel();

  // Restore an existing session silently on load.
  if (cfg.clientId && location.protocol !== 'file:') {
    getMsal().then(async () => {
      if (!account) return;
      renderPanel();
      await loadCalendars().catch(() => null);
      await syncEvents().catch(() => null);
    }).catch(() => null);
  }

  window.outlookCalendar = { connect, disconnect, syncEvents, pushReview, get events() { return outlookEvents; } };
})();
