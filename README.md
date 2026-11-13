# UX Design Review Template

A self-contained, browser-only dashboard for running a UX design review program: track projects through review checkpoints, schedule reviews on a calendar, take intake requests, manage reviewers, and (optionally) sync the review calendar with Outlook.

You don't need to build anything, and there's no backend. Everything runs from `index.html` and a few plain JavaScript files. All data and settings are saved in the browser's `localStorage`.

## What it does

| Page (sidebar) | Purpose |
| --- | --- |
| **Overview / Review dashboard** | KPI cards, configurable charts (bar, line, donut, Gantt), and a project table. Click a chart or a KPI card to open its inspector and change its type, size, or metric. |
| **Review calendar** | The review queue, a milestone timeline, and a week/month calendar. Each project moves through four checkpoints: *Design Review 1*, *Design Review 2*, *Follow-up Design Review*, and *Fit & Finish Review*. Click an event to edit its title, reviewer, links, meeting time, and video link. |
| **Projects** | Project list with editable status labels. |
| **Intake** | Kanban board for incoming review requests. Lanes can be renamed, and cards can be dragged between them. |
| **People** | Review team roster (name, role, email) and reviewer workload. |
| **Reports** | Summary reporting views. |
| **Settings** | Dashboard layout, CSV/JSON import and export, branding, **calendar integration (Outlook)**, sidebar labels and visibility, and color theme. |

### Files

| File | Responsibility |
| --- | --- |
| `index.html` | Markup, styles, sample data, dashboard, charts, and the review calendar |
| `chart-renderers.js` | SVG chart drawing |
| `metrics.js` | KPI cards and the Settings page Save button (branding and shared calendar URL) |
| `task-board.js` | Intake kanban board |
| `people.js` | People page and team roster |
| `reports.js`, `share-pages.js` | Reports page and page sharing |
| `status-labels.js` | Custom project status labels |
| `theme-picker.js` | Color theme |
| `outlook-calendar.js` | Outlook / Microsoft 365 calendar connection |
| `sources/` | Read-only reference files synced from the ChatGPT "UXDR" project. **Do not edit.** |

## Getting started

1. Serve the folder over HTTP. You need this for the Outlook sign-in; everything else also works if you open the file directly:
   ```bash
   cd design-review-template
   python3 -m http.server 8080
   ```
2. Open http://localhost:8080/index.html.
3. Go to **Settings** and set your brand and workspace names, grid columns, and theme, then click **Save**.
4. Replace the sample data. Either use **Settings → Data and sharing → Import CSV/JSON**, or edit projects directly from the Review calendar.
5. Add reviewers on the **People** page. Include their emails so Outlook invites can go to them.

To reset, clear the site's `localStorage` in your browser dev tools.

## Outlook calendar integration

**Settings → Calendar integration → Outlook connection** connects the template to a Microsoft 365 / Outlook.com calendar through Microsoft Graph. The sign-in uses MSAL.js (OAuth with PKCE, popup).

Once you're connected, you can:

- **Pick the design review calendar.** Choose your default calendar, any other calendar you own, or a calendar shared with you (for example, a team "Design Reviews" calendar).
- **See Outlook events on the Review calendar.** Events from the selected calendar show up as blue **Outlook** chips. Clicking a chip opens the event in Outlook. To hide them, turn off *Show Outlook events on Review calendar*.
- **Push reviews to Outlook.** Open a review event on the calendar and click **Add to Outlook calendar**. This creates the event with:
  - the title, scheduled date, and time (10:00 if no time is set)
  - the default length you chose
  - the content, notes, and meeting links
  - the assigned reviewer as an attendee, if their email is in **People**

  After that, the button changes to **Update Outlook event** and updates the same event instead of making a duplicate.
- **Sync now.** Refreshes events from 2 months back to 4 months ahead.

### One-time setup: register an app in Microsoft Entra ID

You (or your M365 admin) need to do this once.

1. Go to https://entra.microsoft.com → **App registrations → New registration**.
2. **Name**: `Design Review Template`.
3. **Supported account types**: choose *Single tenant* for your organization only, or *Multitenant + personal accounts* to also allow Outlook.com.
4. **Redirect URI**: platform **Single-page application (SPA)**. Enter the exact URL you serve the page from, for example `http://localhost:8080/index.html`. Also add your production URL if you host it.
5. Open **API permissions → Add → Microsoft Graph → Delegated** and add:
   - `User.Read`
   - `Calendars.ReadWrite`
   - `Calendars.ReadWrite.Shared` (needed for shared or team calendars)

   Grant admin consent if your tenant requires it.
6. Copy the **Application (client) ID**. Also copy the **Directory (tenant) ID**, or use your domain.

### Connect

1. In **Settings → Calendar integration**, paste the **Application (client) ID**.
2. Set **Tenant**. Use your tenant ID or domain for a single-tenant app, or leave it as `common`.
3. Click **Connect Outlook** and sign in in the popup.
4. Choose the **Design review calendar** and set a **Default length**.

The connection is restored automatically when you reload the page. Use **Disconnect** to sign out.

> [!NOTE]
> - The **Shared calendar URL** field (`https://` or `webcal://`) is still available for reference links. It isn't used for the Graph connection.
> - Tokens are stored by MSAL in `localStorage`. Settings are stored under `outlookCalendarConfig`. Nothing is sent anywhere except Microsoft's login and Graph endpoints.
> - Sign-in won't work from a `file://` URL. Serve the folder over `http(s)` and make sure the URL matches the registered redirect URI exactly.
> - Pop-up blockers can stop sign-in. Allow pop-ups for the site.

### Troubleshooting

| Message | Fix |
| --- | --- |
| `AADSTS50011` redirect URI mismatch | Add the exact page URL, including `/index.html`, as an SPA redirect URI |
| `AADSTS65001` consent required | An admin needs to grant consent for the Graph permissions |
| "Serve the template over http(s)…" | Run a local server (see Getting started) |
| Shared calendar missing from the list | Accept the calendar share in Outlook first, and make sure `Calendars.ReadWrite.Shared` is granted |
| Reviewer not invited | Add the reviewer's email on **People**. Their name there must match the reviewer `@name` |
