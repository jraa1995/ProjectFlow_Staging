# COLONY

Project management tool built on Google Apps Script with Google Sheets as the backing store.

## Stack

- **Runtime:** Google Apps Script (server-side JS)
- **Database:** Google Sheets
- **Frontend:** Vanilla JS, Tailwind CSS (CDN), Chart.js, Font Awesome
- **Client-server:** `google.script.run` — no REST API, everything goes through GAS

## Project Structure

```
core/        # Auth, config, data layer, permissions, board logic, dependencies
engines/     # Automation, triage, notifications, mentions, funnels
services/    # Email notifications
ui/          # HTML templates — index, views, modals, scripts, styles
```

Entry point is `doGet()` in `Code.js`, which serves `ui/Index.html`. All frontend logic lives in `ui/Scripts.html`.

## Caching

Four layers, checked in order: RequestCache (per-request server memo), ScriptCache (GAS built-in), LocalCache (localStorage), IndexedDB (via CacheService).

## Deploy

```bash
clasp push
```

That's it. Open the web app URL in a browser to test. Script ID lives in `.clasp.json`.
