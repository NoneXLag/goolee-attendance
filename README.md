# Goolee.my Attendance Clock

A single-page, mobile-first attendance clock: employees pick their name, enter
a 4-digit PIN, and punch in/out. GPS accuracy ≤ 50 m is mandatory on every
punch — no fallback, no IP-only mode.

## What's in this zip

```
index.html            The attendance app (open this in a browser / host it)
apps-script/Code.gs    Google Apps Script backend (Google Sheets storage)
README.md              This file
```

## 1. Set up the Google Sheet

Create a new Google Sheet with two tabs:

**Employees**

| Name       | PIN  | Active |
|------------|------|--------|
| Aisyah Bt. | 4821 | TRUE   |
| Ravi K.    | 0093 | TRUE   |

- Format the **PIN** column as *Plain text* first, so PINs with a leading
  zero (e.g. `0093`) aren't stripped to a number.
- Set `Active` to `FALSE` to immediately remove someone from the dropdown —
  no further punches are accepted for that name.

**Attendance** — leave the header row, the script appends rows here:

| Timestamp | Name | Action | Location (Google Maps) | Accuracy (m) | Date | Time |
|---|---|---|---|---|---|---|

The **Location (Google Maps)** cell is a clickable `HYPERLINK("...","Open in Maps")`
formula pointing at `https://www.google.com/maps?q=lat,lng` — click it from
the sheet to open the exact punch location on a map.

## 2. Deploy the backend

1. In the Sheet, open **Extensions → Apps Script**.
2. Delete the placeholder code and paste in `apps-script/Code.gs`.
3. Click **Deploy → New deployment → Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
4. Copy the deployment URL (ends in `/exec`).

## 3. Connect the frontend

Open `index.html` and edit the config block near the top of the `<script>`:

```js
const CONFIG = {
  API_URL: "https://script.google.com/macros/s/XXXXX/exec",
  SHEET_URL: "https://docs.google.com/spreadsheets/d/XXXXX/edit",
  ...
};
```

## 4. Host it

Upload `index.html` anywhere that serves static files over HTTPS (GitHub
Pages, Netlify, Vercel, your own web host) and point `attendance.goolee.my`
at it. Geolocation requires HTTPS (or `localhost` for testing) in all modern
browsers.

## Behavior notes

- **GPS policy**: `enableHighAccuracy: true`, 15 s timeout, no cached
  positions, up to 3 retries with a 2 s pause between attempts. Readings
  between 50–100 m and above 100 m are rejected and the user is asked to
  move; punch buttons stay disabled until a reading ≤ 50 m locks in.
- **PIN lockout**: 5 incorrect PINs locks that employee out for 5 minutes.
  This is enforced both client-side (for instant feedback) and
  server-side in `Code.gs` via `PropertiesService`, since Apps Script web
  apps don't expose the caller's IP address — the PRD's "tracked per IP"
  requirement is approximated as "tracked per employee" for that reason.
- **HR-only link**: the "Open Google Sheet →" link is included but hidden
  (`display:none`) by default, since this static page has no real access
  control. Wire it up to your own auth (e.g. an HR PIN gate, or simply
  give HR a separate bookmarked link) before exposing it.
- **PDPA notice**: collapsible, closed by default, as specified.

## What was intentionally left out

Per the PRD's own scope notes: no device binding, no WebAuthn/Face ID, no
selfie-at-punch. These were called out as unnecessary for a team under 50
workers on a PIN-based system.
