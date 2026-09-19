Goolee.my — Attendance System
==============================

A complete GPS-verified attendance system built on Google Sheets + Google Apps Script + a static HTML frontend.

Employees clock in/out from any phone with GPS verification and a 4-digit PIN. HR manages everything from a Google Sheet — no server, no hosting costs, no database.


TABLE OF CONTENTS
-----------------

1. What You Get
2. Architecture
3. Prerequisites
4. Setup — Step by Step
   - Step 1: Create the Google Sheet
   - Step 2: Add the Apps Script Code
   - Step 3: Run Initial Setup
   - Step 4: Deploy the Web App
   - Step 5: Connect the Frontend
   - Step 6: Host the Frontend
   - Step 7: First Login Test
5. Daily Use
6. Configuration Reference
7. File Structure
8. Troubleshooting
9. Security & Privacy Notes


WHAT YOU GET
------------

- GPS-verified punches — clock in/out only works within 100 m accuracy
- PIN protection — 4-digit PIN per employee (SHA-256 hashed over the wire)
- One punch per day — prevents duplicate clock-ins or clock-outs
- Auto-calculated status — On Time, Late, Left Early, Overtime
- Live dashboard — today's snapshot, monthly performance, daily log
- Leave tracking — MC, Emergency, Annual, Unpaid
- Public holidays — excluded from working-day calculations
- Per-employee tabs — auto-generated personal attendance sheets
- Auto-refresh — dashboard updates within ~5 seconds of any punch


ARCHITECTURE
------------

Employee phone (index.html)  →  HTTPS/JSONP  →  Google Apps Script (Code.gs + Dashboard.gs)  →  Google Sheet

- Frontend: a single static index.html hosted anywhere (GitHub Pages, Netlify, Cloudflare Pages)
- Backend: Google Apps Script bound to the Google Sheet
- Storage: the Google Sheet itself
- Auth: employee name + 4-digit PIN (hashed)


PREREQUISITES
-------------

- A Google account (free or Workspace)
- A GitHub account (for hosting the frontend — or Netlify/Cloudflare/Vercel)
- Basic familiarity with copy-paste (no coding needed)
- A logo image (logo.png) — optional but recommended


SETUP — STEP BY STEP
--------------------

Step 1: Create the Google Sheet
-------------------------------

1. Go to sheets.new to create a blank spreadsheet.
2. Rename it to "Goolee Attendance" (or anything you like).
3. Create the following tabs (bottom of the screen, right-click → rename, or click +):
   - Dashboard    (auto-generated overview, leave empty)
   - Employees    (your staff list)
   - Attendance   (raw punch log, auto-written by the app)
   - Leave        (MC & leave records)
   - Holidays     (public & company holidays)
   - Settings     (working hours & rules)

4. Employees tab — paste this into cell A1:

   Name	PIN	Active	Role	Joined Date
   Mus	0000	TRUE	Game Developer	2024-03-01
   Maryam	0000	TRUE	Admin & Marketing	2024-06-15

   Important: Format the PIN column as Plain text (Format → Number → Plain text) so leading zeros are preserved.

5. Attendance tab — paste this into A1:

   Timestamp	Name	Action	Location (Google Maps)	Accuracy (m)	Date	Time

6. Leave tab — paste this into A1:

   Date	Name	Type	Notes

7. Holidays tab — paste this into A1:

   Date	Holiday Name	Type

   Then add your holidays (one per row). Example for Selangor 2026:

   2026-01-01	New Year's Day	Public Holiday
   2026-02-01	Thaipusam	Public Holiday
   2026-02-02	Thaipusam Holiday (in lieu)	Public Holiday
   2026-02-17	Chinese New Year	Public Holiday
   2026-02-18	Chinese New Year Holiday	Public Holiday
   2026-03-07	Nuzul Al-Quran	Public Holiday
   2026-03-21	Hari Raya Aidilfitri	Public Holiday
   2026-03-22	Hari Raya Aidilfitri Holiday	Public Holiday
   2026-03-23	Hari Raya Aidilfitri Holiday (in lieu)	Public Holiday
   2026-05-01	Labour Day	Public Holiday
   2026-05-27	Hari Raya Haji	Public Holiday
   2026-05-31	Wesak Day	Public Holiday
   2026-06-01	Agong's Birthday	Public Holiday
   2026-06-01	Wesak Day Holiday (in lieu)	Public Holiday
   2026-06-17	Awal Muharram	Public Holiday
   2026-08-25	Prophet Muhammad's Birthday (Maulidur Rasul)	Public Holiday
   2026-08-31	Merdeka Day (National Day)	Public Holiday
   2026-09-16	Malaysia Day	Public Holiday
   2026-11-08	Deepavali	Public Holiday
   2026-11-09	Deepavali Holiday (in lieu)	Public Holiday
   2026-12-11	Sultan of Selangor's Birthday	Public Holiday
   2026-12-25	Christmas Day	Public Holiday

   The system also auto-creates a Settings sheet when you run setup.


Step 2: Add the Apps Script Code
--------------------------------

1. In your Google Sheet, go to Extensions → Apps Script.
2. You'll see a default Code.gs file. Rename it to "Code" (click the name).
3. Paste the contents of Code.gs into it.
4. Click + next to Files → Script → name it "Dashboard".
5. Paste the contents of Dashboard.gs into it.
6. Save (Ctrl+S / Cmd+S).

Your Apps Script project should now have two files:
- Code.gs
- Dashboard.gs


Step 3: Run Initial Setup
-------------------------

1. In the Apps Script editor, open the function dropdown at the top.
2. Select "setupDashboard".
3. Click ▶ Run.
4. Google will ask for authorization:
   - Click Review permissions
   - Choose your account
   - Click Advanced → Go to Goolee Attendance (unsafe)
     (this warning is normal for personal scripts)
   - Click Allow
5. Wait ~10 seconds. You should see:
   - The Dashboard tab is now filled with content
   - A Settings tab was created
   - Per-employee tabs (Mus, Maryam, …) were created
6. Reload the Google Sheet tab to see the new structure.

If it fails: open Executions (left sidebar, clock icon) to see the error, then re-run.


Step 4: Deploy the Web App
--------------------------

1. In the Apps Script editor, click Deploy → New deployment.
2. Click the gear icon next to Select type → Web app.
3. Fill in:
   - Description: Goolee Attendance API v1
   - Execute as: Me
   - Who has access: Anyone
     (This is required — the frontend is a public page and cannot authenticate with Google.)
4. Click Deploy.
5. Copy the Web app URL — it looks like:
   https://script.google.com/macros/s/AKfycb.../exec
6. Save it somewhere — you'll need it in Step 5.

Test it: paste the URL in your browser and add ?action=ping:
   https://script.google.com/macros/s/AKfycb.../exec?action=ping
You should see: {"ok":true,"message":"Goolee attendance API is live."}


Step 5: Connect the Frontend
----------------------------

1. Open index.html in a text editor.
2. Find this line (near the bottom, inside <script>):
   API_URL: "https://script.google.com/macros/s/AKfycbzkMOPy0IMjVaBbQcRBmHby5OIR28cUEonUO0nrvgFWfXcIU63pExq3CWkQIqV-6Dtb/exec",
3. Replace the URL with your own Web app URL from Step 4.
4. Save the file.

Tip: If you plan to publish the repo publicly, use a placeholder in the committed file and instruct users to edit it after cloning — your deployment URL is not a secret, but keeping it private avoids random traffic.


Step 6: Host the Frontend
-------------------------

Pick one of these options:

Option A — GitHub Pages (free, easiest)
1. Create a new GitHub repository, e.g. goolee-attendance.
2. Upload index.html and logo.png to the repo root.
3. Go to Settings → Pages.
4. Under Source, choose Deploy from a branch → main → / (root).
5. Click Save. Wait ~1 minute.
6. Your site is live at:
   https://<your-username>.github.io/goolee-attendance/

Option B — Netlify Drop (30 seconds, no account needed for first test)
1. Go to app.netlify.com/drop.
2. Drag the folder containing index.html and logo.png into the browser.
3. Netlify gives you a URL instantly.

Option C — Cloudflare Pages / Vercel
Same idea — drag-and-drop deploy, no build step.

HTTPS is required — browsers only allow GPS geolocation on secure origins.


Step 7: First Login Test
------------------------

1. Open your site on your phone (or desktop browser).
2. Grant location permission when prompted.
3. Wait for the green "Location ready" bar.
4. Select your name from the dropdown.
5. Enter your PIN (default: 0000 for Mus/Maryam, 1111 for Arif, etc.).
6. Tap Clock In.
7. You should see a green checkmark with your name, time, and location accuracy.
8. Check your Google Sheet → Attendance tab → a new row should appear.

Success!


DAILY USE
---------

For Employees
1. Open the site
2. Wait for the green bar
3. Pick name → enter PIN → tap Clock In (morning) or Clock Out (evening)

For HR
- See today: Open Dashboard tab — snapshot at top
- Refresh manually: Menu: Goolee → 🔄 Refresh Dashboard
- Add employee: Add row to Employees tab → Goolee → 🧹 Refresh Employee Cache → 🔄 Refresh Dashboard
- Remove employee: Set their Active cell to FALSE → refresh caches
- Reset PIN: Change the PIN cell in Employees (works immediately)
- Mark leave: Add row to Leave tab → Goolee → 🔄 Refresh Dashboard
- Add holiday: Add row to Holidays tab → Goolee → 🧹 Clear all caches → 🔄 Refresh Dashboard
- Enable auto-refresh: Menu: Goolee → ⏱️ Enable auto-refresh after punch (do this once)


CONFIGURATION REFERENCE
-----------------------

Settings Sheet
- Working Days: Mon,Tue,Wed,Thu,Fri (supports Mon-Fri ranges)
- Work Start Time: 10:00 (24-hour format)
- Work End Time: 19:00 (24-hour format)
- Saturday Working Day: TRUE or FALSE
- Saturday Work Start Time: 09:00 (24-hour format)
- Saturday Work End Time: 12:00 (24-hour format)
- Late After (min): 15
- Early Leave Grace (min): 15
- Company Name: Goolee

Frontend Constants (index.html)
Inside the <script> block:

const CONFIG = {
  API_URL: "...",              // your Apps Script deployment URL
  ACCURACY_MAX_M: 100,         // max GPS accuracy in metres
  GPS_TIMEOUT_MS: 10000,       // how long to wait for a GPS lock
  GPS_RETRIES: 2,              // retries on poor accuracy
  GPS_RETRY_DELAY_MS: 1000,    // delay between retries
  GPS_MAX_AGE_MS: 30000,       // allow a recent cached phone location
  PIN_MAX_TRIES: 5,            // before lockout
  PIN_LOCKOUT_MS: 5 * 60 * 1000, // 5-minute lockout
  REQUEST_TIMEOUT_MS: 10000    // fetch timeout
};

Backend Constants (Code.gs)
const ACCURACY_MAX_M      = 100;   // must match frontend
const PIN_LOCKOUT_TRIES   = 5;
const PIN_LOCKOUT_MINUTES = 5;


FILE STRUCTURE
--------------

goolee-attendance/
├── index.html          ← the frontend (single file)
├── logo.png            ← your company logo
├── README.md           ← this file
└── apps-script/
    ├── Code.gs         ← backend API + punch handler
    └── Dashboard.gs    ← dashboard painter + menu

The Google Sheet lives separately and contains:
- Dashboard — auto-generated
- Employees — you edit
- Attendance — auto-written
- Leave — you edit
- Holidays — you edit
- Settings — you edit (rarely)
- One tab per active employee — auto-generated


TROUBLESHOOTING
---------------

Frontend
- "Location permission needed" → Browser → lock icon → set Location to Allow → reload
- "Location not accurate enough" → Step outside / near a window. Wait 10 s. Tap Retry.
- "Could not reach the attendance server" → Check API_URL in index.html. Test with ?action=ping.
- "Wrong PIN" → 5 tries allowed, then 5-minute lockout. Ask HR to reset.
- "Already Clocked In" → One clock-in per day. Contact HR to correct.
- Buttons stay grey → GPS is still locating. Wait for the green bar.
- Names won't load → Deployment must be set to Who has access: Anyone.

Backend (Apps Script)
- "Missing sheet tab: X" → Create the missing tab with the exact name.
- Dashboard not updating → Run Goolee → 🔄 Refresh Dashboard manually.
- Auto-refresh not working → Run Goolee → 🔍 Diagnose auto-refresh.
- Trigger limit reached → Open Triggers (clock icon) and delete old entries.
- "refreshDashboard is not defined" → Make sure Dashboard.gs is present and saved.
- Punch returns bad_gps → Employee's phone gave poor accuracy — move outdoors.

Redeploying After Code Changes
When you edit Code.gs or Dashboard.gs:
1. Deploy → Manage deployments
2. Click the pencil on the current deployment
3. Under Version, choose New version
4. Click Deploy

The URL stays the same — no frontend change needed.


SECURITY & PRIVACY NOTES
------------------------

- PINs are hashed with SHA-256 over the wire — the raw PIN is never sent to the server.
  (The server compares against the SHA-256 of the stored PIN.)
- PINs are stored in plain text in the Employees sheet. Anyone with Sheet access can read them. This is fine for internal use but do not reuse important PINs.
- The Apps Script deployment is public (Anyone). This is required for the frontend to reach it. Security relies on:
  - Employee name + PIN match
  - GPS accuracy ≤ 100 m
  - One punch per day per employee
  - 5-try lockout per employee
- Location is only captured at the moment of a punch, not tracked continuously.
- Compliance: Suitable for PDPA / GDPR internal use. For stricter setups, add an origin check inside doGet().

Hardening Tips
1. Restrict by Origin — inside doGet(e), check e.parameter.origin against an allowlist.
2. Rotate PINs — change them quarterly.
3. Audit the Attendance tab — it's append-only, so it's a reliable audit trail.
4. Restrict the Sheet — only HR should have edit access; employees never touch the Sheet.


SUPPORT
-------

- Found a bug? Open an Issue.
- Want a feature? Open a Discussion.
- Security concern? Email the maintainer directly.


LICENSE
-------

MIT — free for personal and commercial use. Attribution appreciated but not required.

Built with Google Apps Script, a single HTML file, and zero servers.
