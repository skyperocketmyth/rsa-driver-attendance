# RSA Driver Attendance — Google Apps Script Web App

A mobile-first web app for driver shift tracking, built with Google Apps Script (GAS).
Drivers log shift start, warehouse departure, and shift end. All data is written to a Google Sheet.
A master dashboard provides real-time operational oversight.

---

## Features

- **Stage 1 — Arrived at Warehouse**: Driver registers shift details, captures start odometer photo via camera, selects/enters helper info.
- **Stage 2 — Leaving Warehouse**: Driver records departure time.
- **End of Shift**: Driver logs last-drop time, shift complete time, end odometer photo, and failed drops.
- **Master Dashboard** (`?view=dashboard`): Live stats, active driver table, overtime alerts, shift trend charts, vehicle run-time, punch-out miss tracking.
- **Installable as a mobile app** (PWA) — works on iOS and Android home screens.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Google Apps Script (V8 runtime) |
| Frontend | Vanilla HTML/CSS/JS (single-page app) |
| Data store | Google Sheets |
| Photo storage | Google Drive |
| Charts | Chart.js 4.x |

---

## Setup (one-time)

1. **Create a Google Drive folder** for odometer photos. Copy the folder ID from its URL.
2. Open `Code.js` and paste the folder ID into `DRIVE_FOLDER_ID`.
3. Open `appsscript.json` — confirm the spreadsheet ID in `Code.js` matches yours.
4. **Deploy as a web app**:
   - In Apps Script editor: *Deploy → New deployment → Web app*
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Copy the deployment URL.

| URL | Purpose |
|---|---|
| `https://script.google.com/macros/s/{ID}/exec` | Driver app |
| `https://script.google.com/macros/s/{ID}/exec?view=dashboard` | Operations dashboard |

---

## Sheet Structure

The app reads driver/helper/vehicle dropdowns from a **"Dropdown List"** sheet and writes attendance records to an **"Attendance Data"** sheet (auto-created on first submission).

### Dropdown List columns (read only)
| Col | Data |
|---|---|
| A | Driver Employee ID |
| B | Driver Name |
| E | Helper Employee ID |
| F | Helper Name |
| G | Helper Company |
| J | Vehicle Number |

### Attendance Data columns (written by app)
`Row ID | Shift Date | Driver ID | Driver Name | Helper ID | Helper Name | Helper Company | Vehicle | Start Odo | Start Photo | Fuel | Arrival | Departure | Last Drop | End Time | End Odo | End Photo | Failed Drops | Shift Duration | Overtime`

---

## Install on Mobile

### Android (Chrome)
1. Open the driver app URL in Chrome.
2. Tap the browser menu → **Add to Home screen**.

### iOS (Safari)
1. Open the driver app URL in Safari.
2. Tap the Share icon → **Add to Home Screen**.

---

## Pushing updates

```bash
clasp push          # Push code to Apps Script
clasp deploy        # Create a new deployment version
```

---

## Spreadsheet
`https://docs.google.com/spreadsheets/d/14JFtpxmJt5mEMSnaCBJk7zDGp24HpR7TnSQn2LynA4o`
