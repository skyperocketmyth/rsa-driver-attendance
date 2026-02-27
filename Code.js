// =============================================================================
// DRIVER ATTENDANCE WEB APP — Google Apps Script Backend
// =============================================================================
// SETUP (one-time):
//   1. Create a Google Drive folder for odometer photos
//   2. Copy the folder ID from its URL and paste it into DRIVE_FOLDER_ID below
//   3. Deploy as web app: Execute as "Me", Who has access "Anyone"
//   4. Driver app URL: /exec
//   5. Dashboard URL:  /exec?view=dashboard
// =============================================================================

const SPREADSHEET_ID   = '14JFtpxmJt5mEMSnaCBJk7zDGp24HpR7TnSQn2LynA4o';
const DROPDOWN_SHEET   = 'Dropdown List';
const ATTENDANCE_SHEET = 'Attendance Data';
const DRIVE_FOLDER_ID  = '105-FYZkG7mBsBOP7t3AGvL8zLWyn-WqS'; // <-- Replace before deploy
const OVERTIME_HOURS   = 9;
const TIMEZONE         = 'Asia/Dubai';

// Attendance Data column numbers (1-based, matching sheet columns A-Y)
const COL = {
  ROW_ID:            1,  // A
  SHIFT_DATE:        2,  // B
  DRIVER_ID:         3,  // C
  DRIVER_NAME:       4,  // D
  HELPER_ID:         5,  // E
  HELPER_NAME:       6,  // F
  HELPER_COMPANY:    7,  // G
  VEHICLE:           8,  // H
  START_ODO:         9,  // I
  START_PHOTO:       10, // J
  FUEL:              11, // K
  DESTINATION:       12, // L  — Destination Emirate
  PRIMARY_CUSTOMER:  13, // M  — Primary Customer Name
  TOTAL_DROPS:       14, // N  — Total Drops (planned)
  ARRIVAL:           15, // O  — Arrival at Gate
  DEPARTURE:         16, // P
  LAST_DROP:         17, // Q
  LAST_DROP_PHOTO:   18, // R  — Last Drop Odometer Photo URL
  LAST_DROP_SUBMIT:  19, // S  — Last Drop Submission Timestamp
  FAILED_DROPS:      20, // T
  END_TIME:          21, // U
  END_ODO:           22, // V
  END_PHOTO:         23, // W
  SHIFT_DURATION:    24, // X
  OVERTIME:          25  // Y
};

const HEADERS = [
  'Row ID', 'Shift Date', 'Driver Employee ID', 'Driver Name',
  'Helper Employee ID', 'Helper Name', 'Helper Company', 'Vehicle Number',
  'Start Odometer (km)', 'Start Odometer Photo URL', 'Fuel Taken',
  'Destination Emirate', 'Primary Customer Name', 'Total Drops',
  'Arrival at Gate', 'Departure from Warehouse',
  'Last Drop Date & Time', 'Last Drop Odo Photo URL', 'Last Drop Submission Time',
  'Number of Failed Drops', 'Shift Complete Date & Time',
  'End Odometer (km)', 'End Odometer Photo URL',
  'Shift Duration (hrs)', 'Overtime Hours'
];

// =============================================================================
// ENTRY POINT
// =============================================================================

function doGet(e) {
  var param = (e && e.parameter) ? e.parameter : {};

  // Serve PWA manifest JSON when ?manifest=1 is requested
  if (param.manifest === '1') {
    return buildPwaManifest_();
  }

  // Serve the service worker JS when ?sw=1 is requested.
  // This is required for Chrome on Android to offer "Install App" (standalone mode)
  // instead of just "Add to Home Screen" (opens in browser).
  if (param.sw === '1') {
    return buildServiceWorker_();
  }

  var view = param.view || 'driver';
  var template = HtmlService.createTemplateFromFile('Index');

  try {
    var data = getInitialData();
    template.initialData = JSON.stringify(data);
  } catch (err) {
    template.initialData = JSON.stringify({ drivers: [], helpers: [], vehicles: [], error: err.message });
  }

  // Embed the real deployment exec URL so the client can construct the dashboard link correctly.
  // window.location.href inside the GAS iframe is a googleusercontent.com sandboxed URL,
  // not the exec URL — so we must pass the real URL from the server side.
  var deployUrl = '';
  try { deployUrl = ScriptApp.getService().getUrl(); } catch(ignore) {}
  template.deploymentUrl = deployUrl;

  // Build and embed the app icon as a base64 data URI (used for splash screen + apple-touch-icon)
  template.iconDataUri = buildIconDataUri_();

  template.view = view;

  return template.evaluate()
    .setTitle('RSA Driver App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================================================================
// PWA SUPPORT
// =============================================================================

function buildServiceWorker_() {
  // Minimal service worker — install/activate only.
  // No caching: GAS requires a live network for every request anyway.
  // Its sole purpose is to satisfy Chrome's PWA installability check so
  // Android users get "Install App" (standalone WebAPK) instead of a
  // browser shortcut.
  var js = [
    '// RSA Driver Attendance — Service Worker',
    'self.addEventListener("install", function(e) {',
    '  e.waitUntil(self.skipWaiting());',
    '});',
    'self.addEventListener("activate", function(e) {',
    '  e.waitUntil(clients.claim());',
    '});',
    '// Empty fetch handler satisfies Chrome PWA install criteria.',
    '// All network requests fall through to the browser default.',
    'self.addEventListener("fetch", function(e) {});'
  ].join('\n');
  return ContentService.createTextOutput(js).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// Returns a public Google Drive thumbnail URL for the app icon.
// Creates the SVG file in Drive once, then caches the thumbnail URL in
// Script Properties so subsequent manifest requests are instant.
// Uses the /thumbnail endpoint (returns a JPEG from Google CDN) which is
// more reliable than uc?export=view for SVG files — Chrome requires a
// decodable image with a proper MIME type for PWA manifest icons.
function getOrCreateIconUrl_() {
  try {
    var props  = PropertiesService.getScriptProperties();
    // Use a versioned key so changing the icon URL format forces regeneration.
    var cached = props.getProperty('PWA_ICON_URL_V2');
    if (cached) return cached;

    var svgBlob = Utilities.newBlob(buildIconSvg_(), 'image/svg+xml', 'rsa-app-icon.svg');
    var folder  = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    var file    = folder.createFile(svgBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // thumbnail endpoint: Drive rasterises the SVG to a JPEG served from
    // Google's CDN — reliable MIME type, no auth required for public files.
    var url = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w512';
    props.setProperty('PWA_ICON_URL_V2', url);
    return url;
  } catch (err) {
    // Non-fatal: manifest will have no icons if Drive is unavailable.
    return '';
  }
}

function buildIconSvg_() {
  // RSA Transport icon: dark-blue rounded square + white delivery truck (Material local_shipping)
  return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 192 192">' +
    '<rect width="192" height="192" rx="38" fill="#0D47A1"/>' +
    '<g transform="translate(16,16) scale(6.667)">' +
    '<path fill="white" d="M20 8h-3V4H3c-1.1 0-2 .9-2 2v11h2c0 1.66 1.34 3 3 3s3-1.34 3-3h6c0 ' +
    '1.66 1.34 3 3 3s3-1.34 3-3h2v-5l-3-4zm-.5 1.5l1.96 2.5H17V9.5h2.5zM6 18.5c-.83 0-1.5-.67-' +
    '1.5-1.5s.67-1.5 1.5-1.5 1.5.67 1.5 1.5-.67 1.5-1.5 1.5zm13.5-3H14V6h4.5l3 4v5.5zm-7 3c-' +
    '.83 0-1.5-.67-1.5-1.5s.67-1.5 1.5-1.5 1.5.67 1.5 1.5-.67 1.5-1.5 1.5z"/>' +
    '</g></svg>';
}

function buildIconDataUri_() {
  return 'data:image/svg+xml;base64,' + Utilities.base64Encode(buildIconSvg_());
}

function buildPwaManifest_() {
  var deployUrl = '';
  try { deployUrl = ScriptApp.getService().getUrl(); } catch(e) {}

  // Use a real Drive-hosted URL for the icon so Chrome can load it properly.
  // Data URIs are not accepted as PWA manifest icons by Chrome.
  var iconUrl = getOrCreateIconUrl_();
  // No 'type' declared — Chrome auto-detects from the Content-Type the
  // Drive CDN returns (JPEG). Declaring the wrong type would cause rejection.
  var icons = iconUrl ? [
    { src: iconUrl, sizes: '192x192', purpose: 'any' },
    { src: iconUrl, sizes: '512x512', purpose: 'maskable' }
  ] : [];

  var manifest = {
    name:             'RSA Driver App',
    short_name:       'RSA Driver',
    description:      'Driver shift tracking and attendance management for RSA Transport',
    start_url:        deployUrl || '/',
    display:          'standalone',
    orientation:      'portrait-primary',
    theme_color:      '#0D47A1',
    background_color: '#0D47A1',
    icons:            icons
  };

  return ContentService
    .createTextOutput(JSON.stringify(manifest))
    .setMimeType(ContentService.MimeType.JSON);
}

// =============================================================================
// INITIAL DATA LOAD
// =============================================================================

function getInitialData() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(DROPDOWN_SHEET);

  if (!sheet) {
    throw new Error('Sheet "' + DROPDOWN_SHEET + '" not found in the spreadsheet.');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { drivers: [], helpers: [], vehicles: [] };
  }

  // Read up to column V (22 cols) to capture destination (col S=19) and customer (col V=22)
  var values = sheet.getRange(2, 1, lastRow - 1, 22).getValues();

  var drivers      = [];
  var helpers      = [];
  var vehicles     = [];
  var destinations = [];
  var customers    = [];
  var vehicleSet   = {};
  var destSet      = {};
  var custSet      = {};

  values.forEach(function(row) {
    var driverId   = String(row[0]).trim();
    var driverName = String(row[1]).trim();
    var helperId   = String(row[4]).trim();
    var helperName = String(row[5]).trim();
    var helperCo   = String(row[6]).trim();
    var vehicle    = String(row[9]).trim();
    var dest       = String(row[18]).trim();  // Column S (0-indexed: 18)
    var customer   = String(row[21]).trim();  // Column V (0-indexed: 21)

    if (driverId && driverName) {
      drivers.push({ id: driverId, name: driverName });
    }
    if (helperId && helperName) {
      helpers.push({ id: helperId, name: helperName, company: helperCo });
    }
    if (vehicle && !vehicleSet[vehicle]) {
      vehicleSet[vehicle] = true;
      vehicles.push({ number: vehicle });
    }
    if (dest && !destSet[dest]) {
      destSet[dest] = true;
      destinations.push(dest);
    }
    if (customer && !custSet[customer]) {
      custSet[customer] = true;
      customers.push({ name: customer });
    }
  });

  return { drivers: drivers, helpers: helpers, vehicles: vehicles,
           destinations: destinations, customers: customers };
}

// =============================================================================
// PRIVATE HELPERS
// =============================================================================

function getOrCreateAttendanceSheet_() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ATTENDANCE_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(ATTENDANCE_SHEET);
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange.setBackground('#1565C0');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 220);  // Row ID column wider
  }

  return sheet;
}

function formatDubai_(date) {
  return Utilities.formatDate(date, TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
}

function formatDateOnly_(date) {
  return Utilities.formatDate(date, TIMEZONE, 'dd/MM/yyyy');
}

function generateRowId_(driverId) {
  var stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd-HHmmss');
  return 'SHIFT-' + stamp + '-' + driverId;
}

function saveOdometerPhoto_(base64Data, filename) {
  var cleaned = base64Data.replace(/^data:image\/\w+;base64,/, '');
  var bytes   = Utilities.base64Decode(cleaned);
  var blob    = Utilities.newBlob(bytes, 'image/jpeg', filename);
  var folder  = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?export=view&id=' + file.getId();
}

function parseDubaiDate_(str) {
  // Input: 'DD/MM/YYYY HH:MM:SS'
  if (!str) return null;
  var parts = str.split(' ');
  if (parts.length < 2) return null;
  var d = parts[0].split('/');
  var t = parts[1].split(':');
  if (d.length < 3 || t.length < 2) return null;
  return new Date(
    parseInt(d[2]), parseInt(d[1]) - 1, parseInt(d[0]),
    parseInt(t[0]), parseInt(t[1]), t[2] ? parseInt(t[2]) : 0
  );
}

function findRowByRowId_(sheet, rowId) {
  var finder = sheet.createTextFinder(rowId).findNext();
  if (!finder || finder.getColumn() !== 1) return null;
  return finder.getRow();
}

function formatDatetimeLocalInput_(dtLocal) {
  // Input from datetime-local: 'YYYY-MM-DDTHH:MM'
  if (!dtLocal) return '';
  try {
    var parts = dtLocal.split('T');
    var d     = parts[0].split('-');
    var time  = parts[1] || '00:00';
    return d[2] + '/' + d[1] + '/' + d[0] + ' ' + time;
  } catch(e) {
    return dtLocal;
  }
}

// =============================================================================
// SHIFT START — STAGE 1 (at warehouse)
// =============================================================================

function saveShiftStart(data) {
  try {
    var sheet = getOrCreateAttendanceSheet_();

    // Use the manually entered shift start time
    if (!data.shiftStartTime) {
      return { success: false, error: 'Shift start time is required.' };
    }
    var arrivalStr = formatDatetimeLocalInput_(data.shiftStartTime);
    var shiftDate  = arrivalStr.split(' ')[0];  // Extract 'DD/MM/YYYY' from 'DD/MM/YYYY HH:MM'

    // Guard: check if this driver has ANY incomplete shift (any date)
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var allData = sheet.getRange(2, 1, lastRow - 1, COL.OVERTIME).getValues();
      for (var i = 0; i < allData.length; i++) {
        var r              = allData[i];
        var existingDriver = String(r[COL.DRIVER_ID - 1]).trim();
        var existingEnd    = String(r[COL.END_TIME - 1]).trim();
        if (existingDriver !== data.driverId) continue;
        if (existingEnd) continue;  // shift fully completed, skip

        // Driver has an incomplete shift — determine which stage they're stuck on
        var existingDate      = String(r[COL.SHIFT_DATE - 1]).trim();
        var existingArrival   = String(r[COL.ARRIVAL - 1]).trim();
        var existingDeparture = String(r[COL.DEPARTURE - 1]).trim();
        var existingLastDrop  = String(r[COL.LAST_DROP_SUBMIT - 1]).trim();

        var stuckStage, stageLabel;
        if (existingArrival && !existingDeparture && !existingLastDrop) {
          stuckStage = 2; stageLabel = 'Stage 2 (Departure)';
        } else if (existingDeparture && !existingLastDrop) {
          stuckStage = 3; stageLabel = 'Stage 3 (Last Drop)';
        } else if (existingLastDrop) {
          stuckStage = 4; stageLabel = 'Stage 4 (Shift Complete)';
        } else {
          stuckStage = 1; stageLabel = 'Stage 1';
        }

        return {
          success: false,
          error: 'You have an incomplete shift from ' + existingDate + '. Please complete ' + stageLabel + ' before starting a new shift.',
          incompleteShiftDate:     existingDate,
          incompleteShiftStage:    stageLabel,
          incompleteShiftStageNum: stuckStage
        };
      }
    }

    var rowId    = generateRowId_(data.driverId);
    var photoUrl = saveOdometerPhoto_(data.startPhotoBase64, rowId + '_start.jpg');

    // Build 25-element row (A–Y)
    var row = new Array(25).fill('');
    row[COL.ROW_ID - 1]           = rowId;
    row[COL.SHIFT_DATE - 1]       = shiftDate;
    row[COL.DRIVER_ID - 1]        = data.driverId;
    row[COL.DRIVER_NAME - 1]      = data.driverName;
    row[COL.HELPER_ID - 1]        = data.helperId || '';
    row[COL.HELPER_NAME - 1]      = data.helperName || '';
    row[COL.HELPER_COMPANY - 1]   = data.helperCompany || '';
    row[COL.VEHICLE - 1]          = data.vehicleNumber;
    row[COL.START_ODO - 1]        = Number(data.startOdometer);
    row[COL.START_PHOTO - 1]      = photoUrl;
    row[COL.FUEL - 1]             = data.fuelTaken || '';
    row[COL.DESTINATION - 1]      = data.destinationEmirate || '';
    row[COL.PRIMARY_CUSTOMER - 1] = data.primaryCustomer || '';
    row[COL.TOTAL_DROPS - 1]      = Number(data.totalDrops) || 0;
    row[COL.ARRIVAL - 1]          = arrivalStr;

    sheet.appendRow(row);

    return { success: true, rowId: rowId, arrivalTime: arrivalStr };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// =============================================================================
// SHIFT START — STAGE 2 (departure from warehouse)
// =============================================================================

function getStage1PendingDrivers() {
  try {
    var sheet   = getOrCreateAttendanceSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var now       = new Date();
    var today     = formatDateOnly_(now);
    var yesterday = formatDateOnly_(new Date(now.getTime() - 86400000));
    var values = sheet.getRange(2, 1, lastRow - 1, COL.OVERTIME).getValues();
    var result = [];

    values.forEach(function(row) {
      var shiftDate = String(row[COL.SHIFT_DATE - 1]).trim();
      var arrival   = String(row[COL.ARRIVAL - 1]).trim();
      var departure = String(row[COL.DEPARTURE - 1]).trim();

      // Show pending departures from today and yesterday (covers overnight / early shifts)
      if ((shiftDate === today || shiftDate === yesterday) && arrival && !departure) {
        result.push({
          rowId:         String(row[COL.ROW_ID - 1]).trim(),
          driverId:      String(row[COL.DRIVER_ID - 1]).trim(),
          driverName:    String(row[COL.DRIVER_NAME - 1]).trim(),
          vehicleNumber: String(row[COL.VEHICLE - 1]).trim(),
          helperName:    String(row[COL.HELPER_NAME - 1]).trim(),
          helperCompany: String(row[COL.HELPER_COMPANY - 1]).trim(),
          arrivalTime:   arrival
        });
      }
    });

    return result.sort(function(a, b) { return a.driverName.localeCompare(b.driverName); });
  } catch (err) {
    return [];
  }
}

function saveDeparture(rowId, departureTimeStr) {
  try {
    if (!departureTimeStr) {
      return { success: false, error: 'Departure time is required.' };
    }
    var sheet  = getOrCreateAttendanceSheet_();
    var rowNum = findRowByRowId_(sheet, rowId);
    if (!rowNum) {
      return { success: false, error: 'Shift record not found. Please check with your supervisor.' };
    }

    var departureStr = formatDatetimeLocalInput_(departureTimeStr);
    sheet.getRange(rowNum, COL.DEPARTURE).setValue(departureStr);

    return { success: true, departureTime: departureStr };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// =============================================================================
// STAGE 3 — LAST DROP
// =============================================================================

// Returns drivers eligible for Stage 3: have arrival but no last-drop submission yet.
// Stage 2 (departure) is optional — if skipped it will be auto-filled on saveShiftEnd.
function getActiveDriversForEndShift() {
  try {
    var sheet   = getOrCreateAttendanceSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var values = sheet.getRange(2, 1, lastRow - 1, COL.OVERTIME).getValues();
    var result = [];

    values.forEach(function(row) {
      var arrival      = String(row[COL.ARRIVAL - 1]).trim();
      var departure    = String(row[COL.DEPARTURE - 1]).trim();
      var lastDropSub  = String(row[COL.LAST_DROP_SUBMIT - 1]).trim();

      // Stage 3 candidates: arrival done, last-drop not yet submitted
      if (arrival && !lastDropSub) {
        result.push({
          rowId:         String(row[COL.ROW_ID - 1]).trim(),
          driverId:      String(row[COL.DRIVER_ID - 1]).trim(),
          driverName:    String(row[COL.DRIVER_NAME - 1]).trim(),
          vehicleNumber: String(row[COL.VEHICLE - 1]).trim(),
          departureTime: departure,
          hasDeparture:  departure !== ''
        });
      }
    });

    return result.sort(function(a, b) { return a.driverName.localeCompare(b.driverName); });
  } catch (err) {
    return [];
  }
}

function saveLastDrop(data) {
  try {
    var sheet  = getOrCreateAttendanceSheet_();
    var rowNum = findRowByRowId_(sheet, data.rowId);
    if (!rowNum) {
      return { success: false, error: 'Shift record not found.' };
    }

    if (!data.lastDropTime) {
      return { success: false, error: 'Last drop date & time is required.' };
    }

    var photoUrl          = saveOdometerPhoto_(data.lastDropPhotoBase64, data.rowId + '_lastdrop.jpg');
    var lastDropFormatted = formatDatetimeLocalInput_(data.lastDropTime);
    var submitTime        = formatDubai_(new Date());  // auto-captured server timestamp

    sheet.getRange(rowNum, COL.LAST_DROP).setValue(lastDropFormatted);
    sheet.getRange(rowNum, COL.LAST_DROP_PHOTO).setValue(photoUrl);
    sheet.getRange(rowNum, COL.LAST_DROP_SUBMIT).setValue(submitTime);
    sheet.getRange(rowNum, COL.FAILED_DROPS).setValue(Number(data.failedDrops) || 0);

    return { success: true, submitTime: submitTime };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// =============================================================================
// STAGE 4 — SHIFT COMPLETE
// =============================================================================

// Returns drivers eligible for Stage 4: Stage 3 done but shift not yet completed.
function getStage3PendingDrivers() {
  try {
    var sheet   = getOrCreateAttendanceSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var values = sheet.getRange(2, 1, lastRow - 1, COL.OVERTIME).getValues();
    var result = [];

    values.forEach(function(row) {
      var lastDropSub = String(row[COL.LAST_DROP_SUBMIT - 1]).trim();
      var endTime     = String(row[COL.END_TIME - 1]).trim();

      if (lastDropSub && !endTime) {
        result.push({
          rowId:         String(row[COL.ROW_ID - 1]).trim(),
          driverId:      String(row[COL.DRIVER_ID - 1]).trim(),
          driverName:    String(row[COL.DRIVER_NAME - 1]).trim(),
          vehicleNumber: String(row[COL.VEHICLE - 1]).trim()
        });
      }
    });

    return result.sort(function(a, b) { return a.driverName.localeCompare(b.driverName); });
  } catch (err) {
    return [];
  }
}

function saveShiftEnd(data) {
  try {
    var sheet  = getOrCreateAttendanceSheet_();
    var rowNum = findRowByRowId_(sheet, data.rowId);
    if (!rowNum) {
      return { success: false, error: 'Shift record not found.' };
    }

    // Save end odometer photo
    var photoUrl = saveOdometerPhoto_(data.endPhotoBase64, data.rowId + '_end.jpg');

    // Use manually entered shift complete time
    if (!data.shiftCompleteTime) {
      return { success: false, error: 'Shift complete time is required.' };
    }
    var endTimeStr = formatDatetimeLocalInput_(data.shiftCompleteTime);

    // Auto-fill departure if Stage 2 was skipped (use shift complete time as fallback)
    var existingDeparture = String(sheet.getRange(rowNum, COL.DEPARTURE).getValue()).trim();
    if (!existingDeparture) {
      sheet.getRange(rowNum, COL.DEPARTURE).setValue(endTimeStr);
    }

    // Compute shift duration from arrival time to manually entered end time
    var arrivalStr  = String(sheet.getRange(rowNum, COL.ARRIVAL).getValue()).trim();
    var arrivalDate = parseDubaiDate_(arrivalStr);
    var endDate     = parseDubaiDate_(endTimeStr);

    var shiftDuration = 0;
    var overtime      = 0;
    if (arrivalDate && endDate) {
      shiftDuration = Math.round(((endDate - arrivalDate) / 3600000) * 100) / 100;
      overtime      = Math.round(Math.max(0, shiftDuration - OVERTIME_HOURS) * 100) / 100;
    }

    // Write Stage 4 fields: END_TIME, END_ODO, END_PHOTO, SHIFT_DURATION, OVERTIME
    sheet.getRange(rowNum, COL.END_TIME).setValue(endTimeStr);
    sheet.getRange(rowNum, COL.END_ODO).setValue(Number(data.endOdometer));
    sheet.getRange(rowNum, COL.END_PHOTO).setValue(photoUrl);
    sheet.getRange(rowNum, COL.SHIFT_DURATION).setValue(shiftDuration);
    sheet.getRange(rowNum, COL.OVERTIME).setValue(overtime);

    return { success: true, shiftDuration: shiftDuration, overtime: overtime };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// =============================================================================
// DASHBOARD DATA
// =============================================================================

function getDashboardData() {
  try {
    var sheet   = getOrCreateAttendanceSheet_();
    var lastRow = sheet.getLastRow();

    if (lastRow < 2) return buildEmptyDashboard_();

    var values = sheet.getRange(2, 1, lastRow - 1, COL.OVERTIME).getValues();
    var now    = new Date();
    var today  = formatDateOnly_(now);

    var thirtyDaysAgo = new Date(now.getTime() - 30 * 24 * 3600000);

    var activeDrivers   = [];
    var completedShifts = [];
    var punchOutMisses  = [];
    var vehicleHoursMap = {};
    var trendMap        = {};
    var helperCoMap     = {};
    var failedDropMap   = {};  // keyed by shiftDate
    var overtimeMap     = {};  // keyed by shiftDate

    values.forEach(function(row) {
      var rowId        = String(row[COL.ROW_ID - 1]).trim();
      var shiftDate    = String(row[COL.SHIFT_DATE - 1]).trim();
      var driverId     = String(row[COL.DRIVER_ID - 1]).trim();
      var driverName   = String(row[COL.DRIVER_NAME - 1]).trim();
      var vehicle      = String(row[COL.VEHICLE - 1]).trim();
      var arrivalStr   = String(row[COL.ARRIVAL - 1]).trim();
      var departure    = String(row[COL.DEPARTURE - 1]).trim();
      var lastDropSub  = String(row[COL.LAST_DROP_SUBMIT - 1]).trim();
      var endTime      = String(row[COL.END_TIME - 1]).trim();
      var shiftDur     = Number(row[COL.SHIFT_DURATION - 1]) || 0;
      var overtime     = Number(row[COL.OVERTIME - 1]) || 0;
      var failedDrops  = Number(row[COL.FAILED_DROPS - 1]) || 0;
      var totalDrops   = Number(row[COL.TOTAL_DROPS - 1]) || 0;
      var helperCo     = String(row[COL.HELPER_COMPANY - 1]).trim();

      if (!driverId) return;

      // Active drivers: any incomplete shift (arrival present, no end time)
      if (arrivalStr && !endTime) {
        // Determine current stage
        var currentStage;
        if (!departure) {
          currentStage = 1;  // Arrived at gate, not yet departed
        } else if (!lastDropSub) {
          currentStage = 2;  // Departed, last drop not yet submitted
        } else {
          currentStage = 3;  // Last drop done, shift not yet complete
        }

        var arrivalDateObj = parseDubaiDate_(arrivalStr);
        var runningHours   = arrivalDateObj
          ? Math.round(((now - arrivalDateObj) / 3600000) * 100) / 100
          : 0;

        activeDrivers.push({
          rowId:         rowId,
          driverId:      driverId,
          driverName:    driverName,
          vehicleNumber: vehicle,
          arrivalTime:   arrivalStr,
          departureTime: departure,
          runningHours:  runningHours,
          currentStage:  currentStage,
          isOvertime:    runningHours > OVERTIME_HOURS
        });
      }

      // Punch-out misses: any incomplete shift NOT from today
      if (arrivalStr && !endTime && shiftDate !== today) {
        var missStage;
        if (!departure) {
          missStage = 2;
        } else if (!lastDropSub) {
          missStage = 3;
        } else {
          missStage = 4;
        }
        punchOutMisses.push({
          shiftDate:     shiftDate,
          driverId:      driverId,
          driverName:    driverName,
          vehicleNumber: vehicle,
          stuckStage:    missStage
        });
      }

      // Completed shifts
      if (endTime && shiftDur > 0) {
        completedShifts.push({ shiftDate: shiftDate, shiftDur: shiftDur, overtime: overtime, failedDrops: failedDrops });

        if (vehicle) {
          vehicleHoursMap[vehicle] = (vehicleHoursMap[vehicle] || 0) + shiftDur;
        }

        // Trend (last 30 days)
        var arrivalD = parseDubaiDate_(arrivalStr);
        if (arrivalD && arrivalD >= thirtyDaysAgo) {
          if (!trendMap[shiftDate]) {
            trendMap[shiftDate] = { count: 0, totalDuration: 0, totalOvertime: 0 };
          }
          trendMap[shiftDate].count++;
          trendMap[shiftDate].totalDuration += shiftDur;
          trendMap[shiftDate].totalOvertime  += overtime;
        }

        // Failed drops by date
        if (!failedDropMap[shiftDate]) {
          failedDropMap[shiftDate] = { totalDrops: 0, failedDrops: 0, drivers: {} };
        }
        failedDropMap[shiftDate].totalDrops  += totalDrops;
        failedDropMap[shiftDate].failedDrops += failedDrops;
        if (!failedDropMap[shiftDate].drivers[driverName]) {
          failedDropMap[shiftDate].drivers[driverName] = { totalDrops: 0, failedDrops: 0 };
        }
        failedDropMap[shiftDate].drivers[driverName].totalDrops  += totalDrops;
        failedDropMap[shiftDate].drivers[driverName].failedDrops += failedDrops;

        // Overtime by date/driver (only if overtime > 0)
        if (overtime > 0) {
          if (!overtimeMap[shiftDate]) overtimeMap[shiftDate] = [];
          overtimeMap[shiftDate].push({
            driverName:    driverName,
            shiftDuration: shiftDur,
            overtime:      overtime
          });
        }
      }

      // Today's helper companies
      if (shiftDate === today && helperCo) {
        helperCoMap[helperCo] = (helperCoMap[helperCo] || 0) + 1;
      }
    });

    // Today stats
    var todayCompleted  = completedShifts.filter(function(s) { return s.shiftDate === today; });
    var avgShiftToday   = todayCompleted.length
      ? Math.round((todayCompleted.reduce(function(s, r) { return s + r.shiftDur; }, 0) / todayCompleted.length) * 100) / 100
      : 0;
    var totalFailedToday = todayCompleted.reduce(function(s, r) { return s + r.failedDrops; }, 0);

    // Shift trend (sorted by date)
    var shiftTrendByDate = Object.keys(trendMap).sort().map(function(date) {
      var d = trendMap[date];
      return {
        date:          date,
        shiftCount:    d.count,
        avgDuration:   Math.round((d.totalDuration / d.count) * 100) / 100,
        totalOvertime: Math.round(d.totalOvertime * 100) / 100
      };
    });

    // Vehicle run time (descending)
    var vehicleRunTime = Object.keys(vehicleHoursMap)
      .map(function(v) { return { vehicleNumber: v, totalHours: Math.round(vehicleHoursMap[v] * 100) / 100 }; })
      .sort(function(a, b) { return b.totalHours - a.totalHours; });

    // Top 5 helper companies
    var topHelperCompanies = Object.keys(helperCoMap)
      .map(function(c) { return { company: c, count: helperCoMap[c] }; })
      .sort(function(a, b) { return b.count - a.count; })
      .slice(0, 5);

    // Failed drops by date (sorted desc)
    var failedDropsByDate = Object.keys(failedDropMap).sort().reverse().map(function(date) {
      var d = failedDropMap[date];
      var failureRate = d.totalDrops > 0
        ? Math.round((d.failedDrops / d.totalDrops) * 1000) / 10
        : 0;
      return {
        date:        date,
        totalDrops:  d.totalDrops,
        failedDrops: d.failedDrops,
        failureRate: failureRate,
        drivers: Object.keys(d.drivers).map(function(name) {
          return {
            driverName:  name,
            totalDrops:  d.drivers[name].totalDrops,
            failedDrops: d.drivers[name].failedDrops
          };
        }).sort(function(a, b) { return b.failedDrops - a.failedDrops; })
      };
    });

    // Overtime by date/driver (sorted desc)
    var overtimeByDate = Object.keys(overtimeMap).sort().reverse().map(function(date) {
      return {
        date:    date,
        drivers: overtimeMap[date].sort(function(a, b) { return b.overtime - a.overtime; })
      };
    });

    return {
      activeCount:      activeDrivers.length,
      activeDrivers:    activeDrivers.sort(function(a, b) { return a.driverName.localeCompare(b.driverName); }),
      shiftTrendByDate: shiftTrendByDate,
      vehicleRunTime:   vehicleRunTime,
      punchOutMisses:   punchOutMisses,
      failedDropsByDate: failedDropsByDate,
      overtimeByDate:   overtimeByDate,
      todayStats: {
        avgShiftDuration:     avgShiftToday,
        totalFailedDrops:     totalFailedToday,
        topHelperCompanies:   topHelperCompanies,
        completedShiftsCount: todayCompleted.length
      }
    };
  } catch (err) {
    return { error: err.message };
  }
}

function buildEmptyDashboard_() {
  return {
    activeCount:       0,
    activeDrivers:     [],
    shiftTrendByDate:  [],
    vehicleRunTime:    [],
    punchOutMisses:    [],
    failedDropsByDate: [],
    overtimeByDate:    [],
    todayStats: {
      avgShiftDuration:     0,
      totalFailedDrops:     0,
      topHelperCompanies:   [],
      completedShiftsCount: 0
    }
  };
}

// =============================================================================
// DASHBOARD DETAIL DATA (date-specific: vehicle km + stage timings)
// =============================================================================

function getDashboardDetailData(dateStr) {
  // dateStr: 'DD/MM/YYYY'
  try {
    var sheet   = getOrCreateAttendanceSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return buildEmptyDetailData_();

    var values      = sheet.getRange(2, 1, lastRow - 1, COL.OVERTIME).getValues();
    var vehicleKmMap = {};
    var stageTimings = [];

    values.forEach(function(row) {
      var shiftDate = String(row[COL.SHIFT_DATE - 1]).trim();
      if (shiftDate !== dateStr) return;

      var driverName  = String(row[COL.DRIVER_NAME - 1]).trim();
      var vehicle     = String(row[COL.VEHICLE - 1]).trim();
      var arrivalStr  = String(row[COL.ARRIVAL - 1]).trim();
      var departure   = String(row[COL.DEPARTURE - 1]).trim();
      var lastDropSub = String(row[COL.LAST_DROP_SUBMIT - 1]).trim();
      var endTime     = String(row[COL.END_TIME - 1]).trim();
      var startOdo    = Number(row[COL.START_ODO - 1]) || 0;
      var endOdo      = Number(row[COL.END_ODO - 1])   || 0;

      // Vehicle km: completed shifts with valid odometer readings
      if (endTime && endOdo > 0 && startOdo > 0 && vehicle) {
        var km = Math.max(0, endOdo - startOdo);
        if (!vehicleKmMap[vehicle]) {
          vehicleKmMap[vehicle] = { km: 0, startOdo: startOdo, endOdo: endOdo };
        }
        vehicleKmMap[vehicle].km     += km;
        vehicleKmMap[vehicle].endOdo  = endOdo;
      }

      // Stage timings: any shift that has at least an arrival
      if (arrivalStr) {
        var arrivalDate = parseDubaiDate_(arrivalStr);

        // gap12: arrival → departure (hrs)
        var gap12     = null;
        var gap12Auto = false;
        if (departure && arrivalDate) {
          var depDate = parseDubaiDate_(departure);
          if (depDate) gap12 = Math.round(((depDate - arrivalDate) / 3600000) * 100) / 100;
          // Auto-filled: departure was set equal to endTime by saveShiftEnd
          if (endTime && departure === endTime) gap12Auto = true;
        }

        // gap23: departure → lastDropSubmit (hrs)
        var gap23 = null;
        if (departure && lastDropSub) {
          var dep2 = parseDubaiDate_(departure);
          var ld   = parseDubaiDate_(lastDropSub);
          if (dep2 && ld) gap23 = Math.round(((ld - dep2) / 3600000) * 100) / 100;
        }

        // gap34: lastDropSubmit → endTime (hrs)
        var gap34 = null;
        if (lastDropSub && endTime) {
          var ld2 = parseDubaiDate_(lastDropSub);
          var et  = parseDubaiDate_(endTime);
          if (ld2 && et) gap34 = Math.round(((et - ld2) / 3600000) * 100) / 100;
        }

        stageTimings.push({
          driverName:    driverName,
          vehicleNumber: vehicle,
          arrivalTime:   arrivalStr,
          gap12:         gap12Auto ? null : gap12,
          gap12Auto:     gap12Auto,
          gap23:         gap23,
          gap34:         gap34,
          isComplete:    endTime !== ''
        });
      }
    });

    var vehicleKmByDate = Object.keys(vehicleKmMap)
      .map(function(v) {
        return { vehicleNumber: v, km: vehicleKmMap[v].km, startOdo: vehicleKmMap[v].startOdo, endOdo: vehicleKmMap[v].endOdo };
      })
      .sort(function(a, b) { return b.km - a.km; });

    return {
      dateStr:         dateStr,
      vehicleKmByDate: vehicleKmByDate,
      stageTimings:    stageTimings.sort(function(a, b) { return a.driverName.localeCompare(b.driverName); })
    };
  } catch (err) {
    return { error: err.message };
  }
}

function buildEmptyDetailData_() {
  return { dateStr: '', vehicleKmByDate: [], stageTimings: [] };
}

// =============================================================================
// VEHICLE HOURS FOR DATE RANGE
// =============================================================================

function getVehicleHoursForRange(fromDate, toDate) {
  // fromDate, toDate: 'DD/MM/YYYY'
  try {
    var sheet   = getOrCreateAttendanceSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var values = sheet.getRange(2, 1, lastRow - 1, COL.OVERTIME).getValues();

    // Parse boundary dates for comparison
    var fromD = parseDubaiDate_(fromDate + ' 00:00:00');
    var toD   = parseDubaiDate_(toDate   + ' 23:59:59');
    if (!fromD || !toD) return { error: 'Invalid date range.' };

    var vehicleHoursMap = {};

    values.forEach(function(row) {
      var shiftDate  = String(row[COL.SHIFT_DATE - 1]).trim();
      var endTime    = String(row[COL.END_TIME - 1]).trim();
      var shiftDur   = Number(row[COL.SHIFT_DURATION - 1]) || 0;
      var vehicle    = String(row[COL.VEHICLE - 1]).trim();

      if (!endTime || shiftDur <= 0 || !vehicle) return;

      var shiftDateObj = parseDubaiDate_(shiftDate + ' 00:00:00');
      if (!shiftDateObj) return;
      if (shiftDateObj < fromD || shiftDateObj > toD) return;

      vehicleHoursMap[vehicle] = (vehicleHoursMap[vehicle] || 0) + shiftDur;
    });

    return Object.keys(vehicleHoursMap)
      .map(function(v) { return { vehicleNumber: v, totalHours: Math.round(vehicleHoursMap[v] * 100) / 100 }; })
      .sort(function(a, b) { return b.totalHours - a.totalHours; });
  } catch (err) {
    return { error: err.message };
  }
}
