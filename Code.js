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

// Attendance Data column numbers (1-based, matching sheet columns A-T)
const COL = {
  ROW_ID:         1,  // A
  SHIFT_DATE:     2,  // B
  DRIVER_ID:      3,  // C
  DRIVER_NAME:    4,  // D
  HELPER_ID:      5,  // E
  HELPER_NAME:    6,  // F
  HELPER_COMPANY: 7,  // G
  VEHICLE:        8,  // H
  START_ODO:      9,  // I
  START_PHOTO:    10, // J
  FUEL:           11, // K
  ARRIVAL:        12, // L
  DEPARTURE:      13, // M
  LAST_DROP:      14, // N
  END_TIME:       15, // O
  END_ODO:        16, // P
  END_PHOTO:      17, // Q
  FAILED_DROPS:   18, // R
  SHIFT_DURATION: 19, // S
  OVERTIME:       20  // T
};

const HEADERS = [
  'Row ID', 'Shift Date', 'Driver Employee ID', 'Driver Name',
  'Helper Employee ID', 'Helper Name', 'Helper Company', 'Vehicle Number',
  'Start Odometer (km)', 'Start Odometer Photo URL', 'Fuel Taken',
  'Arrival at Warehouse', 'Departure from Warehouse',
  'Last Drop Date & Time', 'End of Shift Date & Time',
  'End Odometer (km)', 'End Odometer Photo URL', 'Number of Failed Drops',
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
    .setTitle('Driver Attendance')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================================================================
// PWA SUPPORT
// =============================================================================

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

  var manifest = {
    name:             'RSA Driver Attendance',
    short_name:       'RSA Attend',
    description:      'Driver shift tracking and attendance management for RSA Transport',
    start_url:        deployUrl || '/',
    display:          'standalone',
    orientation:      'portrait-primary',
    theme_color:      '#0D47A1',
    background_color: '#0D47A1',
    icons: [
      { src: buildIconDataUri_(), sizes: '192x192', type: 'image/svg+xml', purpose: 'any' },
      { src: buildIconDataUri_(), sizes: '512x512', type: 'image/svg+xml', purpose: 'maskable' }
    ]
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

  var values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();

  var drivers    = [];
  var helpers    = [];
  var vehicles   = [];
  var vehicleSet = {};

  values.forEach(function(row) {
    var driverId   = String(row[0]).trim();
    var driverName = String(row[1]).trim();
    var helperId   = String(row[4]).trim();
    var helperName = String(row[5]).trim();
    var helperCo   = String(row[6]).trim();
    var vehicle    = String(row[9]).trim();

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
  });

  return { drivers: drivers, helpers: helpers, vehicles: vehicles };
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

    // Guard: check if this driver already has an incomplete shift on the same date
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var allData = sheet.getRange(2, 1, lastRow - 1, COL.OVERTIME).getValues();
      for (var i = 0; i < allData.length; i++) {
        var r              = allData[i];
        var existingDriver = String(r[COL.DRIVER_ID - 1]).trim();
        var existingDate   = String(r[COL.SHIFT_DATE - 1]).trim();
        var existingEnd    = String(r[COL.END_TIME - 1]).trim();
        if (existingDriver === data.driverId && existingDate === shiftDate && !existingEnd) {
          return { success: false, error: 'This driver already has an active shift on ' + shiftDate + '. End that shift first.' };
        }
      }
    }

    var rowId    = generateRowId_(data.driverId);
    var photoUrl = saveOdometerPhoto_(data.startPhotoBase64, rowId + '_start.jpg');

    // Build 20-element row (A–T)
    var row = new Array(20).fill('');
    row[COL.ROW_ID - 1]         = rowId;
    row[COL.SHIFT_DATE - 1]     = shiftDate;
    row[COL.DRIVER_ID - 1]      = data.driverId;
    row[COL.DRIVER_NAME - 1]    = data.driverName;
    row[COL.HELPER_ID - 1]      = data.helperId || '';
    row[COL.HELPER_NAME - 1]    = data.helperName || '';
    row[COL.HELPER_COMPANY - 1] = data.helperCompany || '';
    row[COL.VEHICLE - 1]        = data.vehicleNumber;
    row[COL.START_ODO - 1]      = Number(data.startOdometer);
    row[COL.START_PHOTO - 1]    = photoUrl;
    row[COL.FUEL - 1]           = data.fuelTaken || '';
    row[COL.ARRIVAL - 1]        = arrivalStr;

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
// END OF SHIFT
// =============================================================================

function getActiveDriversForEndShift() {
  try {
    var sheet   = getOrCreateAttendanceSheet_();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var values = sheet.getRange(2, 1, lastRow - 1, COL.OVERTIME).getValues();
    var result = [];

    values.forEach(function(row) {
      var arrival   = String(row[COL.ARRIVAL - 1]).trim();
      var departure = String(row[COL.DEPARTURE - 1]).trim();
      var endTime   = String(row[COL.END_TIME - 1]).trim();

      // Include all drivers who started a shift (arrival done) but haven't ended it yet.
      // Stage 2 (departure) is not required — if skipped it will be auto-filled on saveShiftEnd.
      if (arrival && !endTime) {
        result.push({
          rowId:          String(row[COL.ROW_ID - 1]).trim(),
          driverId:       String(row[COL.DRIVER_ID - 1]).trim(),
          driverName:     String(row[COL.DRIVER_NAME - 1]).trim(),
          vehicleNumber:  String(row[COL.VEHICLE - 1]).trim(),
          departureTime:  departure,
          hasDeparture:   departure !== ''
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

    // Compute shift duration from arrival time (col L) to manually entered end time
    var arrivalStr  = String(sheet.getRange(rowNum, COL.ARRIVAL).getValue()).trim();
    var arrivalDate = parseDubaiDate_(arrivalStr);
    var endDate     = parseDubaiDate_(endTimeStr);

    var shiftDuration = 0;
    var overtime      = 0;
    if (arrivalDate && endDate) {
      shiftDuration = Math.round(((endDate - arrivalDate) / 3600000) * 100) / 100;
      overtime      = Math.round(Math.max(0, shiftDuration - OVERTIME_HOURS) * 100) / 100;
    }

    var lastDropFormatted = formatDatetimeLocalInput_(data.lastDropTime);

    // Batch write columns N–T (14 to 20)
    sheet.getRange(rowNum, COL.LAST_DROP, 1, 7).setValues([[
      lastDropFormatted,
      endTimeStr,
      Number(data.endOdometer),
      photoUrl,
      Number(data.failedDrops),
      shiftDuration,
      overtime
    ]]);

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

    values.forEach(function(row) {
      var rowId       = String(row[COL.ROW_ID - 1]).trim();
      var shiftDate   = String(row[COL.SHIFT_DATE - 1]).trim();
      var driverId    = String(row[COL.DRIVER_ID - 1]).trim();
      var driverName  = String(row[COL.DRIVER_NAME - 1]).trim();
      var vehicle     = String(row[COL.VEHICLE - 1]).trim();
      var departure   = String(row[COL.DEPARTURE - 1]).trim();
      var endTime     = String(row[COL.END_TIME - 1]).trim();
      var shiftDur    = Number(row[COL.SHIFT_DURATION - 1]) || 0;
      var overtime    = Number(row[COL.OVERTIME - 1]) || 0;
      var failedDrops = Number(row[COL.FAILED_DROPS - 1]) || 0;
      var helperCo    = String(row[COL.HELPER_COMPANY - 1]).trim();
      var arrivalStr  = String(row[COL.ARRIVAL - 1]).trim();

      if (!driverId) return;

      // Active drivers: departed but not ended
      if (departure && !endTime) {
        var departureDate = parseDubaiDate_(departure);
        var runningHours  = departureDate
          ? Math.round(((now - departureDate) / 3600000) * 100) / 100
          : 0;

        activeDrivers.push({
          rowId:         rowId,
          driverId:      driverId,
          driverName:    driverName,
          vehicleNumber: vehicle,
          departureTime: departure,
          runningHours:  runningHours,
          isOvertime:    runningHours > OVERTIME_HOURS
        });
      }

      // Punch-out misses: departed, no end, NOT today
      if (departure && !endTime && shiftDate !== today) {
        punchOutMisses.push({
          shiftDate:     shiftDate,
          driverId:      driverId,
          driverName:    driverName,
          vehicleNumber: vehicle
        });
      }

      // Completed shifts
      if (endTime && shiftDur > 0) {
        completedShifts.push({ shiftDate: shiftDate, shiftDur: shiftDur, overtime: overtime, failedDrops: failedDrops });

        if (vehicle) {
          vehicleHoursMap[vehicle] = (vehicleHoursMap[vehicle] || 0) + shiftDur;
        }

        // Trend (last 30 days)
        var arrivalDateObj = parseDubaiDate_(arrivalStr);
        if (arrivalDateObj && arrivalDateObj >= thirtyDaysAgo) {
          if (!trendMap[shiftDate]) {
            trendMap[shiftDate] = { count: 0, totalDuration: 0, totalOvertime: 0 };
          }
          trendMap[shiftDate].count++;
          trendMap[shiftDate].totalDuration += shiftDur;
          trendMap[shiftDate].totalOvertime  += overtime;
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

    return {
      activeCount:      activeDrivers.length,
      activeDrivers:    activeDrivers.sort(function(a, b) { return a.driverName.localeCompare(b.driverName); }),
      shiftTrendByDate: shiftTrendByDate,
      vehicleRunTime:   vehicleRunTime,
      punchOutMisses:   punchOutMisses,
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
    activeCount:      0,
    activeDrivers:    [],
    shiftTrendByDate: [],
    vehicleRunTime:   [],
    punchOutMisses:   [],
    todayStats: {
      avgShiftDuration:     0,
      totalFailedDrops:     0,
      topHelperCompanies:   [],
      completedShiftsCount: 0
    }
  };
}
