// ============================================
// FLIGHT PLAN UPDATE ALERT SYSTEM
// ============================================
// This file contains all functions related to:
// - Urgent flight plan update email alerts
// - Flight update status calculations
// - Custom Google Sheets functions

// ============================================
// CONFIGURATION
// ============================================
const ALERT_CONFIG = {
  enabled: true,
  urgentKeyword: "ATNAUJINTI DABAR!!!!", // Status that triggers urgent alert
  statusColumn: "P", // Column where your formula shows the status
  emailRecipient: "matas.miltakis@heston.aero",
  maxAlertsPerCheck: 10 // Maximum flights to include in one email
};

// ============================================
// CHECK FOR URGENT UPDATES - Runs every hour
// ============================================
function checkUrgentFlightUpdates() {
  if (!ALERT_CONFIG.enabled) {
    Logger.log("Urgent flight alerts are disabled");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  let urgentFlights = [];

  // Check all sheets (skip template and old sheets)
  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    // Skip template and old sheets
    if (sheetName === CONFIG.templateSheetName || sheetName.includes('_old_')) {
      return;
    }

    // Get data from the sheet
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return; // No data

    try {
      // Get status column (column P)
      const statusCol = columnLetterToIndex(ALERT_CONFIG.statusColumn);
      const statusRange = sheet.getRange(2, statusCol, lastRow - 1, 1).getValues();
      const flightData = sheet.getRange(2, 1, lastRow - 1, 7).getValues(); // A-G

      // Find urgent flights
      for (let i = 0; i < statusRange.length; i++) {
        const status = statusRange[i][0];

        if (status && status.toString().includes(ALERT_CONFIG.urgentKeyword)) {
          const row = flightData[i];

          // Only add if we have valid data
          if (row[0] && row[1]) { // Check date and code exist
            urgentFlights.push({
              sheet: sheetName,
              date: row[0],        // Column A - LegDate
              code: row[1],        // Column B - Code
              registration: row[2], // Column C - VehicleReg
              departure: row[3],   // Column D - DepString
              arrival: row[4],     // Column E - ArrString
              std: row[5],         // Column F - STDHHMM
              sta: row[6]          // Column G - STAHHMM
            });
          }
        }
      }
    } catch (error) {
      Logger.log(`Error checking sheet ${sheetName}: ${error.toString()}`);
    }
  });

  // Send alert if urgent flights found
  if (urgentFlights.length > 0) {
    Logger.log(`Found ${urgentFlights.length} urgent flight(s) needing update`);
    sendUrgentUpdateAlert(urgentFlights);
  } else {
    Logger.log("No urgent flight updates needed at this time");
  }
}

// ============================================
// SEND URGENT UPDATE ALERT EMAIL
// ============================================
function sendUrgentUpdateAlert(flights) {
  // Limit alerts to prevent overwhelming email
  if (flights.length > ALERT_CONFIG.maxAlertsPerCheck) {
    Logger.log(`Limiting alert to ${ALERT_CONFIG.maxAlertsPerCheck} flights (found ${flights.length})`);
    flights = flights.slice(0, ALERT_CONFIG.maxAlertsPerCheck);
  }

  const now = new Date();
  const timeStr = Utilities.formatDate(now, 'UTC', 'HH:mm');

  // Build email subject
  const subject = `🚨 URGENT: ${flights.length} Flight Plan Update(s) Required NOW`;

  // Build email body
  let body = `⚠️ URGENT: ${flights.length} flight(s) need IMMEDIATE flight plan update\n`;
  body += `(Within 3 hours of departure - must update at STD-4 hours)\n\n`;
  body += `Current time: ${timeStr} UTC\n\n`;
  body += `═══════════════════════════════════════════════════\n\n`;

  flights.forEach((flight, index) => {
    body += `${index + 1}. Flight ${formatValue(flight.code)}\n`;
    body += `   📅 Date: ${formatValue(flight.date)}\n`;
    body += `   ✈️  Registration: ${formatValue(flight.registration)}\n`;
    body += `   🛫 Route: ${formatValue(flight.departure)} → ${formatValue(flight.arrival)}\n`;
    body += `   🕐 STD: ${formatTimeValue(flight.std)} UTC\n`;
    body += `   🕐 STA: ${formatTimeValue(flight.sta)} UTC\n`;
    body += `   ⚠️  ACTION: UPDATE FLIGHT PLAN NOW!\n\n`;
  });

  body += `═══════════════════════════════════════════════════\n\n`;
  body += `📊 View spreadsheet:\n${SpreadsheetApp.getActiveSpreadsheet().getUrl()}\n\n`;
  body += `ℹ️ This is an automated alert from your Flight Schedule system.\n`;
  body += `Flight plans must be updated 4 hours before STD (Scheduled Time of Departure).\n\n`;

  if (flights.length === ALERT_CONFIG.maxAlertsPerCheck) {
    body += `⚠️ Note: This email shows the first ${ALERT_CONFIG.maxAlertsPerCheck} urgent flights.\n`;
    body += `There may be more flights requiring updates. Check the spreadsheet.\n`;
  }

  try {
    GmailApp.sendEmail(ALERT_CONFIG.emailRecipient, subject, body);
    Logger.log(`✅ Urgent alert email sent to: ${ALERT_CONFIG.emailRecipient}`);
  } catch (error) {
    Logger.log(`❌ Failed to send urgent alert email: ${error.toString()}`);
  }
}

// ============================================
// SETUP HOURLY URGENT UPDATE ALERTS
// ============================================
function setupUrgentUpdateAlerts() {
  // Delete existing urgent update triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkUrgentFlightUpdates') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new hourly trigger
  ScriptApp.newTrigger('checkUrgentFlightUpdates')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("✅ Hourly urgent update alert trigger created");

  // Send confirmation email
  const recipient = ALERT_CONFIG.emailRecipient;
  const subject = "✅ Flight Plan Update Alerts Activated";
  const body = `Your urgent flight plan update alert system is now active!\n\n` +
    `⏰ Check frequency: Every hour\n` +
    `🚨 Alert trigger: "${ALERT_CONFIG.urgentKeyword}"\n` +
    `📧 Notifications sent to: ${recipient}\n` +
    `📊 Status column monitored: Column ${ALERT_CONFIG.statusColumn}\n` +
    `⚠️  Max alerts per email: ${ALERT_CONFIG.maxAlertsPerCheck} flights\n\n` +
    `─────────────────────────────────────────\n\n` +
    `How it works:\n` +
    `• System checks all flight sheets every hour\n` +
    `• When status shows "${ALERT_CONFIG.urgentKeyword}"\n` +
    `• You receive an email with flight details\n` +
    `• Alert means: Update flight plan NOW (within 3h of departure)\n\n` +
    `Your formula continues to calculate update times based on:\n` +
    `• Flight plans must be updated STD-4 hours\n` +
    `• Different update windows based on departure time\n` +
    `• All times in UTC timezone\n\n` +
    `─────────────────────────────────────────\n\n` +
    `To disable alerts: Set ALERT_CONFIG.enabled = false\n` +
    `To change email: Update ALERT_CONFIG.emailRecipient\n` +
    `To change column: Update ALERT_CONFIG.statusColumn\n\n` +
    `Spreadsheet: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;

  try {
    GmailApp.sendEmail(recipient, subject, body);
    Logger.log(`✅ Setup confirmation email sent to: ${recipient}`);
  } catch (error) {
    Logger.log(`⚠️ Could not send confirmation email: ${error.toString()}`);
  }
}

// ============================================
// MANUAL TEST - Run this to test alerts
// ============================================
function testUrgentFlightAlerts() {
  Logger.log("═══════════════════════════════════════");
  Logger.log("Testing urgent flight update alerts...");
  Logger.log("═══════════════════════════════════════");
  checkUrgentFlightUpdates();
  Logger.log("═══════════════════════════════════════");
  Logger.log("Test complete. Check logs above for results.");
  Logger.log("If flights were found, an email was sent to: " + ALERT_CONFIG.emailRecipient);
  Logger.log("═══════════════════════════════════════");
}

// ============================================
// CUSTOM FUNCTION: Calculate Flight Update Status
// ============================================
/**
 * Calculates when a flight plan needs to be updated
 * Handles overnight flights correctly
 *
 * @param {string|Date} flightDate - Flight date (e.g., "17-Nov-25")
 * @param {number|Date} stdTime - Scheduled departure time
 * @param {string|Date} todayDate - Current date
 * @param {number|Date} currentTime - Current time
 * @return {string} Update status
 * @customfunction
 */
function FLIGHT_UPDATE_STATUS(flightDate, stdTime, todayDate, currentTime) {
  try {
    // Handle empty or invalid inputs
    if (!flightDate || !stdTime) return ":)";

    // Convert dates to Date objects
    const fDate = new Date(flightDate);
    const tDate = new Date(todayDate);

    // Calculate days difference
    const daysDiff = Math.floor((fDate - tDate) / (1000 * 60 * 60 * 24));

    // Convert times to hours (handle both time formats)
    let stdHours = 0;
    let currentHours = 0;

    if (typeof stdTime === 'number') {
      stdHours = stdTime * 24; // Excel time format
    } else if (stdTime instanceof Date) {
      stdHours = stdTime.getHours() + stdTime.getMinutes() / 60;
    }

    if (typeof currentTime === 'number') {
      currentHours = currentTime * 24;
    } else if (currentTime instanceof Date) {
      currentHours = currentTime.getHours() + currentTime.getMinutes() / 60;
    }

    // Calculate total hours until departure
    const hoursUntil = (daysDiff * 24) + (stdHours - currentHours);

    // URGENT: Less than 3 hours
    if (hoursUntil < 3 && hoursUntil >= 0) {
      return "ATNAUJINTI DABAR!!!!";
    }

    // TOO FAR: More than 24 hours or not today
    if (hoursUntil > 24 || daysDiff > 0) {
      return "TOLI";
    }

    // Determine update window based on STD time
    let updateHour;
    if (stdHours >= 7.167 && stdHours < 13.167) { // 07:10-13:10
      updateHour = 4.083; // 04:05
    } else if (stdHours >= 13.167 && stdHours < 19.167) { // 13:10-19:10
      updateHour = 10.083; // 10:05
    } else if (stdHours >= 19.167) { // 19:10-00:00
      updateHour = 16.083; // 16:05
    } else if (stdHours < 1.167) { // 00:00-01:10
      updateHour = 16.083; // 16:05
    } else { // 01:10-07:10
      updateHour = 22.083; // 22:05
    }

    // Check if we're in update window
    if (currentHours >= updateHour) {
      return "ATNAUJINTI";
    } else {
      const hoursRemaining = updateHour - currentHours;
      return "ATNAUJINTI UZ " + hoursRemaining.toFixed(1) + " VAL";
    }

  } catch (error) {
    return "ERROR: " + error.toString();
  }
}

// ============================================
// CUSTOM FUNCTION: Simple Hours Until Departure
// ============================================
/**
 * Calculates hours until flight departure (handles overnight correctly)
 *
 * @param {string|Date} flightDate - Flight date
 * @param {number|Date} stdTime - Scheduled departure time
 * @param {string|Date} todayDate - Current date
 * @param {number|Date} currentTime - Current time
 * @return {number} Hours until departure
 * @customfunction
 */
function HOURS_UNTIL_DEPARTURE(flightDate, stdTime, todayDate, currentTime) {
  try {
    const fDate = new Date(flightDate);
    const tDate = new Date(todayDate);
    const daysDiff = Math.floor((fDate - tDate) / (1000 * 60 * 60 * 24));

    let stdHours = typeof stdTime === 'number' ? stdTime * 24 : stdTime.getHours() + stdTime.getMinutes() / 60;
    let currentHours = typeof currentTime === 'number' ? currentTime * 24 : currentTime.getHours() + currentTime.getMinutes() / 60;

    return (daysDiff * 24) + (stdHours - currentHours);
  } catch (error) {
    return -1;
  }
}

// ============================================
// HELPER FUNCTIONS
// ============================================

/**
 * Convert column letter to index
 * @param {string} letter - Column letter (e.g., "A", "P", "AA")
 * @return {number} Column index
 */
function columnLetterToIndex(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column = column * 26 + letter.charCodeAt(i) - 64;
  }
  return column;
}

/**
 * Format value for display
 * @param {any} value - Value to format
 * @return {string} Formatted value
 */
function formatValue(value) {
  if (!value || value === '') return 'N/A';
  return value.toString().trim();
}

/**
 * Format time value for display
 * @param {any} timeValue - Time value to format
 * @return {string} Formatted time (HH:MM)
 */
function formatTimeValue(timeValue) {
  if (!timeValue) return 'N/A';

  try {
    // If it's already a string in HH:MM format, return it
    if (typeof timeValue === 'string' && timeValue.match(/^\d{1,2}:\d{2}/)) {
      return timeValue;
    }

    // If it's a Date object
    if (timeValue instanceof Date) {
      const hours = String(timeValue.getUTCHours()).padStart(2, '0');
      const minutes = String(timeValue.getUTCMinutes()).padStart(2, '0');
      return `${hours}:${minutes}`;
    }

    // If it's a number (Excel time format)
    if (typeof timeValue === 'number') {
      const totalMinutes = Math.round(timeValue * 24 * 60);
      const hours = Math.floor(totalMinutes / 60) % 24;
      const minutes = totalMinutes % 60;
      return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
    }

    return timeValue.toString();
  } catch (error) {
    return timeValue ? timeValue.toString() : 'N/A';
  }
}
