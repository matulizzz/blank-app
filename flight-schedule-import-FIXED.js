// ============================================
// FLIGHT SCHEDULE AUTO-IMPORT SCRIPT - FIXED VERSION
// ============================================

// CONFIGURATION - UPDATE THESE VALUES
const CONFIG = {
  templateSheetName: "template", // Name of your template sheet
  gmailLabel: "FlightSchedule", // Gmail label to monitor
  emailSubjectKeyword: "schedule", // Keyword in email subject (case-insensitive)
  processedLabel: "FlightSchedule/Processed", // Label for processed emails

  // Email notification settings
  errorNotificationEmail: "matas.miltakis@heston.aero", // Leave empty to use your own email, or specify: "manager@company.com"
  sendSuccessEmail: false, // Set to true if you want confirmation emails for successful imports
  setupConfirmationEmail: "matas.miltakis@heston.aero", // Email for setup confirmation (leave empty to use errorNotificationEmail or your email)

  // Change detection settings
  enableChangeDetection: true, // Highlight changes when revised schedule arrives
  changeHighlightColor: "#fff2cc", // Yellow background for changed rows
  newFlightColor: "#d9ead3", // Green background for new flights
  removedFlightColor: "#f4cccc", // Red background for removed flights (shown in comparison)

  // Cleanup settings
  autoDeleteOldSheets: true, // Automatically delete "_old_" sheets after X days
  oldSheetRetentionDays: 5, // Delete "_old_" sheets after 5 days
  regularSheetRetentionDays: 90, // Delete regular sheets after 90 days
  sendCleanupNotification: true, // Send email before deleting sheets

  // Smart trigger settings (managed automatically, don't edit)
  expectedEmailTimeUTC: 17, // Expected email arrival hour (17 = 5 PM UTC)
  frequentCheckStartHour: 18, // Start frequent checks at 17:00 UTC (5 PM)
  frequentCheckEndHour: 23, // End frequent checks at 22:00 UTC (10 PM)
  frequentCheckIntervalMinutes: 5, // Check every 5 minutes during frequent window
  normalCheckIntervalMinutes: 30, // Check every 30 minutes during rest of day

  // Column mapping from email data to sheet
  // KEY = field name, VALUE = target column in sheet
  columnMapping: {
    LegDate: "A",        // Flight date
    VehicleReg: "B",     // Aircraft registration (e.g., "G-ABCD", "N12345")
    Code: "C",           // Flight code (e.g., "BA123", "LH456")
    DepString: "D",      // Departure airport
    ArrString: "E",      // Arrival airport
    STDHHMM: "F",        // Scheduled departure time
    STAHHMM: "G",        // Scheduled arrival time
    // Columns H-O contain formulas and will be preserved from template
  },

  // Header aliases - map source column names to our field names
  // Add any variations of header names you see in your emails here
  headerAliases: {
    // LegDate variations
    "LegDate": "LegDate",
    "Leg Date": "LegDate",
    "Date": "LegDate",
    "Flight Date": "LegDate",

    // VehicleReg variations (Aircraft Registration)
    "VehicleReg": "VehicleReg",
    "Vehicle Reg": "VehicleReg",
    "Registration": "VehicleReg",
    "Reg": "VehicleReg",
    "Aircraft": "VehicleReg",
    "AC Reg": "VehicleReg",
    "Tail": "VehicleReg",
    "Tail Number": "VehicleReg",

    // Code variations (Flight Code)
    "Code": "Code",
    "Flight Code": "Code",
    "Flight": "Code",
    "Flight Number": "Code",
    "Flight No": "Code",

    // DepString variations
    "DepString": "DepString",
    "Dep String": "DepString",
    "Departure": "DepString",
    "Dep": "DepString",
    "From": "DepString",
    "Origin": "DepString",

    // ArrString variations
    "ArrString": "ArrString",
    "Arr String": "ArrString",
    "Arrival": "ArrString",
    "Arr": "ArrString",
    "To": "ArrString",
    "Destination": "ArrString",

    // STDHHMM variations
    "STDHHMM": "STDHHMM",
    "STD HHMM": "STDHHMM",
    "STD": "STDHHMM",
    "Dep Time": "STDHHMM",
    "Departure Time": "STDHHMM",

    // STAHHMM variations
    "STAHHMM": "STAHHMM",
    "STA HHMM": "STAHHMM",
    "STA": "STAHHMM",
    "Arr Time": "STAHHMM",
    "Arrival Time": "STAHHMM"
  }
};

// ============================================
// MAIN FUNCTION - Run this daily
// ============================================
function processFlightScheduleEmails() {
  try {
    // Check authorization status first
    checkAuthorization();

    Logger.log("Starting flight schedule import...");

    // Get or create Gmail labels
    const label = getOrCreateLabel(CONFIG.gmailLabel);
    const processedLabel = getOrCreateLabel(CONFIG.processedLabel);

    // Find unprocessed emails
    const threads = GmailApp.search(`label:${CONFIG.gmailLabel} -label:${CONFIG.processedLabel}`);

    if (threads.length === 0) {
      Logger.log("No new schedule emails found.");
      return;
    }

    Logger.log(`Found ${threads.length} unprocessed email(s)`);

    // Process each email thread
    threads.forEach(thread => {
      const messages = thread.getMessages();
      messages.forEach(message => {
        processScheduleEmail(message);
      });

      // Mark as processed
      thread.addLabel(processedLabel);
      thread.removeLabel(label);
    });

    Logger.log("Flight schedule import completed successfully!");

  } catch (error) {
    Logger.log("ERROR: " + error.toString());

    // Check if it's an authorization error
    if (error.toString().includes("Authorization") ||
        error.toString().includes("permission") ||
        error.toString().includes("Access not granted")) {
      sendAuthorizationError();
    } else {
      sendErrorNotification(error);
    }
  }
}

// ============================================
// CHECK AUTHORIZATION STATUS
// ============================================
function checkAuthorization() {
  try {
    // Try to access Gmail (lightweight check)
    GmailApp.getAliases();

    // Try to access Spreadsheet
    SpreadsheetApp.getActiveSpreadsheet();

    return true;
  } catch (error) {
    throw new Error("Authorization check failed: " + error.toString());
  }
}

// ============================================
// SEND AUTHORIZATION ERROR NOTIFICATION
// ============================================
function sendAuthorizationError() {
  const recipient = CONFIG.errorNotificationEmail || Session.getActiveUser().getEmail();

  const subject = "⚠️ URGENT: Flight Schedule Script Needs Re-Authorization";
  const body = `The Flight Schedule Import script has lost authorization and needs to be re-authorized immediately.

❗ ACTION REQUIRED:
1. Open your TESTAVIMAS spreadsheet
2. Go to Extensions → Apps Script
3. Select "testImport" from the dropdown
4. Click Run ▶️
5. Follow the authorization prompts

Until re-authorized, flight schedules will NOT be imported automatically.

Technical details:
- Script: Flight Schedule Importer
- Spreadsheet: ${SpreadsheetApp.getActiveSpreadsheet().getName()}
- Time: ${new Date().toLocaleString()}

If you need help, contact your IT administrator.`;

  try {
    GmailApp.sendEmail(recipient, subject, body);
    Logger.log(`Authorization error notification sent to: ${recipient}`);
  } catch (emailError) {
    Logger.log(`Failed to send authorization error email: ${emailError.toString()}`);
  }
}

// ============================================
// PROCESS INDIVIDUAL EMAIL
// ============================================
function processScheduleEmail(message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const subject = message.getSubject();

  Logger.log(`Processing email: ${subject}`);

  // Try to extract data from attachments first
  let scheduleData = extractDataFromAttachments(message);

  // If no attachments, try to parse email body
  if (!scheduleData || scheduleData.length === 0) {
    scheduleData = extractDataFromEmailBody(message);
  }

  if (!scheduleData || scheduleData.length === 0) {
    Logger.log("WARNING: No valid schedule data found in email");
    return;
  }

  Logger.log(`Extracted ${scheduleData.length} flight records`);

  // Find most common date for sheet name
  const sheetDate = findMostCommonDate(scheduleData);
  const newSheetName = formatSheetDate(sheetDate);

  Logger.log(`Most common date: ${sheetDate}`);
  Logger.log(`Creating new sheet: ${newSheetName}`);

  try {
    // Create new sheet from template
    const newSheet = duplicateTemplateSheet(ss, newSheetName);

    if (!newSheet) {
      throw new Error("Failed to create new sheet");
    }

    Logger.log(`New sheet created successfully: ${newSheetName}`);

    // Check if there's an existing sheet to compare with (for change detection)
    const existingSheet = findExistingSheet(ss, newSheetName);

    // Import data (filtered by the target date) and sort
    const importedCount = importScheduleData(newSheet, scheduleData, sheetDate);

    Logger.log(`Successfully imported and sorted data in sheet: ${newSheetName}`);

    // Detect and highlight changes AFTER sorting (so comparison is accurate)
    if (CONFIG.enableChangeDetection && existingSheet) {
      detectAndHighlightChanges(existingSheet, newSheet);
    }

    // Send success notification if enabled
    sendSuccessNotification(newSheetName, importedCount);

  } catch (error) {
    Logger.log(`ERROR creating/importing to sheet: ${error.toString()}`);
    throw error;
  }
}

// ============================================
// EXTRACT DATA FROM ATTACHMENTS
// ============================================
function extractDataFromAttachments(message) {
  const attachments = message.getAttachments();
  let allData = [];

  attachments.forEach(attachment => {
    const fileName = attachment.getName();
    const contentType = attachment.getContentType();

    Logger.log(`Processing attachment: ${fileName}`);

    // Handle CSV files
    if (fileName.endsWith('.csv') || contentType.includes('text/csv')) {
      const csvData = Utilities.parseCsv(attachment.getDataAsString());
      allData = allData.concat(parseScheduleData(csvData));
    }

    // Handle Excel files
    else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
      const excelData = parseExcelAttachment(attachment);
      allData = allData.concat(parseScheduleData(excelData));
    }
  });

  return allData;
}

// ============================================
// EXTRACT DATA FROM EMAIL BODY
// ============================================
function extractDataFromEmailBody(message) {
  const body = message.getPlainBody();

  // Try to find table-like structure in email
  const lines = body.split('\n').filter(line => line.trim() !== '');

  // Look for patterns matching your schedule format
  const dataLines = [];
  let headerFound = false;

  lines.forEach(line => {
    // Detect header row (contains column names)
    if (!headerFound && line.match(/LegDate|VehicleReg|Code|DepString|ArrString|Registration|Flight/i)) {
      headerFound = true;
      dataLines.push(line.split(/\s+|\t/));
    }
    // Detect data rows (starts with date pattern)
    else if (headerFound && line.match(/^\d{1,2}-[A-Za-z]{3}-\d{2,4}/)) {
      dataLines.push(line.split(/\s+|\t/));
    }
  });

  return parseScheduleData(dataLines);
}

// ============================================
// PARSE EXCEL ATTACHMENT
// ============================================
function parseExcelAttachment(attachment) {
  // Convert Excel to temporary Google Sheet
  const blob = attachment.copyBlob();
  const file = {
    title: 'temp_schedule',
    mimeType: MimeType.GOOGLE_SHEETS
  };

  const tempFile = Drive.Files.insert(file, blob, {convert: true});
  const tempSheet = SpreadsheetApp.openById(tempFile.id);
  const data = tempSheet.getSheets()[0].getDataRange().getValues();

  // Delete temporary file
  DriveApp.getFileById(tempFile.id).setTrashed(true);

  return data;
}

// ============================================
// PARSE SCHEDULE DATA - FIXED VERSION
// ============================================
function parseScheduleData(rawData) {
  if (!rawData || rawData.length < 2) return [];

  const headers = rawData[0].map(h => h.toString().trim());
  const scheduleData = [];

  Logger.log(`Found headers in source data: ${JSON.stringify(headers)}`);

  // Find column indices using improved matching with aliases
  const colIndices = {};
  Object.keys(CONFIG.columnMapping).forEach(fieldName => {
    colIndices[fieldName] = -1; // Default to not found

    // Try to find matching header
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      const normalizedHeader = header.toLowerCase().replace(/[^a-z0-9]/g, '');

      // Check all aliases for this field
      for (const [alias, targetField] of Object.entries(CONFIG.headerAliases)) {
        if (targetField === fieldName) {
          const normalizedAlias = alias.toLowerCase().replace(/[^a-z0-9]/g, '');

          if (normalizedHeader === normalizedAlias) {
            colIndices[fieldName] = i;
            Logger.log(`Matched "${header}" (column ${i}) → ${fieldName} via alias "${alias}"`);
            break;
          }
        }
      }

      if (colIndices[fieldName] >= 0) break; // Found a match, stop searching
    }

    if (colIndices[fieldName] < 0) {
      Logger.log(`WARNING: Could not find column for ${fieldName}`);
    }
  });

  // Show final mapping
  Logger.log("Final column mapping:");
  Object.keys(colIndices).forEach(fieldName => {
    const sourceIndex = colIndices[fieldName];
    const targetCol = CONFIG.columnMapping[fieldName];
    const sourceHeader = sourceIndex >= 0 ? headers[sourceIndex] : "NOT FOUND";
    Logger.log(`  Source "${sourceHeader}" (index ${sourceIndex}) → ${fieldName} → Sheet Column ${targetCol}`);
  });

  // Parse data rows
  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row || row.length === 0 || !row[0]) continue;

    const flightData = {};
    Object.keys(colIndices).forEach(fieldName => {
      const sourceIndex = colIndices[fieldName];
      flightData[fieldName] = sourceIndex >= 0 ? row[sourceIndex] : '';
    });

    // Validate row has essential data
    if (flightData.LegDate || flightData.Code) {
      scheduleData.push(flightData);
    }
  }

  Logger.log(`Parsed ${scheduleData.length} valid flight records`);

  // Show first record as sample
  if (scheduleData.length > 0) {
    Logger.log("Sample first record: " + JSON.stringify(scheduleData[0]));
  }

  return scheduleData;
}

// ============================================
// FIND MOST COMMON DATE
// ============================================
function findMostCommonDate(scheduleData) {
  const dateCounts = {};

  scheduleData.forEach(flight => {
    const dateStr = flight.LegDate ? flight.LegDate.toString().trim() : '';
    if (dateStr) {
      dateCounts[dateStr] = (dateCounts[dateStr] || 0) + 1;
    }
  });

  // Find date with highest count
  let mostCommonDate = '';
  let maxCount = 0;

  Object.keys(dateCounts).forEach(date => {
    if (dateCounts[date] > maxCount) {
      maxCount = dateCounts[date];
      mostCommonDate = date;
    }
  });

  return mostCommonDate || new Date().toLocaleDateString();
}

// ============================================
// FORMAT SHEET DATE
// ============================================
function formatSheetDate(dateStr) {
  try {
    // Parse various date formats
    let date;

    // Format: "29-Mar-25" or "30-Mar-25"
    if (dateStr.match(/^\d{1,2}-[A-Za-z]{3}-\d{2,4}/)) {
      date = new Date(dateStr);
    }
    // Format: "1-Apr-25"
    else if (dateStr.match(/^\d{1,2}-[A-Za-z]{3}-\d{2}/)) {
      date = new Date(dateStr);
    }
    else {
      date = new Date(dateStr);
    }

    // Format as DDMMM (like "29SEP")
    const monthNames = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
                        'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
    const day = String(date.getDate()).padStart(2, '0');
    const month = monthNames[date.getMonth()];

    return `${day}${month}`;

  } catch (error) {
    Logger.log(`Error formatting date: ${error}`);
    const now = new Date();
    const monthNames = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN',
                        'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
    return `${String(now.getDate()).padStart(2, '0')}${monthNames[now.getMonth()]}`;
  }
}

// ============================================
// DUPLICATE TEMPLATE SHEET
// ============================================
function duplicateTemplateSheet(spreadsheet, newName) {
  const template = spreadsheet.getSheetByName(CONFIG.templateSheetName);

  if (!template) {
    throw new Error(`Template sheet "${CONFIG.templateSheetName}" not found!`);
  }

  Logger.log(`Found template sheet: ${CONFIG.templateSheetName}`);

  // Check if sheet with this name already exists
  let existingSheet = spreadsheet.getSheetByName(newName);
  if (existingSheet) {
    // Create timestamp in HHMM format (UTC)
    const now = new Date();
    const hours = String(now.getUTCHours()).padStart(2, '0');
    const minutes = String(now.getUTCMinutes()).padStart(2, '0');
    const timeStamp = `${hours}${minutes}`;

    const oldName = `${newName}_old_${timeStamp}`;
    existingSheet.setName(oldName);

    // Set a custom property to track when this old sheet was created
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty(`oldSheet_${oldName}`, now.toISOString());

    Logger.log(`Renamed existing sheet to: ${oldName} (revision received at ${hours}:${minutes} UTC)`);
  }

  // Duplicate template
  const newSheet = template.copyTo(spreadsheet);
  newSheet.setName(newName);
  Logger.log(`Created new sheet: ${newName}`);

  // Set a custom property to track when this sheet was created
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(`sheet_${newName}`, new Date().toISOString());

  // IMPORTANT: Show the new sheet (in case template is hidden)
  if (newSheet.isSheetHidden()) {
    newSheet.showSheet();
    Logger.log(`New sheet unhidden and now visible`);
  }

  // Activate the new sheet
  spreadsheet.setActiveSheet(newSheet);

  // Move to the rightmost position (end of all sheets)
  const totalSheets = spreadsheet.getSheets().length;
  spreadsheet.moveActiveSheet(totalSheets);

  Logger.log(`Moved sheet ${newName} to rightmost position`);

  return newSheet;
}

// ============================================
// FIND EXISTING SHEET (for change detection)
// ============================================
function findExistingSheet(spreadsheet, newSheetName) {
  // Look for the most recent "_old_" version of this sheet
  const sheets = spreadsheet.getSheets();
  const oldSheets = sheets.filter(sheet =>
    sheet.getName().startsWith(`${newSheetName}_old_`)
  );

  if (oldSheets.length === 0) {
    return null;
  }

  // Sort by timestamp in name (most recent first)
  oldSheets.sort((a, b) => {
    const timestampA = parseInt(a.getName().split('_old_')[1]) || 0;
    const timestampB = parseInt(b.getName().split('_old_')[1]) || 0;
    return timestampB - timestampA;
  });

  return oldSheets[0]; // Return most recent old sheet
}

// ============================================
// IMPORT SCHEDULE DATA TO SHEET
// ============================================
function importScheduleData(sheet, scheduleData, targetDate) {
  if (!scheduleData || scheduleData.length === 0) return 0;

  // Filter data to only include flights matching the target date
  const filteredData = filterDataByDate(scheduleData, targetDate);

  if (filteredData.length === 0) {
    Logger.log("WARNING: No data matches the target date after filtering");
    return 0;
  }

  Logger.log(`Filtered: ${scheduleData.length} total records -> ${filteredData.length} matching date ${targetDate}`);

  // Start from row 2 (row 1 is headers)
  let startRow = 2;

  filteredData.forEach((flight, index) => {
    const row = startRow + index;

    Logger.log(`Row ${row} data BEFORE writing: VehicleReg="${flight.VehicleReg}", Code="${flight.Code}"`);

    // Import data to columns A-G only
    Object.keys(CONFIG.columnMapping).forEach(fieldName => {
      const targetColumn = CONFIG.columnMapping[fieldName];
      const value = flight[fieldName] || '';

      // Only import to columns A-G (not H-O which have formulas)
      if (targetColumn.charCodeAt(0) <= 'G'.charCodeAt(0)) {
        sheet.getRange(targetColumn + row).setValue(value);

        // Extra logging for VehicleReg and Code
        if (fieldName === 'VehicleReg' || fieldName === 'Code') {
          Logger.log(`  Writing ${fieldName}="${value}" to column ${targetColumn} (row ${row})`);
        }
      }
    });
  });

  Logger.log(`Imported ${filteredData.length} rows to sheet`);

  // Sort by column B (Code) A-Z after import
  // IMPORTANT: Sort BEFORE change detection so we compare sorted data
  sortDataByColumnB(sheet, startRow, filteredData.length);

  return filteredData.length; // Return count for later use
}

// ============================================
// FILTER DATA BY TARGET DATE
// ============================================
function filterDataByDate(scheduleData, targetDate) {
  // Normalize target date for comparison
  const normalizedTarget = normalizeDate(targetDate);

  Logger.log(`Filtering data for target date: ${targetDate} (normalized: ${normalizedTarget})`);

  // Filter to only include flights matching the target date
  const filtered = scheduleData.filter(flight => {
    const flightDate = flight.LegDate ? flight.LegDate.toString().trim() : '';

    if (!flightDate) {
      return false; // Skip rows with no date
    }

    const normalizedFlight = normalizeDate(flightDate);
    const matches = normalizedFlight === normalizedTarget;

    if (!matches) {
      Logger.log(`Excluding flight with date: ${flightDate} (doesn't match ${targetDate})`);
    }

    return matches;
  });

  return filtered;
}

// ============================================
// NORMALIZE DATE FOR COMPARISON
// ============================================
function normalizeDate(dateStr) {
  try {
    // Parse date string to Date object
    let date;

    // Format: "29-Mar-25" or "30-Mar-25"
    if (dateStr.match(/^\d{1,2}-[A-Za-z]{3}-\d{2,4}/)) {
      date = new Date(dateStr);
    }
    // Format: "1-Apr-25"
    else if (dateStr.match(/^\d{1,2}-[A-Za-z]{3}-\d{2}/)) {
      date = new Date(dateStr);
    }
    else {
      date = new Date(dateStr);
    }

    // Return in consistent format: YYYY-MM-DD
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');

    return `${year}-${month}-${day}`;

  } catch (error) {
    Logger.log(`Error normalizing date ${dateStr}: ${error}`);
    return dateStr; // Return original if parsing fails
  }
}

// ============================================
// SORT DATA BY COLUMN B (A-Z)
// ============================================
function sortDataByColumnB(sheet, startRow, numRows) {
  if (numRows === 0) return;

  try {
    Logger.log(`Sorting ${numRows} rows by column B (VehicleReg)...`);

    // Define the range to sort (columns A-M only, starting from row 2)
    // This includes data (A-G) and formulas (H-M) so they stay together
    // Columns N-O will remain unsorted
    const lastRow = startRow + numRows - 1;
    const sortRange = sheet.getRange(startRow, 1, numRows, 13); // A-M is 13 columns

    // Sort by column B (which is column 2), ascending (A-Z)
    sortRange.sort({column: 2, ascending: true});

    Logger.log(`Sorted ${numRows} rows by column B (VehicleReg) A-Z (columns A-M only)`);

    // After sorting, trim ALL empty rows in the entire sheet (including template empty rows)
    trimEmptyRows(sheet);

  } catch (error) {
    Logger.log(`Error sorting data: ${error.toString()}`);
    // Don't fail the whole import if sorting fails
  }
}

// ============================================
// TRIM EMPTY ROWS (optimized - only checks A-G for data)
// ============================================
function trimEmptyRows(sheet) {
  try {
    const maxRows = sheet.getMaxRows();

    if (maxRows < 2) {
      return; // Nothing to trim
    }

    // Get all data from columns A-G (the data columns we care about)
    const dataRange = sheet.getRange(2, 1, maxRows - 1, 7); // A-G only
    const data = dataRange.getValues();

    // Find the last row that has ANY data in columns A-G
    let lastDataRow = 1; // Start at 1 (header row)

    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      // Check if this row has any non-empty cell in columns A-G
      const hasData = row.some(cell => {
        return cell !== '' && cell !== null && cell !== undefined;
      });

      if (hasData) {
        lastDataRow = i + 2; // +2 because data starts at row 2
        break;
      }
    }

    // Calculate how many rows to delete
    const rowsToDelete = maxRows - lastDataRow;

    if (rowsToDelete > 0) {
      // Delete all empty rows at once (single operation - very fast)
      sheet.deleteRows(lastDataRow + 1, rowsToDelete);
    }

  } catch (error) {
    // Fail silently
  }
}

// ============================================
// DETECT AND HIGHLIGHT CHANGES
// ============================================
function detectAndHighlightChanges(oldSheet, newSheet) {
  if (!CONFIG.enableChangeDetection) return;

  try {
    Logger.log("Detecting changes between old and new schedules (both already sorted)...");

    // Get data from both sheets (columns A-G, starting row 2)
    // Both sheets are already sorted by column B, so direct comparison is accurate
    const oldLastRow = oldSheet.getLastRow();
    const newLastRow = newSheet.getLastRow();

    if (oldLastRow < 2 || newLastRow < 2) {
      Logger.log("One or both sheets are empty - skipping change detection");
      return;
    }

    const oldData = oldSheet.getRange(2, 1, oldLastRow - 1, 7).getValues();
    const newData = newSheet.getRange(2, 1, newLastRow - 1, 7).getValues();

    // Create lookup maps using Code (column B, index 1) as key
    // Store ALL occurrences since there might be duplicate codes
    const oldFlights = {};
    oldData.forEach((row, index) => {
      const code = row[1] ? row[1].toString().trim() : ''; // Column B (VehicleReg)
      if (code) {
        // Create a unique key combining code and other fields for accurate matching
        const rowKey = createRowKey(row);
        oldFlights[rowKey] = { row: row, rowIndex: index + 2, code: code };
      }
    });

    const newFlights = {};
    newData.forEach((row, index) => {
      const code = row[1] ? row[1].toString().trim() : ''; // Column B (VehicleReg)
      if (code) {
        const rowKey = createRowKey(row);
        newFlights[rowKey] = { row: row, rowIndex: index + 2, code: code };
      }
    });

    let newCount = 0;
    let changedCount = 0;
    let removedCount = 0;

    // Check for new flights (in new but not in old)
    Object.keys(newFlights).forEach(rowKey => {
      const newFlight = newFlights[rowKey];

      if (!oldFlights[rowKey]) {
        // Check if this is truly new or just modified
        const sameCodeInOld = Object.values(oldFlights).find(f => f.code === newFlight.code);

        if (!sameCodeInOld) {
          // Completely new flight code
          highlightRow(newSheet, newFlight.rowIndex, CONFIG.newFlightColor);
          newCount++;
        } else {
          // Same code exists but data changed
          highlightRow(newSheet, newFlight.rowIndex, CONFIG.changeHighlightColor);
          changedCount++;
        }
      }
      // If exact match exists in oldFlights, no highlighting needed
    });

    // Check for removed flights (in old but not in new)
    Object.keys(oldFlights).forEach(rowKey => {
      if (!newFlights[rowKey]) {
        const oldFlight = oldFlights[rowKey];
        const sameCodeInNew = Object.values(newFlights).find(f => f.code === oldFlight.code);

        if (!sameCodeInNew) {
          // Completely removed
          removedCount++;
        }
        // If same code exists in new but with different data, already counted as changed
      }
    });

    // Add note to new sheet if changes detected
    if (newCount > 0 || changedCount > 0 || removedCount > 0) {
      const note = `Changes from previous version:\n✅ New flights: ${newCount}\n📝 Modified flights: ${changedCount}\n❌ Removed flights: ${removedCount}`;
      newSheet.getRange("A1").setNote(note);
      Logger.log(note);

      // Send change notification
      sendChangeNotification(newSheet.getName(), newCount, changedCount, removedCount);
    } else {
      Logger.log("No changes detected between versions");
    }

  } catch (error) {
    Logger.log(`Error detecting changes: ${error.toString()}`);
    // Don't fail the whole import if change detection fails
  }
}

// ============================================
// CREATE ROW KEY (for accurate comparison)
// ============================================
function createRowKey(row) {
  // Create a unique key from all important fields
  // This ensures we match exact duplicates, not just same code
  return row.map((cell, index) => {
    if (index <= 6) { // Only columns A-G
      return (cell || '').toString().trim();
    }
    return '';
  }).join('|');
}

// ============================================
// HIGHLIGHT ROW
// ============================================
function highlightRow(sheet, rowIndex, color) {
  try {
    // Highlight columns A-M (13 columns)
    const range = sheet.getRange(rowIndex, 1, 1, 13);
    range.setBackground(color);
  } catch (error) {
    Logger.log(`Error highlighting row ${rowIndex}: ${error.toString()}`);
  }
}

// ============================================
// SEND CHANGE NOTIFICATION
// ============================================
function sendChangeNotification(sheetName, newCount, changedCount, removedCount) {
  const recipient = CONFIG.errorNotificationEmail || Session.getActiveUser().getEmail();

  const subject = `📊 Schedule Changes Detected: ${sheetName}`;
  const body = `A revised schedule was received for ${sheetName}. Here are the changes:

✅ New flights: ${newCount}
📝 Modified flights: ${changedCount}
❌ Removed flights: ${removedCount}

Changed rows are highlighted:
- Green background = New flights
- Yellow background = Modified flights
- (Removed flights are not shown in the new schedule)

View the sheet: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}#gid=${SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getSheetId()}

The previous version has been saved as "${sheetName}_old_[timestamp]" for reference.`;

  try {
    GmailApp.sendEmail(recipient, subject, body);
    Logger.log(`Change notification sent to: ${recipient}`);
  } catch (emailError) {
    Logger.log(`Failed to send change notification: ${emailError.toString()}`);
  }
}

// ============================================
// HELPER FUNCTIONS
// ============================================
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
    Logger.log(`Created new label: ${labelName}`);
  }
  return label;
}

function sendErrorNotification(error) {
  // Determine recipient email
  const recipient = CONFIG.errorNotificationEmail || Session.getActiveUser().getEmail();

  const subject = "Flight Schedule Import Error";
  const body = `An error occurred while importing flight schedules:\n\n${error.toString()}\n\nPlease check the script logs for details.\n\nView logs: Script Editor > Executions`;

  try {
    GmailApp.sendEmail(recipient, subject, body);
    Logger.log(`Error notification sent to: ${recipient}`);
  } catch (emailError) {
    Logger.log(`Failed to send error email: ${emailError.toString()}`);
  }
}

// ============================================
// SEND SUCCESS NOTIFICATION (Optional)
// ============================================
function sendSuccessNotification(sheetName, recordCount) {
  if (!CONFIG.sendSuccessEmail) return; // Only send if enabled

  const recipient = CONFIG.errorNotificationEmail || Session.getActiveUser().getEmail();

  const subject = `Flight Schedule Imported: ${sheetName}`;
  const body = `Flight schedule successfully imported!\n\nSheet name: ${sheetName}\nFlights imported: ${recordCount}\nTimestamp: ${new Date().toLocaleString()}\n\nSpreadsheet: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;

  try {
    GmailApp.sendEmail(recipient, subject, body);
    Logger.log(`Success notification sent to: ${recipient}`);
  } catch (emailError) {
    Logger.log(`Failed to send success email: ${emailError.toString()}`);
  }
}

// ============================================
// MANUAL TEST FUNCTION
// ============================================
function testImport() {
  Logger.log("Running manual test...");
  processFlightScheduleEmails();
}

// ============================================
// SETUP FUNCTION - Run once to create trigger
// ============================================
function setupEmailTrigger() {
  // Delete existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  // Create smart triggers based on expected email time
  setupSmartTriggers();

  // Create daily cleanup trigger (runs at midnight UTC)
  ScriptApp.newTrigger('cleanupOldSheets')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();

  Logger.log("Smart triggers created for email monitoring");
  Logger.log("Daily cleanup trigger created (runs at midnight UTC)");

  // Send confirmation
  const recipient = CONFIG.setupConfirmationEmail || CONFIG.errorNotificationEmail || Session.getActiveUser().getEmail();
  GmailApp.sendEmail(
    recipient,
    "Flight Schedule Import - Trigger Active",
    "Your flight schedule import script is now active!\n\n" +
    "✅ Smart email monitoring:\n" +
    "   - Every 5 minutes between 17:00-22:00 UTC (frequent check window)\n" +
    "   - Every 1 hour during rest of the day\n" +
    "✅ Weekly health check: Every Monday at 8 AM\n" +
    "✅ Daily cleanup: Removes old sheets automatically\n" +
    "   - '_old_' sheets deleted after 5 days\n" +
    "   - Regular sheets deleted after 90 days\n\n" +
    "When a new schedule email arrives:\n" +
    "1. Apply the 'FlightSchedule' label (via your Gmail filter)\n" +
    "2. Within 5-60 minutes, the script will automatically process it\n" +
    "3. A new sheet will be created in your TESTAVIMAS spreadsheet\n" +
    "4. Changes from previous version will be highlighted\n\n" +
    "You'll receive notifications for:\n" +
    "- Schedule changes (new/modified/removed flights)\n" +
    "- Authorization issues (requires immediate action)\n" +
    "- Import errors\n" +
    "- Weekly status updates\n" +
    "- Sheet cleanup actions"
  );

  Logger.log(`Setup confirmation email sent to: ${recipient}`);
}

// ============================================
// SETUP SMART TRIGGERS
// ============================================
function setupSmartTriggers() {
  const freqStart = CONFIG.frequentCheckStartHour;
  const freqEnd = CONFIG.frequentCheckEndHour;

  Logger.log(`Setting up smart triggers with Google's 20-trigger limit`);

  // STRATEGY: Use single recurring trigger that runs every 5 minutes
  // The main function will decide internally whether to process or skip

  // Create a single trigger that runs every 5 minutes
  ScriptApp.newTrigger('processFlightScheduleEmailsSmart')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log("Created single smart trigger that runs every 5 minutes");
  Logger.log(`Will process: Every 5 min (${freqStart}:00-${freqEnd}:00 UTC), every 60 min (rest of day)`);

  // Create weekly health check trigger (runs every Monday at 8 AM)
  ScriptApp.newTrigger('weeklyHealthCheck')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  Logger.log("Created weekly health check trigger");
}

// ============================================
// SMART PROCESSING (decides when to actually process)
// ============================================
function processFlightScheduleEmailsSmart() {
  const now = new Date();
  const currentHour = now.getUTCHours();
  const currentMinute = now.getUTCMinutes();

  const freqStart = CONFIG.frequentCheckStartHour;
  const freqEnd = CONFIG.frequentCheckEndHour;

  // Check if we're in frequent check window (17:00-22:00 UTC)
  const inFrequentWindow = (currentHour >= freqStart && currentHour <= freqEnd);

  if (inFrequentWindow) {
    // During 17:00-22:00 UTC: Process every time (every 5 min)
    Logger.log(`[${currentHour}:${String(currentMinute).padStart(2, '0')} UTC] In frequent window (${freqStart}:00-${freqEnd}:00) - processing`);
    processFlightScheduleEmails();
  } else {
    // Outside frequent window: Only process every 30 minutes (at :00 and :30)
    if (currentMinute >= 0 && currentMinute <= 4) {
      // Process at top of hour (00:00, 01:00, 02:00, etc.)
      Logger.log(`[${currentHour}:${String(currentMinute).padStart(2, '0')} UTC] Outside frequent window - 30-min check (:00) - processing`);
      processFlightScheduleEmails();
    } else if (currentMinute >= 30 && currentMinute <= 34) {
      // Process at half hour (00:30, 01:30, 02:30, etc.)
      Logger.log(`[${currentHour}:${String(currentMinute).padStart(2, '0')} UTC] Outside frequent window - 30-min check (:30) - processing`);
      processFlightScheduleEmails();
    } else {
      Logger.log(`[${currentHour}:${String(currentMinute).padStart(2, '0')} UTC] Outside frequent window - skipping (next check at ${currentMinute < 30 ? currentHour + ':30' : (currentHour + 1) % 24 + ':00'})`);
      // Skip this run
    }
  }
}

// ============================================
// CLEANUP OLD SHEETS (runs daily)
// ============================================
function cleanupOldSheets() {
  if (!CONFIG.autoDeleteOldSheets) {
    Logger.log("Auto-cleanup disabled");
    return;
  }

  try {
    Logger.log("Running daily sheet cleanup...");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const now = new Date();
    const scriptProperties = PropertiesService.getScriptProperties();

    const sheetsToDelete = [];
    const sheetsDeleted = [];

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();

      // Skip template sheet
      if (sheetName === CONFIG.templateSheetName) {
        return;
      }

      // Check "_old_" sheets
      if (sheetName.includes('_old_')) {
        const createdDateStr = scriptProperties.getProperty(`oldSheet_${sheetName}`);
        if (createdDateStr) {
          const createdDate = new Date(createdDateStr);
          const daysSinceCreated = (now - createdDate) / (1000 * 60 * 60 * 24);

          if (daysSinceCreated >= CONFIG.oldSheetRetentionDays) {
            sheetsToDelete.push({ sheet: sheet, type: 'old', age: Math.floor(daysSinceCreated) });
          }
        }
      }
      // Check regular sheets (like "28SEP", "29SEP")
      else {
        const createdDateStr = scriptProperties.getProperty(`sheet_${sheetName}`);
        if (createdDateStr) {
          const createdDate = new Date(createdDateStr);
          const daysSinceCreated = (now - createdDate) / (1000 * 60 * 60 * 24);

          if (daysSinceCreated >= CONFIG.regularSheetRetentionDays) {
            sheetsToDelete.push({ sheet: sheet, type: 'regular', age: Math.floor(daysSinceCreated) });
          }
        }
      }
    });

    if (sheetsToDelete.length === 0) {
      Logger.log("No old sheets to delete");
      return;
    }

    // Send notification before deletion
    if (CONFIG.sendCleanupNotification) {
      sendCleanupNotification(sheetsToDelete);
    }

    // Delete sheets
    sheetsToDelete.forEach(item => {
      try {
        const sheetName = item.sheet.getName();
        ss.deleteSheet(item.sheet);
        sheetsDeleted.push(`${sheetName} (${item.age} days old)`);

        // Remove from properties
        if (item.type === 'old') {
          scriptProperties.deleteProperty(`oldSheet_${sheetName}`);
        } else {
          scriptProperties.deleteProperty(`sheet_${sheetName}`);
        }

        Logger.log(`Deleted sheet: ${sheetName} (${item.age} days old)`);
      } catch (error) {
        Logger.log(`Error deleting sheet ${item.sheet.getName()}: ${error.toString()}`);
      }
    });

    Logger.log(`Cleanup complete: ${sheetsDeleted.length} sheets deleted`);

  } catch (error) {
    Logger.log(`Error during cleanup: ${error.toString()}`);
  }
}

// ============================================
// SEND CLEANUP NOTIFICATION
// ============================================
function sendCleanupNotification(sheetsToDelete) {
  const recipient = CONFIG.errorNotificationEmail || Session.getActiveUser().getEmail();

  const oldSheets = sheetsToDelete.filter(s => s.type === 'old').map(s => `  - ${s.sheet.getName()} (${s.age} days old)`).join('\n');
  const regularSheets = sheetsToDelete.filter(s => s.type === 'regular').map(s => `  - ${s.sheet.getName()} (${s.age} days old)`).join('\n');

  const subject = `🗑️ Flight Schedule Cleanup: ${sheetsToDelete.length} sheets will be deleted`;
  const body = `Automatic cleanup is removing old sheets from your TESTAVIMAS spreadsheet:

${oldSheets ? `"_old_" sheets (older than ${CONFIG.oldSheetRetentionDays} days):\n${oldSheets}\n\n` : ''}${regularSheets ? `Regular sheets (older than ${CONFIG.regularSheetRetentionDays} days):\n${regularSheets}\n\n` : ''}These sheets will be permanently deleted in the next few minutes.

If you need to keep any of these sheets:
1. Open the spreadsheet immediately
2. Rename the sheet to prevent deletion (add "KEEP-" prefix)

Spreadsheet: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}

This is an automatic cleanup process. To disable or adjust:
- Set CONFIG.autoDeleteOldSheets = false (to disable)
- Adjust CONFIG.oldSheetRetentionDays (currently ${CONFIG.oldSheetRetentionDays})
- Adjust CONFIG.regularSheetRetentionDays (currently ${CONFIG.regularSheetRetentionDays})`;

  try {
    GmailApp.sendEmail(recipient, subject, body);
    Logger.log(`Cleanup notification sent to: ${recipient}`);
  } catch (emailError) {
    Logger.log(`Failed to send cleanup notification: ${emailError.toString()}`);
  }
}

// ============================================
// HIDE TEMPLATE SHEET
// ============================================
function hideTemplateSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const template = ss.getSheetByName(CONFIG.templateSheetName);

    if (template && !template.isSheetHidden()) {
      template.hideSheet();
      Logger.log(`Template sheet "${CONFIG.templateSheetName}" is now hidden`);
    } else if (template && template.isSheetHidden()) {
      Logger.log(`Template sheet "${CONFIG.templateSheetName}" is already hidden`);
    }
  } catch (error) {
    Logger.log(`Could not hide template sheet: ${error.toString()}`);
    // Don't fail setup if hiding fails
  }
}

// ============================================
// UNHIDE TEMPLATE SHEET (for maintenance)
// ============================================
function unhideTemplateSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const template = ss.getSheetByName(CONFIG.templateSheetName);

    if (template && template.isSheetHidden()) {
      template.showSheet();
      Logger.log(`Template sheet "${CONFIG.templateSheetName}" is now visible`);
    } else {
      Logger.log(`Template sheet "${CONFIG.templateSheetName}" is already visible`);
    }
  } catch (error) {
    Logger.log(`Could not unhide template sheet: ${error.toString()}`);
  }
}

// ============================================
// WEEKLY HEALTH CHECK
// ============================================
function weeklyHealthCheck() {
  Logger.log("Running weekly health check...");

  const recipient = CONFIG.errorNotificationEmail || Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // Check authorization
    checkAuthorization();

    // Check triggers are active
    const triggers = ScriptApp.getProjectTriggers();
    const mainTrigger = triggers.find(t => t.getHandlerFunction() === 'processFlightScheduleEmails');

    if (!mainTrigger) {
      throw new Error("Main trigger not found! Email monitoring is not active.");
    }

    // Count recent imports (sheets created in last 7 days)
    const sheets = ss.getSheets();
    const oneWeekAgo = new Date();
    oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);

    let recentSheets = 0;
    sheets.forEach(sheet => {
      const created = new Date(sheet.getSheetId()); // Approximation
      // Count sheets that aren't 'template' and might be recent
      if (sheet.getName() !== CONFIG.templateSheetName &&
          !sheet.getName().includes('_old_')) {
        recentSheets++;
      }
    });

    // Send weekly status email
    const subject = "✅ Flight Schedule Import - Weekly Status";
    const body = `Weekly health check completed successfully!

📊 Status: All systems operational
✅ Authorization: Valid
✅ Triggers: Active
📅 Monitoring: Running every 10 minutes

📈 Statistics:
- Active sheets in spreadsheet: ${sheets.length}
- Template sheet: ${CONFIG.templateSheetName}

🔧 Configuration:
- Gmail label: ${CONFIG.gmailLabel}
- Spreadsheet: ${ss.getName()}
- Next check: Next Monday at 8 AM

Everything is working correctly. No action required.

---
This is an automated weekly status email. To disable, set CONFIG.sendWeeklyStatus = false`;

    GmailApp.sendEmail(recipient, subject, body);
    Logger.log("Weekly health check completed successfully");

  } catch (error) {
    Logger.log("Weekly health check failed: " + error.toString());

    // Send warning email
    const subject = "⚠️ Flight Schedule Import - Health Check Failed";
    const body = `Weekly health check detected an issue:

❌ Issue: ${error.toString()}

🔧 Recommended actions:
1. Open your TESTAVIMAS spreadsheet
2. Go to Extensions → Apps Script
3. Check the Executions log for recent errors
4. Re-run authorization if needed (testImport function)

The script may not be working correctly until this is resolved.

---
Time: ${new Date().toLocaleString()}
Spreadsheet: ${ss.getUrl()}`;

    GmailApp.sendEmail(recipient, subject, body);
  }
}

// ============================================
// ALTERNATIVE: Setup daily trigger (backup option)
// ============================================
function setupDailyTrigger() {
  // Delete existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processFlightScheduleEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new daily trigger at 6:00 AM
  ScriptApp.newTrigger('processFlightScheduleEmails')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();

  Logger.log("Daily trigger created successfully! Will run every day at 6:00 AM");
}

// ============================================
// DEBUG FUNCTION - Check what headers are in the last email
// ============================================
function debugCheckLastEmail() {
  const label = GmailApp.getUserLabelByName(CONFIG.gmailLabel);
  if (!label) {
    Logger.log("Label not found");
    return;
  }

  const threads = label.getThreads(0, 1); // Get most recent thread
  if (threads.length === 0) {
    Logger.log("No emails found with this label");
    return;
  }

  const message = threads[0].getMessages()[0];
  Logger.log("Email subject: " + message.getSubject());

  // Check attachments
  const attachments = message.getAttachments();
  if (attachments.length > 0) {
    attachments.forEach(att => {
      Logger.log("\nAttachment: " + att.getName());

      if (att.getName().endsWith('.csv')) {
        const csvData = Utilities.parseCsv(att.getDataAsString());
        Logger.log("Headers found: " + JSON.stringify(csvData[0]));
        Logger.log("First data row: " + JSON.stringify(csvData[1]));
      }
    });
  } else {
    Logger.log("No attachments found");
  }
}
