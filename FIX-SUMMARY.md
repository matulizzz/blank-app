# Flight Schedule Import Script - Fix Summary

## Problem Identified

The **Code** (flight code) and **VehicleReg** (aircraft registration) columns were getting switched during import. This happened because the header matching logic was too generic and couldn't reliably distinguish between similar column names in the source data.

## Root Cause

The original `parseScheduleData()` function used fuzzy matching:
```javascript
colIndices[key] = headers.findIndex(h =>
  h.toLowerCase().includes(key.toLowerCase()) ||
  h.toLowerCase().replace(/[^a-z]/g, '') === key.toLowerCase().replace(/[^a-z]/g, '')
);
```

**Problems:**
- If the source data had headers like "Registration" and "Code", the script couldn't reliably map them
- The matching logic didn't account for common variations in header names
- No explicit alias mapping for different naming conventions

## Solution Implemented

### 1. **Added Header Alias System**
Created a comprehensive `headerAliases` configuration that maps all common variations of column names to the correct field:

```javascript
headerAliases: {
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
  // ... and more
}
```

### 2. **Improved Header Matching Logic**
The new `parseScheduleData()` function:
- Normalizes both source headers and aliases (removes special chars, lowercase)
- Uses exact matching after normalization
- Tries all aliases for each field before giving up
- Logs the complete mapping so you can verify correctness

### 3. **Enhanced Debug Logging**
Added comprehensive logging at every step:

```
Found headers in source data: ["LegDate","Registration","Flight Code",...]
Matched "Registration" (column 1) → VehicleReg via alias "Registration"
Matched "Flight Code" (column 2) → Code via alias "Flight Code"
Final column mapping:
  Source "Registration" (index 1) → VehicleReg → Sheet Column B
  Source "Flight Code" (index 2) → Code → Sheet Column C
```

### 4. **Added Debug Function**
New function `debugCheckLastEmail()` to inspect the actual headers in your emails:
```javascript
function debugCheckLastEmail() {
  // Shows actual headers from the most recent email
}
```

## How to Use the Fixed Version

### Step 1: Replace Your Script
1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Replace the entire code with the contents of `flight-schedule-import-FIXED.js`
4. Save (Ctrl+S or Cmd+S)

### Step 2: Test with Debug Function
Before processing emails, check what headers your emails actually have:

1. In Apps Script editor, select function: **debugCheckLastEmail**
2. Click **Run** ▶️
3. Check the **Execution log** to see:
   - Actual headers in your email/CSV
   - First data row sample

### Step 3: Update Header Aliases (if needed)
If `debugCheckLastEmail()` shows headers not in the alias list:

1. Add them to `CONFIG.headerAliases`
2. Example: If your email has "A/C Registration", add:
   ```javascript
   "A/C Registration": "VehicleReg",
   ```

### Step 4: Test Import
1. Select function: **testImport**
2. Click **Run** ▶️
3. Check the **Execution log** for detailed mapping info
4. Verify the imported sheet has correct data in correct columns

## Verification Checklist

After running an import, verify:

- [ ] Column A = Flight Date (LegDate)
- [ ] Column B = Flight Code (like "BA123", "LH456") ← Code from CSV
- [ ] Column C = Aircraft Registration (like "LYAAA", "LYBBB", "9HMMM", "7OMMM") ← VehicleRegistration from CSV
- [ ] Column D = Departure Airport
- [ ] Column E = Arrival Airport
- [ ] Column F = Departure Time
- [ ] Column G = Arrival Time

## Common Header Variations Already Supported

### Flight Code (Column B):
- Code, Flight Code, Flight, Flight Number, Flight No

### Aircraft Registration (Column C):
- VehicleReg, VehicleRegistration, Vehicle Reg, Vehicle Registration
- Registration, Reg, Aircraft, AircraftReg, Aircraft Reg, AC Reg
- Tail, Tail Number

### Date (Column A):
- LegDate, Leg Date, Date, Flight Date

### Departure (Column D):
- DepString, Dep String, Departure, Dep, From, Origin

### Arrival (Column E):
- ArrString, Arr String, Arrival, Arr, To, Destination

### Times (Columns F & G):
- STDHHMM, STD HHMM, STD, Dep Time, Departure Time
- STAHHMM, STA HHMM, STA, Arr Time, Arrival Time

## Troubleshooting

### If columns are still switched:

1. **Run debugCheckLastEmail()** to see actual headers
2. **Check the import logs** in Execution history:
   - Look for "Final column mapping" section
   - Verify each source header maps to correct field
3. **Add missing aliases** if headers aren't recognized
4. **Check source data order** - the script maps by name, not position

### Reading the Logs

Key log messages to look for:
```
✅ Good: "Matched 'Registration' (column 1) → VehicleReg via alias 'Registration'"
❌ Bad: "WARNING: Could not find column for VehicleReg"

✅ Good: "Row 2 data BEFORE writing: VehicleReg='G-ABCD', Code='BA123'"
❌ Bad: "Row 2 data BEFORE writing: VehicleReg='BA123', Code='G-ABCD'"  (SWITCHED!)
```

### If data is switched in the source:

If your email/CSV actually has Registration in column 3 and Code in column 2:
1. The script will still map correctly by name (not position)
2. Check the "Final column mapping" log to confirm
3. The script writes to sheet based on field name, not source position

## Technical Notes

### Sorting
- **Column B (Code)** is used for sorting - flights are sorted by flight code A-Z
- This sorts flights alphabetically by flight number (BA123, LH456, etc.)

### Column Mapping Flow
```
Source CSV/Email Headers (e.g., "Code", "VehicleRegistration")
  ↓ (normalize + match aliases)
Flight Data Object { VehicleReg: "LYAAA", Code: "BA123", ... }
  ↓ (use CONFIG.columnMapping)
Sheet Columns { B: "BA123", C: "LYAAA", ... }
```

**Corrected Mapping:**
- Code (BA123, LH456) → Column B
- VehicleReg (LYAAA, 9HMMM) → Column C

The fix ensures step 1 (matching) is accurate, so the rest flows correctly.

## Questions to Answer

Please check your source data and answer these:

1. **What are the actual column headers in your email/CSV?**
   - Run `debugCheckLastEmail()` to find out

2. **Which column should contain flight code?**
   - Examples: "BA123", "LH456", "AA789"
   - This should go to Column B

3. **Which column should contain aircraft registration?**
   - Examples: "LYAAA", "LYBBB", "9HMMM", "7OMMM"
   - This should go to Column C

4. **Are they currently being switched?**
   - If Column B shows "LYAAA" instead of "BA123", they're switched
   - The fix should resolve this

## Contact

If the issue persists after applying this fix, please provide:
1. Output from `debugCheckLastEmail()`
2. Output from `testImport()` execution log (particularly the "Final column mapping" section)
3. Screenshot showing which columns have wrong data

This will help identify if there's a unique header name that needs to be added to the alias list.
