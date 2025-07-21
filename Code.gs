function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🔍 Προτιμήσεις")
    .addItem("Άνοιγμα Επιλογέα Περιοχών", "showPreferenceSidebar")
    .addToUi();
  
  // Auto-show sidebar
  showPreferenceSidebar();
}

function showPreferenceSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle("Κατάταξη Περιοχών")
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getAllLocations() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ΠΕΡΙΟΧΕΣ");
  const data = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues(); // A:B

  return data
    .filter(([loc, seats]) => {
      const trimmed = String(loc || "").trim();
      return trimmed && trimmed !== "-" && Number(seats) > 0;
    })
    .map(([loc]) => String(loc).trim());
}

function submitOrderedList(locations) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MAIN");
  const row = SpreadsheetApp.getActiveRange().getRow();
  if (row < 2) return "invalid";

  const top10 = locations.slice(0, 10);
  const formatted = top10.map((loc, i) => `${i + 1}. ${loc}`).join(", ");

  sheet.getRange(row, 7).setValue(formatted); // ✅ Column G for the ranked list

  // ✅ Columns H–Q (8–17) for raw top 10
  sheet.getRange(row, 8, 1, 10).clearContent();
  if (top10.length > 0) {
    sheet.getRange(row, 8, 1, top10.length).setValues([top10]);
  }

  return "ok";
}

function runSeatAllocation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const main = ss.getSheetByName("MAIN");
  const pool = ss.getSheetByName("ΠΕΡΙΟΧΕΣ");

  // STEP 1: Load seat pool from ΠΕΡΙΟΧΕΣ
  const poolData = pool.getRange(1, 1, pool.getLastRow(), 2).getValues(); // A:B
  const seatMap = new Map();

  poolData.forEach(([loc, seats]) => {
    const trimmedLoc = String(loc || "").trim();
    const seatCount = Number(seats);
    if (trimmedLoc && !isNaN(seatCount)) {
      seatMap.set(trimmedLoc, seatCount);
    }
  });

  // STEP 2: Load user data from MAIN (columns E–Q)
  const startRow = 2;
  const numRows = main.getLastRow() - 1;
  const data = main.getRange(startRow, 5, numRows, 13).getValues(); // E–Q

  const users = data.map((row, i) => ({
    sheetRow: startRow + i,
    moria: Number(row[0]) || 0, // Column E
    preferences: [...new Set(row.slice(3, 13).map(v => String(v || "").trim()).filter(Boolean))], // H–Q
    assigned: ""
  }));

  // STEP 3: Sort users by ΜΟΡΙΑ descending
  users.sort((a, b) => b.moria - a.moria);

  // STEP 4: Allocate locations
  for (const user of users) {
    for (const pref of user.preferences) {
      if (!seatMap.has(pref)) continue;
      const available = seatMap.get(pref);
      if (available > 0) {
        seatMap.set(pref, available - 1);
        user.assigned = pref;
        Logger.log(`✅ Assigned "${pref}" to row ${user.sheetRow} (ΜΟΡΙΑ: ${user.moria})`);
        break;
      }
    }
    if (!user.assigned) {
      Logger.log(`❌ Unassigned (row ${user.sheetRow}, ΜΟΡΙΑ: ${user.moria})`);
    }
  }

  // STEP 5: Write results to MAIN (columns R and S)
  const resultRange = main.getRange(startRow, 18, numRows, 2); // R:S
  const resultValues = users.map(u => [
    u.assigned || "",
    u.assigned ? "Assigned" : "Unassigned"
  ]);
  resultRange.setValues(resultValues);

  // STEP 6: Update seats in ΠΕΡΙΟΧΕΣ column B
  poolData.forEach(([loc], i) => {
    const trimmedLoc = String(loc || "").trim();
    const remaining = seatMap.get(trimmedLoc);
    Logger.log(`↪ ${trimmedLoc}: ${remaining}`);
    pool.getRange(i + 1, 2).setValue(remaining ?? "");
  });

  Logger.log("🎯 Seat allocation complete.");
}
