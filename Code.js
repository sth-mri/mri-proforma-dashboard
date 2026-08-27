// --- NEW HEADLESS API ROUTER ---
function doPost(e) {
  try {
    // Parse the incoming request from GitHub
    let request = JSON.parse(e.postData.contents);
    let action = request.action;
    let payload = request.payload;
    let result;

    // Route the request to the correct function
    if (action === "getSettings") {
      result = getSettings();
    } else if (action === "getScannerSchedule") {
      result = getScannerSchedule(payload);
    } else if (action === "saveSlotToSheet") {
      result = saveSlotToSheet(payload);
    } else if (action === "getGlobalSchedule") {
      result = getGlobalSchedule();
    }

    // Return the data as a clean JSON package
    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: result }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    // If an error occurs, send it back to GitHub to display
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Ensure you handle CORS for cross-origin requests
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

const START_TIME_MINS = 450; 
const ROW_OFFSET = 7;
const DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
const BLOCK_MINS = 5;

// --- ROUTING ENGINE ---
function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'decoder') {
    return HtmlService.createHtmlOutputFromFile('Decoder').setTitle('MRI Proforma Decoder').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('MRI Proforma Dashboard').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getScriptUrl() { return ScriptApp.getService().getUrl(); }
function timeToMins(timeStr) {
  if (!timeStr) return 0;
  let p = timeStr.split(':');
  return parseInt(p[0]) * 60 + parseInt(p[1]);
}
function minsToTime(mins) {
  let h = Math.floor(mins / 60);
  let m = mins % 60;
  return (h < 10 ? '0'+h : h) + ':' + (m < 10 ? '0'+m : m);
}

// --- SHEET CHANGE DETECTOR (Flags Cache as Outdated) ---
function onEdit(e) {
  // If ANY manual edit happens on the sheet, flag the cache to be rebuilt!
  PropertiesService.getDocumentProperties().setProperty('scheduleLastModified', Date.now().toString());
}
function triggerCacheUpdate() {
  PropertiesService.getDocumentProperties().setProperty('scheduleLastModified', Date.now().toString());
}

function getSettings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName("Settings");
    if (!settingsSheet) return [];
    const dataRange = settingsSheet.getDataRange();
    const data = dataRange.getValues();
    const backgrounds = dataRange.getBackgrounds(); 
    let activities = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) activities.push({ name: data[i][0].toString(), color: backgrounds[i][1], note: data[i][2] ? data[i][2].toString() : "" });
    }
    return activities;
  } catch(e) { return []; }
}

function getDayMapping(sheet) {
  try {
    let lastCol = sheet.getLastColumn();
    let maxCols = sheet.getMaxColumns();
    let numCols = Math.max(7, lastCol - 1); 
    if ((numCols + 1) > maxCols) numCols = maxCols - 1; 
    if (numCols < 7) numCols = 7; 
    
    let row6Vals = sheet.getRange(6, 2, 1, numCols).getValues()[0];
    let dayMap = [];
    let dayInfo = {};
    DAYS.forEach(d => dayInfo[d] = { startIdx: -1, span: 0 });
    
    let currentDayStr = DAYS[0];
    for (let i = 0; i < numCols; i++) {
      let val = row6Vals[i] ? row6Vals[i].toString().trim() : "";
      let matchedDay = DAYS.find(d => d.toLowerCase() === val.toLowerCase());
      if (matchedDay) currentDayStr = matchedDay;
      dayMap.push(currentDayStr);
      if(dayInfo[currentDayStr].startIdx === -1) dayInfo[currentDayStr].startIdx = i;
      dayInfo[currentDayStr].span++;
    }
    return { numCols, dayMap, dayInfo };
  } catch(e) {
    let dayInfo = {};
    DAYS.forEach((d, i) => dayInfo[d] = { startIdx: i, span: 1 });
    return { numCols: 7, dayMap: DAYS, dayInfo: dayInfo };
  }
}

function getScannerSchedule(scannerName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(scannerName);
    if (!sheet) throw new Error("Scanner sheet not found");

    const headerRange = sheet.getRange("A1:A4");
    const hVals = headerRange.getValues();
    const hColors = headerRange.getFontColors();
    const headers = [
      { text: hVals[0][0], color: hColors[0][0] }, { text: hVals[1][0], color: hColors[1][0] },
      { text: hVals[2][0], color: hColors[2][0] }, { text: hVals[3][0], color: hColors[3][0] }
    ];

    let mapping = getDayMapping(sheet);
    let sheetMaxRows = sheet.getMaxRows();
    let maxRows = 156; 
    if (ROW_OFFSET + maxRows - 1 > sheetMaxRows) maxRows = sheetMaxRows - ROW_OFFSET + 1;
    if (maxRows < 1) maxRows = 1;

    const dataRange = sheet.getRange(ROW_OFFSET, 2, maxRows, mapping.numCols);
    const values = dataRange.getValues();
    const backgrounds = dataRange.getBackgrounds();
    const rtvGrid = dataRange.getRichTextValues();
    const mergedRanges = dataRange.getMergedRanges();
    
    let slots = [];
    let coveredCells = {};

    function parseFormat(run, defSize, defColor, defBold) {
      if (!run) return { size: defSize, color: defColor, bold: defBold, italic: false };
      let ts = run.getTextStyle();
      return { size: ts.getFontSize() || defSize, color: ts.getForegroundColor() || defColor, bold: ts.isBold(), italic: ts.isItalic() };
    }

    mergedRanges.forEach(range => {
      let rStart = range.getRow() - ROW_OFFSET;
      let rEnd = rStart + range.getNumRows();
      let cIdx = range.getColumn() - 2; 
      let cSpan = range.getNumColumns();
      
      if(rStart < 0 || rStart >= maxRows || cIdx < 0 || cIdx >= mapping.numCols) return; 

      let val = range.getValue().toString().trim();
      if(!val) return;
      
      let day = mapping.dayMap[cIdx];
      let colOffset = cIdx - mapping.dayInfo[day].startIdx;
      
      let leftPct = (colOffset / mapping.dayInfo[day].span) * 100;
      let widthPct = (cSpan / mapping.dayInfo[day].span) * 100;
      widthPct = Math.min(widthPct, 100 - leftPct); 
      
      let trackMode = (cSpan === mapping.dayInfo[day].span) ? "Full" : (colOffset + 1).toString();
      let parts = val.split('\n');
      let rtvRuns = null;
      try { rtvRuns = rtvGrid[rStart][cIdx].getRuns(); } catch(e) {}
      
      slots.push({
        day: day, trackMode: trackMode, leftPct: leftPct, widthPct: widthPct,
        startIndex: rStart, endIndex: rEnd,
        startTime: minsToTime(START_TIME_MINS + (rStart * BLOCK_MINS)),
        endTime: minsToTime(START_TIME_MINS + (rEnd * BLOCK_MINS)),
        activity: parts[0].trim(), note: parts.slice(1).join('\n').trim(), color: range.getBackground(),
        actFormat: parseFormat(rtvRuns ? rtvRuns[0] : null, 12, "#000000", true),
        noteFormat: parseFormat(rtvRuns && rtvRuns.length > 1 ? rtvRuns[rtvRuns.length-1] : (rtvRuns ? rtvRuns[0] : null), 12, "#ff0000", false)
      });

      for(let r = rStart; r < rEnd; r++) {
        for(let c = cIdx; c < cIdx + cSpan; c++) coveredCells[`${r}-${c}`] = true;
      }
    });

    for (let r = 0; r < maxRows; r++) {
      for (let c = 0; c < mapping.numCols; c++) {
        if (!coveredCells[`${r}-${c}`]) {
          let val = values[r][c].toString().trim();
          if (val !== "") {
            let day = mapping.dayMap[c];
            let colOffset = c - mapping.dayInfo[day].startIdx;
            let parts = val.split('\n');
            let rtvRuns = null;
            try { rtvRuns = rtvGrid[r][c].getRuns(); } catch(e) {}

            slots.push({
              day: day, trackMode: (mapping.dayInfo[day].span > 1) ? (colOffset + 1).toString() : "Full",
              leftPct: (colOffset / mapping.dayInfo[day].span) * 100,
              widthPct: (1 / mapping.dayInfo[day].span) * 100,
              startIndex: r, endIndex: r + 1,
              startTime: minsToTime(START_TIME_MINS + (r * BLOCK_MINS)),
              endTime: minsToTime(START_TIME_MINS + ((r + 1) * BLOCK_MINS)),
              activity: parts[0].trim(), note: parts.slice(1).join('\n').trim(), color: backgrounds[r][c],
              actFormat: parseFormat(rtvRuns ? rtvRuns[0] : null, 12, "#000000", true),
              noteFormat: parseFormat(rtvRuns && rtvRuns.length > 1 ? rtvRuns[rtvRuns.length-1] : (rtvRuns ? rtvRuns[0] : null), 12, "#ff0000", false)
            });
          }
        }
      }
    }
    return { headers: headers, slots: slots, dayInfo: mapping.dayInfo };
  } catch (error) { throw new Error("Backend Error: " + error.message); }
}

function saveSlotToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(data.scanner);
  let mapping = getDayMapping(sheet);
  
  let startMins = timeToMins(data.startTime);
  let endMins = timeToMins(data.endTime);
  let startRow = ((startMins - START_TIME_MINS) / BLOCK_MINS) + ROW_OFFSET;
  let numRows = (endMins - startMins) / BLOCK_MINS;
  
  let baseCol = 2 + mapping.dayInfo[data.day].startIdx;
  let targetColOffset = 0;
  let targetColSpan = mapping.dayInfo[data.day].span;
  if (data.trackMode !== "Full") { targetColOffset = parseInt(data.trackMode) - 1; targetColSpan = 1; }
  
  let col = baseCol + targetColOffset;
  const targetRange = sheet.getRange(startRow, col, numRows, targetColSpan);
  
  if (!data.forceOverwrite) {
    let existingValues = targetRange.getValues().flat();
    let hasData = existingValues.some(val => val !== "");
    let isPartOfMerge = targetRange.isPartOfMerge();
    if (hasData || isPartOfMerge) return "COLLISION";
  }

  let columnRange = sheet.getRange(ROW_OFFSET, col, 156, targetColSpan); 
  let mergedRanges = columnRange.getMergedRanges();
  let targetStartRow = startRow;
  let targetEndRow = startRow + numRows - 1;
  
  mergedRanges.forEach(range => {
    let mStart = range.getRow();
    let mEnd = mStart + range.getNumRows() - 1;
    if (targetStartRow <= mEnd && targetEndRow >= mStart) {
      range.breakApart(); range.clearContent(); range.setBackground(null); 
    }
  });

  targetRange.breakApart(); targetRange.clearContent(); targetRange.setBackground(null);
  
  let actText = (data.activity || "").trim();
  let noteText = (data.note || "").trim();
  let fullText = actText + (noteText ? "\n" + noteText : "");
  if (!fullText) fullText = " "; 

  let rtv = SpreadsheetApp.newRichTextValue().setText(fullText);
  if (actText.length > 0) {
    let actStyle = SpreadsheetApp.newTextStyle().setForegroundColor(data.format.act.color).setFontSize(parseInt(data.format.act.size)).setBold(data.format.act.bold).setItalic(data.format.act.italic).build();
    rtv.setTextStyle(0, actText.length, actStyle);
  }
  if (noteText.length > 0) {
    let noteStyle = SpreadsheetApp.newTextStyle().setForegroundColor(data.format.note.color).setFontSize(parseInt(data.format.note.size)).setBold(data.format.note.bold).setItalic(data.format.note.italic).build();
    let startIdx = actText.length > 0 ? actText.length + 1 : 0;
    rtv.setTextStyle(startIdx, fullText.length, noteStyle);
  }

  let topCell = sheet.getRange(startRow, col);
  topCell.setRichTextValue(rtv.build());
  if (numRows > 1 || targetColSpan > 1) targetRange.merge();
  targetRange.setBackground(data.color);
  targetRange.setHorizontalAlignment("center");
  targetRange.setVerticalAlignment("middle");
  targetRange.setWrap(true);
  
  triggerCacheUpdate(); // Flag cache for rebuild!
  return "Success";
}

// =========================================
// DECODER CACHE SYSTEM
// =========================================

function getInitialDecoderData() {
  let props = PropertiesService.getDocumentProperties();
  let lastMod = parseInt(props.getProperty('scheduleLastModified') || '0');
  let cacheTs = parseInt(props.getProperty('cacheTimestamp') || '0');

  let cachedData = loadCacheFromSheet();

  if (!cachedData || cachedData.length === 0) {
    return { status: "empty", data: [] }; // First time running
  }

  if (lastMod > cacheTs) {
    return { status: "outdated", data: cachedData }; // Send old data immediately, trigger background update
  }

  return { status: "fresh", data: cachedData }; // Data is up to date!
}

function generateNewCache() {
  let freshData = getGlobalScheduleCore(); 
  saveCacheToSheet(freshData);
  return freshData;
}

function loadCacheFromSheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("DecoderCache");
  if (!sheet) return null;
  let lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  
  // Reassemble JSON string chunks
  let chunks = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let jsonStr = chunks.map(r => r[0]).join('');
  try { return JSON.parse(jsonStr); } catch(e) { return null; }
}

function saveCacheToSheet(dataObj) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("DecoderCache");
  if (!sheet) {
    sheet = ss.insertSheet("DecoderCache");
    sheet.hideSheet(); // Keep it hidden so it doesn't clutter the file
  }
  sheet.clear();
  
  // Break massive JSON into 50k character chunks (Google Sheets Cell Limit)
  let jsonStr = JSON.stringify(dataObj);
  let chunks = [];
  for (let i = 0; i < jsonStr.length; i += 50000) {
    chunks.push([jsonStr.substring(i, i + 50000)]);
  }
  
  sheet.getRange(2, 1, chunks.length, 1).setValues(chunks);
  let now = Date.now();
  sheet.getRange(1, 1).setValue(now);
  PropertiesService.getDocumentProperties().setProperty('cacheTimestamp', now.toString());
}

function getGlobalScheduleCore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scanners = ["HM1", "HM2", "HM3", "HM4", "NM1", "NM2", "NM3", "WM"];
  let allSlots = [];
  
  scanners.forEach(scannerName => {
    try {
      let sheet = ss.getSheetByName(scannerName);
      if(!sheet) return;
      
      let mapping = getDayMapping(sheet);
      let sheetMaxRows = sheet.getMaxRows();
      let maxRows = 156; 
      if (ROW_OFFSET + maxRows - 1 > sheetMaxRows) maxRows = sheetMaxRows - ROW_OFFSET + 1;
      if (maxRows < 1) maxRows = 1;

      const dataRange = sheet.getRange(ROW_OFFSET, 2, maxRows, mapping.numCols);
      const values = dataRange.getValues();
      const backgrounds = dataRange.getBackgrounds();
      const mergedRanges = dataRange.getMergedRanges();
      let coveredCells = {};

      mergedRanges.forEach(range => {
        let rStart = range.getRow() - ROW_OFFSET;
        let rEnd = rStart + range.getNumRows();
        let cIdx = range.getColumn() - 2;
        if(rStart < 0 || rStart >= maxRows || cIdx < 0 || cIdx >= mapping.numCols) return;
        
        let val = range.getValue().toString().trim();
        if(!val) return;
        
        let parts = val.split('\n');
        allSlots.push({
          scanner: scannerName, day: mapping.dayMap[cIdx],
          startTime: minsToTime(START_TIME_MINS + (rStart * BLOCK_MINS)), endTime: minsToTime(START_TIME_MINS + (rEnd * BLOCK_MINS)),
          activity: parts[0].trim(), note: parts.slice(1).join('\n').trim(), color: range.getBackground()
        });

        for(let r = rStart; r < rEnd; r++) {
          for(let c = cIdx; c < cIdx + range.getNumColumns(); c++) coveredCells[`${r}-${c}`] = true;
        }
      });

      for (let r = 0; r < maxRows; r++) {
        for (let c = 0; c < mapping.numCols; c++) {
          if (!coveredCells[`${r}-${c}`]) {
            let val = values[r][c].toString().trim();
            if (val !== "") {
              let parts = val.split('\n');
              allSlots.push({
                scanner: scannerName, day: mapping.dayMap[c],
                startTime: minsToTime(START_TIME_MINS + (r * BLOCK_MINS)), endTime: minsToTime(START_TIME_MINS + ((r + 1) * BLOCK_MINS)),
                activity: parts[0].trim(), note: parts.slice(1).join('\n').trim(), color: backgrounds[r][c]
              });
            }
          }
        }
      }
    } catch(e) { } 
  });
  return allSlots;
}

// --- CLEANUP TOOL ---
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🛠️ MRI Tools').addItem('Auto-Merge Colored Cells (Current Tab)', 'autoMergeColors').addToUi();
}

function autoMergeColors() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  let mapping = getDayMapping(sheet);
  
  const startRow = 7; const numRows = 156; const startCol = 2; const numCols = mapping.numCols; 
  const gridRange = sheet.getRange(startRow, startCol, numRows, numCols);
  const mergedRanges = gridRange.getMergedRanges();
  const isMergedMap = Array(numRows).fill(null).map(() => Array(numCols).fill(false));
  
  mergedRanges.forEach(mRange => {
    let rStart = mRange.getRow() - startRow;
    let rEnd = rStart + mRange.getNumRows();
    let cStart = mRange.getColumn() - startCol;
    let cEnd = cStart + mRange.getNumColumns();
    for(let r = rStart; r < rEnd; r++) {
      for(let c = cStart; c < cEnd; c++) {
        if(r >= 0 && r < numRows && c >= 0 && c < numCols) isMergedMap[r][c] = true;
      }
    }
  });
  
  const backgrounds = gridRange.getBackgrounds();
  const values = gridRange.getValues();
  let mergedCount = 0;
  
  for (let c = 0; c < numCols; c++) {
    let blockStartRow = -1; let blockColor = null; let blockText = null;
    for (let r = 0; r < numRows; r++) {
      let cellColor = backgrounds[r][c]; let cellText = values[r][c].toString().trim(); let isMerged = isMergedMap[r][c];
      
      if (isMerged) {
        if (blockStartRow !== -1 && (r - blockStartRow) > 1 && blockColor !== '#ffffff' && blockColor !== '#000000') {
           sheet.getRange(startRow + blockStartRow, startCol + c, r - blockStartRow, 1).mergeVertically(); mergedCount++;
        }
        blockStartRow = -1; blockColor = null; blockText = null; continue;
      }
      if (blockStartRow === -1) { blockStartRow = r; blockColor = cellColor; blockText = cellText; continue; }
      
      if (cellColor === blockColor && (cellText === "" || cellText === blockText) && cellColor !== '#ffffff' && cellColor !== '#000000') {
      } else {
        if ((r - blockStartRow) > 1 && blockColor !== '#ffffff' && blockColor !== '#000000') {
          sheet.getRange(startRow + blockStartRow, startCol + c, r - blockStartRow, 1).mergeVertically(); mergedCount++;
        }
        blockStartRow = r; blockColor = cellColor; blockText = cellText;
      }
    }
    if (blockStartRow !== -1 && (numRows - blockStartRow) > 1 && blockColor !== '#ffffff' && blockColor !== '#000000') {
      sheet.getRange(startRow + blockStartRow, startCol + c, numRows - blockStartRow, 1).mergeVertically(); mergedCount++;
    }
  }
  
  triggerCacheUpdate(); // Flag cache for rebuild!
  ui.alert(`✅ Cleanup Complete!\n\nMerged ${mergedCount} new blocks of color on the ${sheet.getName()} tab.`);
}
