/**
 * Gedeelde hulpfuncties voor de VP-invoegtoepassing (Office Add-in).
 * Wordt zowel door het onzichtbare functiebestand (functions.js, voor de ribbon-knoppen)
 * als door het taakvenster (taskpane.js, voor knop 13) gebruikt.
 *
 * Dit zijn pure JavaScript-functies (geen Excel-API-aanroepen erin) - dezelfde logica als
 * in de Office Scripts-versie (01 t/m 13), 1-op-1 overgezet naar de reguliere Excel
 * JavaScript API die Office Add-ins gebruiken (Excel.run/context.sync, in plaats van de
 * ExcelScript-API van Office Scripts - de methodenamen verschillen op een paar plekken,
 * zie de kanttekeningen in functions.js).
 */

function excelSerial(year, month, day, hour, minute, second) {
  hour = hour || 0; minute = minute || 0; second = second || 0;
  const dayMs = Date.UTC(year, month - 1, day);
  const epochMs = Date.UTC(1899, 11, 30);
  const days = Math.round((dayMs - epochMs) / 86400000);
  const dayFraction = (hour * 3600 + minute * 60 + second) / 86400;
  return days + dayFraction;
}

function parseDutchDateTime(text) {
  const s = text.trim();

  let m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m) {
    const d = m[1], mo = m[2], y = m[3], h = m[4], mi = m[5], se = m[6];
    const year = y.length === 2 ? 2000 + parseInt(y) : parseInt(y);
    return {
      serial: excelSerial(year, parseInt(mo), parseInt(d), parseInt(h), parseInt(mi), se ? parseInt(se) : 0),
      hasDate: true,
      hasTime: true,
    };
  }

  m = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/);
  if (m) {
    const d = m[1], mo = m[2], y = m[3];
    const year = y.length === 2 ? 2000 + parseInt(y) : parseInt(y);
    return { serial: excelSerial(year, parseInt(mo), parseInt(d)), hasDate: true, hasTime: false };
  }

  m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m) {
    const h = m[1], mi = m[2], se = m[3];
    const seconds = parseInt(h) * 3600 + parseInt(mi) * 60 + (se ? parseInt(se) : 0);
    return { serial: seconds / 86400, hasDate: false, hasTime: true };
  }

  return null;
}

function parseCurrencyNumber(text) {
  let s = text.trim().replace(/[€$]/g, "").trim();
  if (s === "") return null;

  if (/^-?\d{1,3}(\.\d{3})*(,\d+)?$/.test(s)) {
    s = s.replace(/\./g, "").replace(",", "."); // NL-notatie: 1.234,56
  } else if (/^-?\d+([.,]\d+)?$/.test(s)) {
    s = s.replace(",", ".");
  } else {
    return null;
  }

  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function parsePlainNumber(text) {
  const s = text.trim();
  if (s === "" || s.length === 1) return null; // origineel sloeg losse tekens over

  let normalized;
  if (/^-?\d{1,3}(\.\d{3})*(,\d+)?$/.test(s)) {
    normalized = s.replace(/\./g, "").replace(",", ".");
  } else if (/^-?\d+([.,]\d+)?$/.test(s)) {
    normalized = s.replace(",", ".");
  } else {
    return null;
  }

  const n = parseFloat(normalized);
  return isNaN(n) ? null : n;
}

/**
 * Bepaalt bij welke Ploeg of Post een rij van een VP-contactpersonenexport hoort - zelfde
 * voorrangsregel als BepaalMapNaam in de oude VBA: heeft de rij een Ploeg ingevuld, dan telt
 * die (ongeacht wat er in de Post-kolom staat); anders telt de Post, als die gevuld is.
 * Gebruikt door zowel het overzicht (welke groepen komen voor) als het filteren (welke rijen
 * horen bij de gekozen groep) in taskpane.js.
 */
function determineRowGroup(row, ploegCol, postCol) {
  const ploeg = ploegCol >= 0 && row[ploegCol] != null ? String(row[ploegCol]).trim() : "";
  if (ploeg) return { type: "ploeg", value: ploeg, label: ploeg };
  const post = postCol >= 0 && row[postCol] != null ? String(row[postCol]).trim() : "";
  if (post) return { type: "post", value: post, label: `Post ${post}` };
  return null;
}

function parseCellRef(cell) {
  const m = cell.match(/^([A-Z]+)(\d+)$/);
  if (!m) return { row: 0, col: 0 };
  let col = 0;
  for (let i = 0; i < m[1].length; i++) {
    col = col * 26 + (m[1].charCodeAt(i) - 64);
  }
  return { row: parseInt(m[2], 10) - 1, col: col - 1 };
}

/**
 * Bouwt een boolean-grid (zelfde afmeting als usedRange) dat aangeeft welke cellen bij een
 * samengevoegde cel horen. Verwacht dat usedRange al geladen is met
 * "rowIndex,columnIndex,rowCount,columnCount" en mergedAreas met "isNullObject,areas/items/address".
 */
function buildMergedMask(usedRange, mergedAreas) {
  const rowCount = usedRange.rowCount;
  const colCount = usedRange.columnCount;
  const mask = [];
  for (let r = 0; r < rowCount; r++) mask.push(new Array(colCount).fill(false));
  if (mergedAreas.isNullObject) return mask;

  const startRow = usedRange.rowIndex;
  const startCol = usedRange.columnIndex;

  for (const area of mergedAreas.areas.items) {
    const addr = area.address;
    const cellPart = addr.includes("!") ? addr.split("!")[1] : addr;
    const parts = cellPart.split(":");
    const fromCell = parts[0];
    const toCellRaw = parts[1];
    const from = parseCellRef(fromCell);
    const to = parseCellRef(toCellRaw || fromCell);

    const aRow = from.row - startRow;
    const aCol = from.col - startCol;
    const aRows = to.row - from.row + 1;
    const aCols = to.col - from.col + 1;

    for (let r = Math.max(aRow, 0); r < Math.min(aRow + aRows, rowCount); r++) {
      for (let c = Math.max(aCol, 0); c < Math.min(aCol + aCols, colCount); c++) {
        mask[r][c] = true;
      }
    }
  }
  return mask;
}
