/**
 * VP-invoegtoepassing (Office Add-in) - functiebestand voor de ribbon-knoppen.
 *
 * Dit draait onzichtbaar op de achtergrond zodra je op een ribbon-knop klikt (Action
 * xsi:type="ExecuteFunction" in manifest.xml). Elke Knop-functie is de 1-op-1 vervanger
 * van de gelijknamige Office Script (zie de map "office_scripts" van de eerdere
 * levering) - zelfde logica, maar herschreven van de ExcelScript-API (Office Scripts)
 * naar de reguliere Excel JavaScript API die Office Add-ins gebruiken. Belangrijkste
 * verschillen tussen die twee API's die hier zijn toegepast (en al eerder tot
 * fouten leidden toen we dit voor Office Scripts uitzochten - dus voor de zekerheid
 * hier expliciet gecheckt tegen de officiële Microsoft-documentatie):
 *
 *  - getUsedRangeOrNullObject() bestaat in de reguliere Excel JS API wel (in Office
 *    Scripts niet - daar is het gewoon getUsedRange(), met undefined als resultaat).
 *  - values/formulas/numberFormat zijn hier properties (met .load()/.sync() eromheen),
 *    geen aparte getValues()/getNumberFormats()-methoden zoals bij Office Scripts.
 *  - getMergedAreasOrNullObject() (met OrNullObject) i.p.v. Office Scripts' getMergedAreas().
 *  - Randen: format.borders.getItem("EdgeTop") met .style/.weight als properties, i.p.v.
 *    getRangeBorder(...).setStyle()/.setWeight().
 *
 * Elke actie draait binnen Excel.run(...) en moet altijd eindigen met event.completed(),
 * anders blijft Excel denken dat de knop nog bezig is.
 */

Office.onReady();

// --- Knop 1: Niet lege cellen leeg maken -----------------------------------------------
async function knop1(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject");
      await context.sync();
      if (usedRange.isNullObject) return;

      usedRange.load("values,formulas,rowIndex,columnIndex,rowCount,columnCount");
      const mergedAreas = usedRange.getMergedAreasOrNullObject();
      mergedAreas.load("isNullObject,areas/items/address");
      await context.sync();

      const mask = buildMergedMask(usedRange, mergedAreas);
      const values = usedRange.values;
      const formulas = usedRange.formulas;

      for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < values[r].length; c++) {
          const f = formulas[r][c];
          if (typeof f === "string" && f.startsWith("=")) { values[r][c] = f; continue; } // formules laten staan
          if (!mask[r][c] && values[r][c] === "") {
            values[r][c] = "";
          }
        }
      }

      usedRange.values = values;
      await context.sync();
    });
  } catch (error) {
    console.error("Knop1 (Niet lege cellen leeg maken):", error);
  }
  event.completed();
}
Office.actions.associate("Knop1", knop1);

// --- Knop 2: Samenvoeging cellen opheffen -----------------------------------------------
async function knop2(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject");
      await context.sync();
      if (usedRange.isNullObject) return;

      usedRange.unmerge();
      await context.sync();
    });
  } catch (error) {
    console.error("Knop2 (Samenvoeging cellen opheffen):", error);
  }
  event.completed();
}
Office.actions.associate("Knop2", knop2);

// --- Knop 3: Automatische kolombreedte --------------------------------------------------
// (De Office Script-versie had hier een optioneel scriptparameter "heftEerstSamenvoegingOp"
// omdat het Automate-paneel daar automatisch een invoerveld voor toont. Een ribbon-knop in
// een Add-in kan dat niet; deze versie heft dus nooit eerst samenvoegingen op. Wil je dat
// gedrag terug, voeg dan een keuzevakje toe aan het taakvenster net als bij knop 13.)
async function knop3(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject");
      await context.sync();
      if (usedRange.isNullObject) return;

      usedRange.format.autofitColumns();
      await context.sync();
    });
  } catch (error) {
    console.error("Knop3 (Automatische kolombreedte):", error);
  }
  event.completed();
}
Office.actions.associate("Knop3", knop3);

// --- Knop 4: Tekst -> Datum/Tijd ---------------------------------------------------------
async function knop4(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject");
      await context.sync();
      if (usedRange.isNullObject) return;

      usedRange.load("values,formulas,numberFormat");
      await context.sync();

      const values = usedRange.values;
      const formulas = usedRange.formulas;
      const numberFormats = usedRange.numberFormat;

      for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < values[r].length; c++) {
          const f = formulas[r][c];
          if (typeof f === "string" && f.startsWith("=")) { values[r][c] = f; continue; }

          const v = values[r][c];
          if (typeof v !== "string" || v.trim() === "") continue;

          const parsed = parseDutchDateTime(v);
          if (!parsed) continue;

          values[r][c] = parsed.serial;
          if (parsed.hasDate && parsed.hasTime) numberFormats[r][c] = "dd-mm-yyyy hh:mm";
          else if (parsed.hasTime) numberFormats[r][c] = "hh:mm";
          else numberFormats[r][c] = "dd-mm-yyyy";
        }
      }

      usedRange.values = values;
      usedRange.numberFormat = numberFormats;
      await context.sync();
    });
  } catch (error) {
    console.error("Knop4 (Tekst -> Datum/Tijd):", error);
  }
  event.completed();
}
Office.actions.associate("Knop4", knop4);

// --- Knop 5: Tekst -> Valuta --------------------------------------------------------------
async function knop5(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject");
      await context.sync();
      if (usedRange.isNullObject) return;

      usedRange.load("values,formulas,numberFormat");
      await context.sync();

      const values = usedRange.values;
      const formulas = usedRange.formulas;
      const numberFormats = usedRange.numberFormat;

      for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < values[r].length; c++) {
          const f = formulas[r][c];
          if (typeof f === "string" && f.startsWith("=")) { values[r][c] = f; continue; }

          const v = values[r][c];
          if (typeof v !== "string" || v.trim() === "") continue;
          if (!/[€$]/.test(v)) continue;

          const n = parseCurrencyNumber(v);
          if (n === null) continue;

          values[r][c] = n;
          numberFormats[r][c] = "€ #,##0.00";
        }
      }

      usedRange.values = values;
      usedRange.numberFormat = numberFormats;
      await context.sync();
    });
  } catch (error) {
    console.error("Knop5 (Tekst -> Valuta):", error);
  }
  event.completed();
}
Office.actions.associate("Knop5", knop5);

// --- Knop 6: Tekst -> Percentage -----------------------------------------------------------
async function knop6(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject");
      await context.sync();
      if (usedRange.isNullObject) return;

      usedRange.load("values,formulas,numberFormat");
      await context.sync();

      const values = usedRange.values;
      const formulas = usedRange.formulas;
      const numberFormats = usedRange.numberFormat;

      for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < values[r].length; c++) {
          const f = formulas[r][c];
          if (typeof f === "string" && f.startsWith("=")) { values[r][c] = f; continue; }

          const v = values[r][c];
          if (typeof v !== "string") continue;

          const match = v.trim().match(/^(-?\d+(?:[.,]\d+)?)\s*%$/);
          if (!match) continue;

          const n = parseFloat(match[1].replace(",", "."));
          if (isNaN(n)) continue;

          values[r][c] = n / 100;
          numberFormats[r][c] = "0.00%";
        }
      }

      usedRange.values = values;
      usedRange.numberFormat = numberFormats;
      await context.sync();
    });
  } catch (error) {
    console.error("Knop6 (Tekst -> Percentage):", error);
  }
  event.completed();
}
Office.actions.associate("Knop6", knop6);

// --- Knop 7: Tekst -> Getal ------------------------------------------------------------------
async function knop7(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject");
      await context.sync();
      if (usedRange.isNullObject) return;

      usedRange.load("values,formulas");
      await context.sync();

      const values = usedRange.values;
      const formulas = usedRange.formulas;

      for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < values[r].length; c++) {
          const f = formulas[r][c];
          if (typeof f === "string" && f.startsWith("=")) { values[r][c] = f; continue; }

          const v = values[r][c];
          if (typeof v !== "string" || v.trim() === "") continue;

          const n = parsePlainNumber(v);
          if (n !== null) values[r][c] = n;
        }
      }

      usedRange.values = values;
      await context.sync();
    });
  } catch (error) {
    console.error("Knop7 (Tekst -> Getal):", error);
  }
  event.completed();
}
Office.actions.associate("Knop7", knop7);

// --- Knop 8: Alle tekstbewerkingen (combineert 1, 4, 5, 6, 7 in één doorloop) -----------------
async function knop8(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject");
      await context.sync();
      if (usedRange.isNullObject) return;

      usedRange.load("values,formulas,numberFormat,rowIndex,columnIndex,rowCount,columnCount");
      const mergedAreas = usedRange.getMergedAreasOrNullObject();
      mergedAreas.load("isNullObject,areas/items/address");
      await context.sync();

      const mask = buildMergedMask(usedRange, mergedAreas);
      const values = usedRange.values;
      const formulas = usedRange.formulas;
      const numberFormats = usedRange.numberFormat;

      for (let r = 0; r < values.length; r++) {
        for (let c = 0; c < values[r].length; c++) {
          const f = formulas[r][c];
          if (typeof f === "string" && f.startsWith("=")) { values[r][c] = f; continue; }

          const v = values[r][c];

          // 1. Lege cellen opschonen (niet-samengevoegd)
          if (!mask[r][c] && v === "") {
            values[r][c] = "";
            continue;
          }

          if (typeof v !== "string" || v.trim() === "") continue;
          const s = v.trim();

          // 2. Tekst -> Datum/Tijd
          const dt = parseDutchDateTime(s);
          if (dt) {
            values[r][c] = dt.serial;
            if (dt.hasDate && dt.hasTime) numberFormats[r][c] = "dd-mm-yyyy hh:mm";
            else if (dt.hasTime) numberFormats[r][c] = "hh:mm";
            else numberFormats[r][c] = "dd-mm-yyyy";
            continue;
          }

          // 3. Tekst -> Percentage
          const pctMatch = s.match(/^(-?\d+(?:[.,]\d+)?)\s*%$/);
          if (pctMatch) {
            const n = parseFloat(pctMatch[1].replace(",", "."));
            if (!isNaN(n)) {
              values[r][c] = n / 100;
              numberFormats[r][c] = "0.00%";
              continue;
            }
          }

          // 4. Tekst -> Valuta (alleen bij €/$ teken)
          if (/[€$]/.test(s)) {
            const cur = parseCurrencyNumber(s);
            if (cur !== null) {
              values[r][c] = cur;
              numberFormats[r][c] = "€ #,##0.00";
              continue;
            }
          }

          // 5. Tekst -> Getal
          const num = parsePlainNumber(s);
          if (num !== null) {
            values[r][c] = num;
          }
        }
      }

      usedRange.values = values;
      usedRange.numberFormat = numberFormats;
      await context.sync();
    });
  } catch (error) {
    console.error("Knop8 (Alle tekstbewerkingen):", error);
  }
  event.completed();
}
Office.actions.associate("Knop8", knop8);

// --- Knop 9/10/11: Uitlijnen ---------------------------------------------------------------
async function setAlignment(horizontal) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("isNullObject");
    await context.sync();
    if (usedRange.isNullObject) return;

    usedRange.format.horizontalAlignment = horizontal;
    usedRange.format.verticalAlignment = "Center";
    await context.sync();
  });
}

async function knop9(event) {
  try {
    await setAlignment("Left");
  } catch (error) {
    console.error("Knop9 (Links uitlijnen):", error);
  }
  event.completed();
}
Office.actions.associate("Knop9", knop9);

async function knop10(event) {
  try {
    await setAlignment("Center");
  } catch (error) {
    console.error("Knop10 (Centreren):", error);
  }
  event.completed();
}
Office.actions.associate("Knop10", knop10);

async function knop11(event) {
  try {
    await setAlignment("Right");
  } catch (error) {
    console.error("Knop11 (Rechts uitlijnen):", error);
  }
  event.completed();
}
Office.actions.associate("Knop11", knop11);

// --- Knop 12: Omlijnen ----------------------------------------------------------------------
async function knop12(event) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject");
      await context.sync();
      if (usedRange.isNullObject) return;

      const edges = ["EdgeLeft", "EdgeTop", "EdgeRight", "EdgeBottom", "InsideHorizontal", "InsideVertical"];
      for (const edge of edges) {
        const border = usedRange.format.borders.getItem(edge);
        border.style = "Continuous";
        border.weight = "Thin";
      }
      await context.sync();
    });
  } catch (error) {
    console.error("Knop12 (Omlijnen):", error);
  }
  event.completed();
}
Office.actions.associate("Knop12", knop12);

// Knop 13 ("Contactpersonen naar Outlook") staat niet hier: die opent het taakvenster
// (zie manifest.xml, Action xsi:type="ShowTaskpane") - de voorbereidingslogica en de link
// naar de Power Automate-flow staan in ../taskpane/taskpane.js.
