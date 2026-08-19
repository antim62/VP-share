/**
 * Taakvenster voor knop 13 ("Contactpersonen naar Outlook").
 * Poort van Office Script "13 - Voorbereiden voor Outlook-export.ts" naar de reguliere
 * Excel JavaScript API (zie de kanttekening bovenin functions.js over de API-verschillen).
 */

Office.onReady(() => {
  document.getElementById("btn-voorbereiden").addEventListener("click", voorbereiden);
});

async function voorbereiden() {
  const statusEl = document.getElementById("status");
  statusEl.className = "status";
  statusEl.textContent = "Bezig...";

  try {
    const resultaat = await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject");
      await context.sync();
      if (usedRange.isNullObject) {
        return { fout: "Werkblad is leeg, niets te doen." };
      }

      usedRange.load("values,formulas,numberFormat,address");
      await context.sync();

      const values = usedRange.values;

      if (!values[0] || values[0][0] !== "Persoon - Voornaam") {
        return { fout: 'Dit lijkt geen geldig VP-contactpersonenbestand: cel A1 bevat niet "Persoon - Voornaam".' };
      }

      // MapNaam bepalen (zelfde logica als BepaalMapNaam in de oude VBA / script 13)
      const headerRow = values[0];
      const ploegCol = headerRow.indexOf("Groep categorieën - Ploeg");
      const postCol = headerRow.indexOf("Groep categorieën - Post");
      let mapNaam = "Brandweer Brabant-Noord";
      if (values.length > 1) {
        const ploeg = ploegCol >= 0 ? String(values[1][ploegCol]).trim() : "";
        const post = postCol >= 0 ? String(values[1][postCol]).trim() : "";
        if (ploeg) mapNaam = ploeg;
        else if (post) mapNaam = `Post ${post}`;
      }

      // Datum/tijd-tekst omzetten (rij 0 = koprij, die slaan we over)
      const formulas = usedRange.formulas;
      const numberFormats = usedRange.numberFormat;

      for (let r = 1; r < values.length; r++) {
        for (let c = 0; c < values[r].length; c++) {
          const f = formulas[r][c];
          if (typeof f === "string" && f.startsWith("=")) { values[r][c] = f; continue; }

          const v = values[r][c];
          if (typeof v !== "string" || v.trim() === "") continue;

          const parsed = parseDutchDateTime(v);
          if (!parsed) continue;

          values[r][c] = parsed.serial;
          numberFormats[r][c] = parsed.hasDate ? "dd-mm-yyyy" : "hh:mm";
        }
      }
      usedRange.values = values;
      usedRange.numberFormat = numberFormats;

      // Van het bereik een tabel maken, of een bestaande tabel bijwerken
      const tables = sheet.tables;
      tables.load("items");
      await context.sync();

      let table;
      if (tables.items.length > 0) {
        table = tables.items[0];
        table.resize(usedRange);
      } else {
        table = sheet.tables.add(usedRange, true);
      }
      table.name = "VPContacten";
      await context.sync();

      return {
        aantal: values.length - 1,
        mapNaam: mapNaam,
      };
    });

    if (resultaat.fout) {
      statusEl.className = "status error";
      statusEl.textContent = resultaat.fout;
      return;
    }

    statusEl.className = "status ok";
    statusEl.textContent =
      `Klaar: ${resultaat.aantal} contactpersonen verwerkt, tabel "VPContacten" bijgewerkt. ` +
      `Outlook-contactenlijst wordt: "${resultaat.mapNaam}". Sla het bestand nu op in OneDrive/VP ` +
      `en start daarna de Power Automate-flow (link hieronder).`;
  } catch (error) {
    console.error(error);
    statusEl.className = "status error";
    statusEl.textContent = "Er ging iets mis: " + (error && error.message ? error.message : error);
  }
}
