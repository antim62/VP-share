/**
 * Taakvenster voor knop 13 ("Contactpersonen naar Outlook").
 *
 * Werkt in twee stappen, zodat je één keer ALLE contactpersonen vanuit VP kunt exporteren en
 * er hier telkens een andere Ploeg/Post uit kunt kiezen (in plaats van vooraf in VP per Post
 * of Ploeg apart te moeten exporteren):
 *
 *   1. "Overzicht laden" leest het actieve werkblad (moet de ruwe VP-contactpersonenexport
 *      zijn) en verzamelt alle unieke Ploegen/Posten die erin voorkomen (zelfde
 *      voorrangsregel als de oude BepaalMapNaam: een rij met een Ploeg telt als die Ploeg,
 *      ongeacht de Post-kolom; anders telt de Post als die gevuld is). Onthoudt de naam van
 *      dat bronwerkblad, zodat stap 2 daar altijd naar terugleest - ook als er inmiddels een
 *      ander werkblad actief is (bijv. na een eerdere run, wanneer Excel het nieuwe
 *      "VPContacten"-werkblad activeert).
 *   2. "Voorbereiden voor deze selectie" filtert de rijen op de gekozen Ploeg/Post, zet
 *      datumtekst om (zelfde logica als de oude Office Script "13 - Voorbereiden..."), en
 *      zet het resultaat op een schoon werkblad "VPContacten" als tabel - klaar voor de
 *      Power Automate-flow. Een vorige run op dat werkblad wordt eerst volledig gewist.
 */

let sourceSheetName = null;
let ploegColIdx = -1;
let postColIdx = -1;

Office.onReady(() => {
  document.getElementById("btn-laden").addEventListener("click", overzichtLaden);
  document.getElementById("btn-voorbereiden").addEventListener("click", voorbereiden);
});

async function overzichtLaden() {
  const statusEl = document.getElementById("status");
  const select = document.getElementById("groep-select");
  const btnVoorbereiden = document.getElementById("btn-voorbereiden");

  statusEl.className = "status";
  statusEl.textContent = "Bezig met inlezen...";
  select.innerHTML = "";
  select.disabled = true;
  btnVoorbereiden.disabled = true;
  sourceSheetName = null;

  try {
    const result = await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      const usedRange = sheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject,values");
      await context.sync();

      if (usedRange.isNullObject) {
        return { fout: "Het actieve werkblad is leeg. Zet de ruwe VP-contactpersonenexport actief en probeer opnieuw." };
      }

      const values = usedRange.values;
      if (!values[0] || values[0][0] !== "Persoon - Voornaam") {
        return {
          fout:
            'Dit lijkt geen geldige VP-contactpersonenexport: cel A1 bevat niet "Persoon - Voornaam". ' +
            'Zorg dat het werkblad met de (ongefilterde) VP-export actief is voordat je op "Overzicht laden" klikt.',
        };
      }

      const headerRow = values[0];
      const ploegCol = headerRow.indexOf("Groep categorieën - Ploeg");
      const postCol = headerRow.indexOf("Groep categorieën - Post");
      if (ploegCol < 0 && postCol < 0) {
        return { fout: 'Kolommen "Groep categorieën - Ploeg" en "Groep categorieën - Post" niet gevonden in dit bestand.' };
      }

      const gevonden = new Map(); // "type|waarde" -> {type, value, label, count}
      for (let r = 1; r < values.length; r++) {
        const groep = determineRowGroup(values[r], ploegCol, postCol);
        if (!groep) continue;
        const key = groep.type + "|" + groep.value;
        if (!gevonden.has(key)) gevonden.set(key, { ...groep, count: 0 });
        gevonden.get(key).count++;
      }

      return {
        sheetName: sheet.name,
        ploegCol,
        postCol,
        groepen: Array.from(gevonden.values()).sort((a, b) => a.label.localeCompare(b.label, "nl")),
      };
    });

    if (result.fout) {
      statusEl.className = "status error";
      statusEl.textContent = result.fout;
      return;
    }

    if (result.groepen.length === 0) {
      statusEl.className = "status error";
      statusEl.textContent = "Geen enkele rij heeft een Ploeg of Post ingevuld - er valt niets te selecteren.";
      return;
    }

    sourceSheetName = result.sheetName;
    ploegColIdx = result.ploegCol;
    postColIdx = result.postCol;

    for (const g of result.groepen) {
      const opt = document.createElement("option");
      opt.value = `${g.type}|${g.value}`;
      opt.textContent = `${g.label} (${g.count})`;
      select.appendChild(opt);
    }
    select.disabled = false;
    btnVoorbereiden.disabled = false;

    statusEl.className = "status ok";
    statusEl.textContent =
      `${result.groepen.length} groep(en) gevonden op werkblad "${result.sheetName}". ` +
      `Kies hierboven een Ploeg of Post en klik op "2. Voorbereiden voor deze selectie".`;
  } catch (error) {
    console.error(error);
    statusEl.className = "status error";
    statusEl.textContent = "Er ging iets mis: " + (error && error.message ? error.message : error);
  }
}

async function voorbereiden() {
  const statusEl = document.getElementById("status");
  const select = document.getElementById("groep-select");

  if (!sourceSheetName || !select.value) {
    statusEl.className = "status error";
    statusEl.textContent = 'Klik eerst op "1. Overzicht laden" en kies daarna een Ploeg/Post.';
    return;
  }

  const scheidingIdx = select.value.indexOf("|");
  const selType = select.value.slice(0, scheidingIdx);
  const selValue = select.value.slice(scheidingIdx + 1);

  statusEl.className = "status";
  statusEl.textContent = "Bezig...";

  try {
    const resultaat = await Excel.run(async (context) => {
      const sourceSheet = context.workbook.worksheets.getItemOrNullObject(sourceSheetName);
      sourceSheet.load("isNullObject");
      await context.sync();
      if (sourceSheet.isNullObject) {
        return { fout: `Bronwerkblad "${sourceSheetName}" bestaat niet meer - klik opnieuw op "1. Overzicht laden".` };
      }

      const usedRange = sourceSheet.getUsedRangeOrNullObject();
      usedRange.load("isNullObject,values,formulas,numberFormat");
      await context.sync();
      if (usedRange.isNullObject) {
        return { fout: `Bronwerkblad "${sourceSheetName}" is leeg.` };
      }

      const values = usedRange.values;
      const formulas = usedRange.formulas;
      const numberFormats = usedRange.numberFormat;
      const headerRow = values[0];

      // Filteren op de gekozen groep, en tegelijk datumtekst omzetten (zelfde logica als
      // voorheen) - alleen voor de rijen die overblijven.
      const outValues = [headerRow.slice()];
      const outFormats = [numberFormats[0].slice()];

      for (let r = 1; r < values.length; r++) {
        const groep = determineRowGroup(values[r], ploegColIdx, postColIdx);
        if (!groep || groep.type !== selType || groep.value !== selValue) continue;

        const rowValues = values[r].slice();
        const rowFormats = numberFormats[r].slice();
        for (let c = 0; c < rowValues.length; c++) {
          const f = formulas[r][c];
          if (typeof f === "string" && f.startsWith("=")) { rowValues[c] = f; continue; }

          const v = rowValues[c];
          if (typeof v !== "string" || v.trim() === "") continue;

          const parsed = parseDutchDateTime(v);
          if (!parsed) continue;

          rowValues[c] = parsed.serial;
          rowFormats[c] = parsed.hasDate ? "dd-mm-yyyy" : "hh:mm";
        }
        outValues.push(rowValues);
        outFormats.push(rowFormats);
      }

      if (outValues.length <= 1) {
        return { fout: "Geen enkele rij komt overeen met de gekozen selectie - niets om weg te schrijven." };
      }

      // Uitvoer-werkblad "VPContacten": bestaand werkblad eerst helemaal leegmaken (incl. een
      // eventuele oude tabel), of nieuw aanmaken als het nog niet bestaat.
      let outSheet = context.workbook.worksheets.getItemOrNullObject("VPContacten");
      outSheet.load("isNullObject");
      await context.sync();

      if (outSheet.isNullObject) {
        outSheet = context.workbook.worksheets.add("VPContacten");
      } else {
        const bestaandeTabellen = outSheet.tables;
        bestaandeTabellen.load("items");
        await context.sync();
        for (const t of bestaandeTabellen.items) t.delete();

        const oudGebruikt = outSheet.getUsedRangeOrNullObject();
        oudGebruikt.load("isNullObject");
        await context.sync();
        if (!oudGebruikt.isNullObject) oudGebruikt.clear(Excel.ClearApplyTo.All);
      }

      const rowCount = outValues.length;
      const colCount = headerRow.length;
      const outRange = outSheet.getRangeByIndexes(0, 0, rowCount, colCount);
      outRange.values = outValues;
      outRange.numberFormat = outFormats;

      const table = outSheet.tables.add(outRange, true);
      table.name = "VPContacten";
      outRange.format.autofitColumns();
      outSheet.activate();
      await context.sync();

      return {
        aantal: outValues.length - 1,
        mapNaam: selType === "post" ? `Post ${selValue}` : selValue,
      };
    });

    if (resultaat.fout) {
      statusEl.className = "status error";
      statusEl.textContent = resultaat.fout;
      return;
    }

    statusEl.className = "status ok";
    statusEl.textContent =
      `Klaar: ${resultaat.aantal} contactpersonen van "${resultaat.mapNaam}" op werkblad "VPContacten" ` +
      `gezet als tabel. Sla het bestand nu op in OneDrive/VP en start daarna de Power Automate-flow.`;
  } catch (error) {
    console.error(error);
    statusEl.className = "status error";
    statusEl.textContent = "Er ging iets mis: " + (error && error.message ? error.message : error);
  }
}
