# VP-invoegtoepassing als Office Add-in (eigen ribbon-tab "VP")

Dit is de volledige vervanger voor de oude `.xlam`-invoegtoepassing: een echte, eigen
ribbon-tab **"VP"** met dezelfde 5 groepen en 13 knoppen (zelfde teksten en tooltips als het
origineel), die niet via macro's werkt maar via een Office Add-in (HTML/JavaScript). Dit *kan*
wél gewoon werken op een omgeving waar macro's zijn uitgeschakeld, want een Add-in is geen
macro/VBA en draait in een eigen, beveiligde sandbox.

Dit bouwt voort op de eerder geleverde 13 Office Scripts (map `office_scripts` uit de vorige
levering) - dezelfde bewerkingslogica is hier herschreven naar de Excel JavaScript API die
Add-ins gebruiken, en gekoppeld aan een eigen ribbon-tab in plaats van losse knoppen in het
Automate-paneel.

## Wat zit erin

```
manifest.xml                 - het "installatiebestand": beschrijft de ribbon-tab, groepen,
                                knoppen, iconen en welke bestanden erbij horen.
assets/                       - icoon-16/32/80.png: het echte VP-logo (tab en groepen).
assets/knop/                    knopN-16/32/80.png: 13 eigen ontworpen iconen, één per knop
                                (zie "Eigen iconen" hieronder).
src/shared/vp-helpers.js      - gedeelde reken-/parseerlogica (datum/tijd, valuta, percentage,
                                getal, samengevoegde cellen), gebruikt door beide onderstaande
                                bestanden.
src/functions/functions.html  - onzichtbaar "achtergrondvenstertje" dat de knoppen 1 t/m 12
src/functions/functions.js      afhandelt (elke knop = één JS-functie).
src/taskpane/taskpane.html    - het paneel dat opent bij knop 13 ("Contactpersonen naar
src/taskpane/taskpane.js        Outlook"): bereidt de Excel-tabel voor en linkt door naar de
src/taskpane/taskpane.css       Power Automate-flow (zie de eerder geleverde gidsen daarover).
package.json                  - hulpscripts om lokaal te testen (certificaat, mini-webserver,
                                manifest-check).
```

## Belangrijke keuzes / beperkingen t.o.v. het origineel

- **Knop 13 doet zelf geen Outlook meer.** Dat kon in VBA via lokale Outlook-automatisering
  (COM); een Add-in draait in een sandbox zonder die toegang. Knop 13 opent nu een taakvenster
  dat (a) de Excel-tabel voorbereidt (zelfde als het losse Office Script) en (b) doorlinkt naar
  de Power Automate-flow die het echte contacten/agenda-werk doet. Zie de eerder geleverde
  bestanden `Power Automate flow - Contactpersonen naar Outlook.md` en
  `Power Automate - stappenplan voor beginners.md`.
- **Geen scheidingsstreepje** vóór knop 8 ("Alle tekstbewerkingen") binnen de groep
  Tekstbewerkingen - Add-in-ribbons ondersteunen geen losse scheider binnen een groep. Puur
  optisch verschil, geen functioneel verschil.
- **Knop 3** (Automatische kolombreedte) heft nooit meer automatisch eerst samenvoegingen op
  (dat was in de Office Script-versie een parameter die het Automate-paneel als invoerveld
  toonde - een ribbon-knop in een Add-in kan dat niet vanzelf). Wil je dat terug, dat is een
  kleine uitbreiding (een keuzevakje in een taakvenster, zoals bij knop 13) - zeg het maar.
- **Eigen icoon per knop** (zie "Eigen iconen" hieronder) - het origineel gebruikte per knop een
  ander Office-systeemicoon (`imageMso`, bijv. `MergeCenter`, `BordersAll`); die kunnen Add-ins
  niet hergebruiken, dus is hiervoor een nieuwe, bijpassende set gemaakt.

## Eigen iconen

`assets/icon-16/32/80.png` zijn het echte VP-logo (gebruikt voor de tab en de 4 groepen),
automatisch uitgesneden/geschaald vanaf het logo dat je aanleverde. Bij 16px verliest een
gedetailleerd logo onvermijdelijk wat scherpte - dat is inherent aan zo'n klein formaat.

`assets/knop/knopN-16/32/80.png` zijn 13 nieuw ontworpen iconen, één per knop, in hetzelfde
blauw als het VP-logo zodat het geheel als één stijl oogt: een gum (leeg maken), 4 losse
vakjes (samenvoeging opheffen), een balk met pijlen (kolombreedte), een kalender (datum/tijd),
€ (valuta), % (percentage), # (getal), een toverstaf (alle tekstbewerkingen), links/midden/
rechts-uitgelijnde streepjes (uitlijnen), een dichtgetekend rasterblokje (omlijnen) en een
contactpersoon-silhouet (Outlook). Het origineel gebruikte per knop een Office-systeemicoon
(`imageMso`) - die kunnen Add-ins helaas niet hergebruiken (geen directe koppeling mogelijk),
vandaar deze nieuwe set in plaats daarvan.

Wil je op termijn een ander icoon voor een specifieke knop: vervang het bijbehorende
`knopN-16/32/80.png`-drietal in `assets/knop/` door je eigen PNG's op dezelfde afmetingen -
er hoeft dan verder niets in `manifest.xml` te veranderen, de bestandsnamen blijven gelijk.

## Lokaal testen

Vereist: [Node.js](https://nodejs.org) (LTS-versie is prima).

1. Open een terminal in deze map en installeer de hulpprogramma's:
   ```
   npm install
   ```
2. Eenmalig een lokaal ontwikkelcertificaat installeren (nodig omdat Office alleen add-ins
   over https vertrouwt, ook lokaal):
   ```
   npm run cert
   ```
   Klik "Ja"/"Trust" als Windows/macOS om bevestiging vraagt.
3. Start de lokale server:
   - Windows: `npm start`
   - macOS/Linux: `npm run start:mac`

   Laat dit venster openstaan - zolang deze server draait, is `https://localhost:3000/...`
   bereikbaar (precies de URL's die in `manifest.xml` staan).
4. Controleer of het manifest geldig is:
   ```
   npm run validate
   ```
5. Laad de add-in in Excel (de exacte menunaam verschilt per Excel-versie/build - soms
   "Invoegen", soms "Startpagina"/"Start"):
   - **Excel op het web**: open een werkmap → **Startpagina** → **Invoegtoepassingen** →
     **Meer invoegtoepassingen** → tabblad **Mijn invoegtoepassingen** →
     **Invoegtoepassing uploaden** → kies `manifest.xml`. Werkt meestal direct, geen extra
     instelling nodig.
   - **Excel desktop (Windows/Mac)**: hier werkt "Invoegtoepassing uploaden" op een
     IT-beheerde laptop vaak niet direct (melding over "catalogi voor invoegtoepassingen").
     Gebruik dan de **gedeelde-map-catalogus**: zet `manifest.xml` in een (eventueel lokale)
     map, voeg die map toe bij **Bestand → Opties → Vertrouwenscentrum → Instellingen voor
     het Vertrouwenscentrum → Vertrouwde catalogi voor invoegtoepassingen**, vink
     "Weergeven in menu" aan, herstart Excel. Ga daarna naar **Start/Invoegen →
     Invoegtoepassingen ophalen** → tabblad **"GEDEELDE MAP"** → kies de VP-invoegtoepassing.
   - Of automatisch via de command line (werkt meestal alleen bij Excel op het web/Mac,
     minder betrouwbaar op een beheerde Windows-laptop): `npm run sideload`.
6. Je zou nu bovenin een tabblad **"VP"** moeten zien met de 5 groepen en 13 knoppen.

**Testen wijzigen**: pas een `.js`-bestand aan, sla op, en druk in Excel op de knop nogmaals (of
sluit/heropen het werkblad) - de lokale server serveert automatisch de nieuwste versie, een
herstart is niet nodig. Wijzig je iets in `manifest.xml` zelf (bijv. een knoptekst), dan moet je
de add-in wel opnieuw uploaden/sideloaden.

## Van test naar productie

Zolang `manifest.xml` naar `https://localhost:3000/...` verwijst, werkt dit alleen op dit ene
apparaat terwijl deze terminal openstaat - niet praktisch voor dagelijks gebruik of om met
collega's te delen.

Voor blijvend gebruik heb je een "echte" hosting-plek nodig voor de bestanden in `assets/` en
`src/` (statische bestanden, geen server-logica nodig - een simpele webserver volstaat). Met
jouw eigen Proxmox-omgeving kan dat prima zelf: bijvoorbeeld een kleine nginx- of IIS-VM met een
geldig HTTPS-certificaat (Let's Encrypt werkt hiervoor), of - als je bij Veiligheidsregio
Brabant-Noord toegang hebt tot Microsoft 365-opties - een Azure Static Web App / SharePoint-site
die statische bestanden mag serveren.

Stappen zodra je een hosting-plek hebt:

1. Zet de hele inhoud van deze map (behalve `package.json`/`README.md`, die zijn alleen voor
   lokaal testen) op die server, met dezelfde mapstructuur.
2. Vervang in `manifest.xml` **elke** `https://localhost:3000` door je echte domeinnaam, bijv.
   `https://vp-addin.jouw-domein.nl`.
3. Run `npm run validate` nogmaals om te checken dat het manifest nog klopt.
4. Test opnieuw via "Invoegtoepassing uploaden" met het bijgewerkte manifest.

## Beheerde werklaptop zonder netwerkdeling (geen UNC-share mogelijk)

Draai je op een door IT beheerde laptop waarop netwerkdeling (SMB) is uitgeschakeld, dan valt de
gedeelde-map-catalogus (zie "Lokaal testen" hierboven) af - die vereist een échte
Windows-netwerkshare (\\\\server\\share), geen gewoon lokaal mapje. Twee alternatieven:

- **Voor nu: Excel op het web.** Dat werkte al zonder enige Trust Center-instelling nodig te
  hebben (zie stap 5 hierboven) - waarschijnlijk omdat dit via een andere, tenant-brede
  instelling loopt in plaats van een lokale Windows-GPO. Prima stopgap zolang je op de
  beheerde laptop werkt.
- **Voor blijvend gebruik: vraag IT om Centralized Deployment** (zie hieronder) - dat is de
  bedoelde, ondersteunde route voor precies dit scenario (add-in met een eigen ribbon-tab op
  een beheerd apparaat) en heeft aan jouw kant helemaal geen Trust Center-instelling nodig; IT
  wijst 'm centraal toe en hij verschijnt vanzelf.

**Let op - een SharePoint-catalogus lijkt een logische derde optie (jullie delen toch al via
Teams/SharePoint) maar werkt hier niet**: Microsoft ondersteunt add-ins met een eigen
ribbon-tab (technisch: het `VersionOverrides`-onderdeel van het manifest, wat wij hier
gebruiken) uitdrukkelijk niet via een SharePoint-app-catalogus - dat mechanisme is alleen voor
oudere, taakvenster-only add-ins zonder ribbon-knoppen. Bovendien vereist het instellen ervan
sowieso SharePoint-beheerdersrechten, dus dit is voor jou als eindgebruiker sowieso geen
zelfservice-optie. Sla deze route dus over en ga voor de twee opties hierboven.

## Simpel delen met één collega (zonder IT)

Zolang het manifest naar `localhost` wijst, werkt de add-in alleen op jouw eigen pc, met jouw
lokale server aan. Voor een collega heb je twee dingen nodig: (1) de bestanden ergens hosten
waar beide pc's bij kunnen, en (2) bij de collega dezelfde soort catalogus-instelling die jij
net op je eigen pc hebt gedaan.

**Snelste weg zonder IT-goedkeuring: gratis statische hosting (bijv. GitHub Pages).** Jij bent
toch al thuis in Linux/homelab-spul, dus dit kost je een paar minuten:

1. Zet de map `assets/` en `src/` (dus niet `package.json`/`README.md`, die zijn alleen voor
   lokaal testen) in een GitHub-repository en zet **Settings → Pages** aan. Je krijgt een vaste
   URL terug, bijv. `https://jouwgebruikersnaam.github.io/vp-addin/`.
2. Vervang in `manifest.xml` elke `https://localhost:3000` door die URL.
3. `npm run validate` om te checken dat het manifest nog klopt.
4. Test 'm zelf nogmaals (zie "Lokaal testen" hierboven, alleen heb je nu geen lokale server
   meer nodig - de bestanden staan online).
5. Stuur je collega **alleen het bijgewerkte `manifest.xml`-bestand** (niet de hele map - de
   code staat nu al online).
6. Je collega doet exact wat jij net op je eigen pc deed: `manifest.xml` in een (eventueel
   lokale) map zetten, die map toevoegen bij **Vertrouwde catalogi voor invoegtoepassingen**,
   Excel herstarten, en de add-in toevoegen via **Invoegtoepassingen ophalen → GEDEELDE MAP**.
   Geen Node.js, geen lokale server nodig aan die kant.

**Let op**: bij een gratis publieke host zoals GitHub Pages is de add-in-code (niet je
Excel-data, alleen de JavaScript/HTML-bestanden hierboven) in principe door iedereen met de URL
te bekijken. Er staan geen wachtwoorden/gevoelige gegevens in, dus functioneel is dat geen
probleem, maar check even of dat voor een brandweer-tool binnen jullie regels past. Twijfel je:
vraag IT naar een interne hosting-plek (zie hieronder) in plaats van een publieke.

## Centraal uitrollen via IT (zodat collega's 'm niet zelf hoeven te uploaden)

Zodra het manifest naar een stabiele, door IT vertrouwde hosting-plek wijst, kan de IT-afdeling
de add-in organisatiebreed (of voor een groep gebruikers) uitrollen via **Centralized
Deployment** in het Microsoft 365 beheercentrum (**admin.microsoft.com** → **Instellingen** →
**Integrated apps** → **Aangepaste apps uploaden**, daar het `manifest.xml`-bestand aanleveren).
Dat is dezelfde route als waarmee bedrijven ook andere Office-invoegtoepassingen uitrollen, dus
IT hoeft hier niets bijzonders voor te doen - alleen het manifest vertrouwen en toewijzen aan de
juiste gebruikers/groepen.

## Vastlopen tijdens het testen?

Stuur een screenshot van de foutmelding (in Excel zelf, of via **F12**-devtools als je in Excel
op het web test) en ik kijk met je mee - dit is precies dezelfde aanpak als bij de Power
Automate-flow.
