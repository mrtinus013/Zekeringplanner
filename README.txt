ZEKERINGPLANNER v0.4
====================

GitHub Pages-ready webapp voor de Enexis GFF-zekeringwissellijst.

WERKING
- Importeer de actuele GFF_zekeringwissel_20251229.xlsx lokaal in de browser.
- De app kiest automatisch werkblad "data" en herkent de huidige kolommen:
  D = Netstation/LS-kast
  F = Huidige waarde
  G = LS Veldnummer
  I = Opmerking specialisme
  J = Opmerkingen vanuit Engineering
  K = Zekering vervangen medewerker
- Locaties worden gekoppeld aan de bestaande Enexis POI-bronnen voor stations en verdeelkasten.

STATUSSEN
- Wit / geen kleur + geen medewerker: open, normaal uitvoeren.
- Groen OF een naam in "Zekering vervangen medewerker": gereed, niet meenemen.
- Geel: niet kunnen wisselen; opmerking specialisme tonen + herinnering "Mail Kevin".
- Oranje: Engineering-opmerking (vaak verzwaring/vermindering); apart tonen, niet automatisch meenemen.
- Blauw: wordt opgepakt door Henri van der Vleuten tijdens onderhoud; apart tonen.

AFVINKEN
- Iedere normale open richting heeft een knop "Gedaan".
- Afgevinkte regels verdwijnen direct uit de route en boodschappenlijst.
- Ze worden lokaal opgeslagen in de browser en verschijnen onder "Uitgevoerd".
- Het overzicht kan op Vandaag of Alles worden gezet en gekopieerd worden.
- Met "Herstel" kan een foutieve afvinking worden teruggedraaid.
- Dit schrijft NIET terug naar SharePoint/Excel.

BOODSCHAPPENLIJST
- Standaard 3 zekeringen per richting (aanpasbaar in Instellingen).
- Restbak/onbekende waarden worden niet als zekeringtype opgeteld maar als controlepunt gemeld.
- Geel, Engineering, Henri en afgevinkte regels tellen standaard niet mee.

MAIL KEVIN
- Bij gele regels staat een knop "Maak mail".
- In Instellingen kan optioneel Kevins e-mailadres worden ingevuld.
- Zonder e-mailadres opent de mailapp met onderwerp en tekst maar zonder ontvanger.

PRIVACY
- Excel-data wordt alleen in de browser verwerkt en lokaal opgeslagen.
- Zet de Excel NIET in een publieke GitHub-repository.
- Alleen de POI-bronnen worden online opgehaald.

INSTALLATIE GITHUB PAGES
1. Upload index.html, app.js, styles.css, manifest.webmanifest en README.txt naar een repository.
2. Settings > Pages > Deploy from a branch > main / root.
3. Open de GitHub Pages-link.
4. Op iPhone: Safari > Deel > Zet op beginscherm.
