ZEKERINGPLANNER v0.2
====================

Doel
----
Webapp om op basis van de GFF-zekeringwissellijst de resterende werkzaamheden rond een gekozen station/verdeelkast te plannen.

Specifiek afgestemd op GFF_zekeringwissel_20251229.xlsx
------------------------------------------------------
Werkblad: data
- G: te wisselen station
- H: LS_VELDNUMMER (richting)
- I: LS_SMELTVEILIGHEID_INOM (zekeringwaarde)
- J: opmerking

Status
------
- Groen = gereed, wordt altijd uitgesloten.
- Geel = actie nodig, standaard apart getoond en niet meegenomen in de boodschappenlijst.
- Geen kleur / wit = open.

De aangeleverde versie van het bestand bevatte bij analyse:
- 271 open regels
- 31 gele actie-regels
- 235 groene/gereed-regels
- 335 unieke wissellocaties

POI-koppeling
-------------
De app leest zowel poi_e_station.csv als poi_e_verdeelkast.csv uit de bestaande POI-zoeker. Daardoor kunnen ook SK/VK-locaties worden gekoppeld.

Privacy
-------
De Excel wordt alleen lokaal in de browser ingelezen en opgeslagen in localStorage. De werklijst wordt NIET naar GitHub verstuurd. Dit is bewust zo gedaan omdat de huidige GitHub Pages-repository publiek is.

Gebruik
-------
1. Open index.html via GitHub Pages/een webserver.
2. Kies de Excel. Eenmaal ingelezen blijft de verwerkte lijst lokaal op dat apparaat staan.
3. Kies een startstation/verdeelkast of gebruik Mijn locatie.
4. Kies straal en maximaal aantal locaties.
5. Klik “Maak lijst in de buurt”.
6. Onder Stations zie je open richtingen; Actie nodig toont geel; Boodschappenlijst telt benodigde zekeringen.

Aantal zekeringen
-----------------
De Excel bevat geen aantalkolom. Standaard rekent de app daarom met 3 zekeringen per richting (3 fasen). Dit is aanpasbaar onder Instellingen.

Opmerking over SharePoint
-------------------------
Een publieke GitHub Pages-app kan het beveiligde SharePoint-bestand niet betrouwbaar rechtstreeks uitlezen zonder Microsoft/Entra-authenticatie en CORS-configuratie. Daarom is lokale import gebruikt. Voor live synchronisatie is een interne hosting of Entra-appregistratie nodig.
