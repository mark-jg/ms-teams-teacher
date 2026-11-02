# Docent Tools voor Teams (v0.1.0)

Chrome-extensie voor Microsoft Teams (web) voor docenten: team-overstijgende analyses van opdrachten. Alles draait lokaal in je browser; er wordt niets naar een eigen server gestuurd. Zusje van [Docent Tools voor Magister](https://github.com/mark-jg/magister-docent), met dezelfde opzet.

## Installeren (lokaal laden)

1. Download de nieuwste `teams-docent-tools-vX.Y.Z.zip` van de Releases-pagina, of kloon deze map.
2. Pak de zip uit in een map die je laat staan (bijvoorbeeld `Documenten/teams-docent-tools`).
3. Open Chrome en ga naar `chrome://extensions`.
4. Zet rechtsboven **Ontwikkelaarsmodus** aan.
5. Klik op **Uitgepakte extensie laden** en kies de map met `manifest.json`.
6. Open (of herlaad) Teams in de browser: `https://teams.microsoft.com` of `https://teams.cloud.microsoft`. Rechts verschijnt een handvat "Docent Tools".

Werkt ook in Edge (`edge://extensions`). Bijwerken gaat zoals bij de Magister-extensie: de extensie kijkt om de 6 uur op GitHub of er een nieuwere release is en kan zichzelf bijwerken nadat je één keer de extensiemap hebt aangewezen (Instellingen > Updates).

## Wat zit erin

| Functie | Hoe |
|---|---|
| **Opdrachten over alle teams**: per klas, per leerling (over al je klassen heen), per opdracht, komende deadlines, nakijkwerk en signalen | Alt+O, of via het zijpaneel of het extensie-icoon |
| Export naar Excel: het geopende tabblad of alles in één werkmap (zes bladen, inclusief alle inleveringen) | knoppen bovenin het overzicht |
| Filters: periode (schooljaar, 90 dagen, 30 dagen, alles), klassen (klik een chip; Alt+klik toont alleen die klas), zoeken | bovenin het overzicht |
| Zijpaneel met status van de koppeling, modules aan/uit en thema | Alt+K of het handvat |
| Inspector (ontwikkelaarsmodus): frames, tokens (alleen host en scopes), API-calls van Teams | zijpaneel > Ontwikkelaar |

### De analyses

Per klas en per leerling wordt geteld: toegewezen, ingeleverd (en of dat op tijd of **te laat** was), **ontbreekt** (deadline voorbij, niets ingeleverd), open (deadline nog niet voorbij), nagekeken, na te kijken, vrijgesteld, en het gemiddelde percentage van de maximale punten over de nagekeken opdrachten.

- **Leerlingen** is de team-overstijgende kern: een leerling die in meerdere van je klassen zit, staat één keer, met de optelsom over die klassen. Klik op een rij voor de uitsplitsing per klas en de lijst van opdrachten.
- **Deadlines** zet de deadlines van al je klassen op één tijdlijn, zodat je ziet waar leerlingen die je in meerdere vakken hebt tegelijk mee te maken krijgen.
- **Nakijken** laat zien wat er over alle klassen op je ligt te wachten, langst wachtende bovenaan.
- **Signalen** zijn leerlingen die over alle klassen opgeteld boven een drempel komen (standaard 2× ontbreekt of 3× te laat; instelbaar). De kolom "In klassen" laat zien of het in één vak speelt of breder.

Concept- en ingeplande opdrachten tellen niet mee. Vrijstelling (excused) telt niet als ontbrekend en niet mee in het gemiddelde.

## Hoe het werkt

De extensie gebruikt de **Microsoft Graph education-API** (`/education/me/classes`, `/education/classes/{id}/assignments`, `.../submissions?$expand=outcomes`, `.../members`), met het token dat de Teams-webclient zelf al heeft. Een script in de pagina (`content/netlog.js`) onthoudt per host het Bearer-token dat Teams meestuurt, en valt terug op de MSAL-cache van Teams in localStorage. Het token blijft in het geheugen van de pagina; de extensie krijgt alleen de antwoorden. Er worden alleen leesaanroepen gedaan (GET, en POST naar `$batch` voor namen).

De opgehaalde klassen, opdrachten en inleveringen (alleen de velden die de analyse nodig heeft) blijven in `chrome.storage.local`, zodat het overzicht de volgende keer direct opent. Standaard wordt na 6 uur opnieuw opgehaald; "Vernieuwen" kan altijd; wissen kan in de instellingen.

### Als de koppeling niet werkt

Het zijpaneel toont onder **Koppeling** wat er aan de hand is:

- *Nog geen Graph-token gezien*: Teams heeft in deze pagina nog geen aanroep naar Graph gedaan. Klik op Teams of op de app Opdrachten en controleer opnieuw.
- *Graph-token gevonden maar zonder toegang tot de education-API (403)*: het token van de Teams-client heeft de scopes `EduAssignments`/`EduRoster` niet. Zet dan de ontwikkelaarsmodus aan, open de Inspector en kopieer het rapport; daarin staat welke tokens en API's er wél zijn (bijvoorbeeld de interne API van de app Opdrachten), zodat we die route kunnen inbouwen.

Dit is de belangrijkste onzekerheid van deze eerste versie: de documentatie beschrijft de education-API, maar of het token van de Teams-webclient die mag aanroepen, blijkt pas in de praktijk op jouw school.

## Privacy

- Geen eigen server, geen telemetrie. De enige externe aanroepen zijn naar Microsoft Graph (binnen je eigen sessie) en, voor updates, naar GitHub (alleen versienummer en downloadlink; uit te zetten).
- Tokens verlaten de pagina niet. Het Inspector-rapport bevat geen tokens en geen waarden: id's, GUID's en e-mailadressen worden vervangen, van API-antwoorden blijven alleen veldnamen over.
- Leerlinggegevens staan alleen in de lokale opslag van de extensie op jouw apparaat en in Excel-bestanden die je zelf exporteert.

## Ontwikkelen

- `node test/analyse.test.js` test de berekeningen (`content/analyse.js`) zonder browser.
- Nieuwe analyses: `content/analyse.js` (pure functies), weergave in `content/modules/opdrachten.js`.
- Nieuwe versie: regel toevoegen onder Wijzigingen, dan `scripts/release.sh X.Y.Z` (zet versie, tagt, pusht; GitHub Actions bouwt de zip en de release).

## Wijzigingen

- 0.1.0: eerste versie. Opdrachten over alle teams (klassen, leerlingen, opdrachten, deadlines, nakijken, signalen) via Microsoft Graph met het token van Teams zelf; Excel-export; zijpaneel met koppelingsstatus; Inspector; zelf bijwerken via GitHub Releases.
