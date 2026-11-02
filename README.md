# Docent Tools voor Teams (v0.5.0)

Chrome-extensie voor Microsoft Teams (web) voor docenten: team-overstijgende analyses van opdrachten. Alles draait lokaal in je browser; er wordt niets naar een eigen server gestuurd. Zusje van [Docent Tools voor Magister](https://github.com/mark-jg/magister-docent), met dezelfde opzet.

## Installeren (lokaal laden)

1. Download de nieuwste `teams-docent-tools-vX.Y.Z.zip` van de Releases-pagina, of kloon deze map.
2. Pak de zip uit in een map die je laat staan (bijvoorbeeld `Documenten/teams-docent-tools`).
3. Open Chrome en ga naar `chrome://extensions`.
4. Zet rechtsboven **Ontwikkelaarsmodus** aan.
5. Klik op **Uitgepakte extensie laden** en kies de map met `manifest.json`.
6. Open (of herlaad) Teams in de browser: `https://teams.microsoft.com` of `https://teams.cloud.microsoft`. Rechts verschijnt een handvat "Docent Tools".

Werkt ook in Edge (`edge://extensions`). Na elke update of herlaad van de extensie moet een open Teams-tabblad één keer herladen worden (F5); de extensie meldt dat zelf. Bijwerken gaat zoals bij de Magister-extensie: de extensie kijkt om de 6 uur op GitHub of er een nieuwere release is en kan zichzelf bijwerken nadat je één keer de extensiemap hebt aangewezen (Instellingen > Updates).

## Wat zit erin

| Functie | Hoe |
|---|---|
| **Opdrachten over alle teams**: per klas (afgeleid uit wie samen in dezelfde teams zit), per team, per leerling (over al je teams heen), per opdracht, komende deadlines, nakijkwerk en signalen | Alt+O, of via het zijpaneel of het extensie-icoon |
| **Cijfers**: per leerling het gemiddelde per team, gewogen volgens de cijfercategorieën van dat team, plus een totaal; klik voor categorieën en opdrachten | tabblad Cijfers |
| **Rubrics**: rubric-rijen die in meerdere opdrachten of teams voorkomen, met per leerling het gekozen niveau over al die voorkomens | tabblad Rubrics |
| **Rapporten**: per leerling een A4-rapport met cijfers per vak per periode, studentnummer, klas, opleiding, cohort en leerjaar, afdrukbaar (ook naar PDF) in de huisstijl van Mediacollege Amsterdam; opmerkingen per periode typ je in het rapport zelf | knop Rapporten in het overzicht |
| Export naar Excel: het geopende tabblad of alles in één werkmap (zes bladen, inclusief alle inleveringen) | knoppen bovenin het overzicht |
| Filters: klassen (automatisch of zelf gekozen, wordt onthouden), periode (schooljaar, 90 dagen, 30 dagen, alles), "Alleen problemen", zoeken; sorteren door op een kolomkop te klikken | bovenin het overzicht |
| Zijpaneel met status van de koppeling, modules aan/uit en thema | Alt+K of het handvat |
| Inspector (ontwikkelaarsmodus): frames, tokens (alleen host en scopes), API-calls van Teams | zijpaneel > Ontwikkelaar |

### De analyses

Per klas en per leerling wordt geteld: toegewezen, ingeleverd (en of dat op tijd of **te laat** was), **ontbreekt** (deadline voorbij, niets ingeleverd), open (deadline nog niet voorbij), nagekeken, na te kijken, vrijgesteld, en het gemiddelde percentage van de maximale punten over de nagekeken opdrachten.

- **Leerlingen** is de team-overstijgende kern: een leerling die in meerdere van je klassen zit, staat één keer, met de optelsom over die klassen. Klik op een rij voor de uitsplitsing per klas en de lijst van opdrachten.
- **Deadlines** zet de deadlines van al je klassen op één tijdlijn, zodat je ziet waar leerlingen die je in meerdere vakken hebt tegelijk mee te maken krijgen.
- **Nakijken** laat zien wat er over alle klassen op je ligt te wachten, langst wachtende bovenaan.
- **Signalen** zijn leerlingen die over alle klassen opgeteld boven een drempel komen (standaard 2× ontbreekt of 3× te laat; instelbaar). De kolom "In klassen" laat zien of het in één vak speelt of breder.

### Klassen bij een team per vak

Op scholen met een team per vak is de klas niet één team. De extensie leidt klassen daarom af uit **wie samen in dezelfde teams zit**: twee teams horen bij dezelfde klas als het kleinste team voor minstens de helft uit dezelfde leerlingen bestaat (en minstens drie gemeenschappelijk). Elke leerling krijgt de klas waarvan die de meeste teams volgt; een leerling die in 4A zit en één keuzevak met 4B deelt, blijft dus 4A. De klas krijgt als naam het woord dat de teamnamen gemeen hebben, bij voorkeur met een cijfer erin (bijvoorbeeld "4A" uit "Engels 4A" en "Wiskunde 4A").

Standaard wordt de **klascode vóór het streepje** in de teamnaam gebruikt: "MIP26A - Burgerschap" en "MIP26A - Rekenen" vormen samen klas MIP26A (patroon: letters, jaar of cohort, optioneel een letter). Teams zonder herkenbare code, zoals "KD - Verdieping Beeldgebruik", sluiten via de overlap-regel aan bij de klas waarmee ze de meeste leerlingen delen. Een ander patroon (regex met één groep) kan in de instellingen; `^$` schakelt de code uit. Klassen zonder leerlingen (bijvoorbeeld een team dat door twee klassen wordt gedeeld) staan niet in de keuzelijst. Het tabblad **Klassen** toont per klas de teams, de klaslijst (klik op de rij) en de totalen; de keuzelijst **Klas** bovenin beperkt alle tabbladen tot één klas; de kolom Klas staat ook bij Leerlingen en Signalen en in de Excel-bladen Klassen en Klasindeling.

### Welke teams tellen mee

Standaard bepaalt de extensie dat automatisch met vier signalen uit Graph: is het een **klasteam** (teamtype Klas), is het **gearchiveerd**, ben jij **eigenaar**, en hoort het bij **dit schooljaar** (team aangemaakt sinds juni vóór het schooljaar, of opdrachten met een datum in dit schooljaar). Alleen eigen, actieve klasteams van dit jaar tellen mee; bovenin staat hoeveel er verborgen zijn en waarom. Via **Klassen** kies je zelf (zoekveld, vinkjes, Alles, Niets); die keuze wordt onthouden tot je weer op Automatisch klikt. Gearchiveerde teams worden standaard niet eens opgehaald (instelbaar), en opdrachten van vóór het vorige schooljaar vallen weg.

**Alleen problemen** laat per tabblad alleen rijen zien met ontbrekend werk, te laat of nakijkwerk; de tellingen bovenin veranderen niet mee. Standaardvolgorde: Leerlingen en Signalen op ontbrekend werk, Opdrachten op deadline (nieuwste eerst), Deadlines op datum, Nakijken op wachttijd.

### Cijfers en rubrics

**Cijfers** toont per leerling en per team het gemiddelde als percentage van de maximale punten. Heeft het team cijfercategorieën met een weging (Opdrachten › Instellingen in Teams), dan wordt het gemiddelde per categorie berekend en gewogen opgeteld, genormaliseerd over de categorieën waarin de leerling cijfers heeft; opdrachten zonder categorie staan wel in de lijst maar tellen niet mee in de weging. Zonder wegingen geldt het gewone gemiddelde. Totaal is het gemiddelde over de teams. Rood onder 55%, oranje onder 70%.

**Rubrics** haalt per opdracht de rubric op en herkent rijen (criteria) op hun tekst, zonder leestekens en hoofdletters. Rijen die in meerdere opdrachten of teams voorkomen zijn "gedeeld": klik erop en je ziet per leerling het gekozen niveau in elke opdracht, met een gemiddelde (punten van het niveau gedeeld door de hoogste punten in de rubric; zonder punten de positie van het niveau). Zo is een rij als "Samenwerken" over vakken heen te volgen. Het Excel-blad Rubric-niveaus per leerling bevat alles.

### Rapporten

De knop **Rapporten** opent een keuzevenster: klassen, periodes en afronding. **Rapporten openen** maakt een afdrukbare pagina met één A4 per leerling; **Naar Excel** geeft dezelfde cijfers als tabel. Op het rapport staan:

- Opleiding en titel (bijvoorbeeld "Resultaten jaar 1"), met het logo van Mediacollege Amsterdam.
- Studentnummer (de cijfers vóór de @ van het e-mailadres), roepnaam, achternaam, klas, cohort en leerjaar. Cohort en leerjaar komen uit de klascode: MIP26A is cohort 2026, en in schooljaar 2026-2027 leerjaar 1.
- Per vak (teamnaam zonder klascode) het cijfer per periode: het gewogen gemiddelde van de opdrachten met een deadline in die periode, omgerekend naar een cijfer met de cesuur (standaard 55% = 5,5, lineair naar 1 en 10). Onvoldoendes in magenta. Vakken zijn te groeperen (bijvoorbeeld Praktijkvakken en AVO-vakken).
- Een vaste toelichting en per periode een vak voor opmerkingen, dat je in het rapport typt; opmerkingen worden per leerling en periode bewaard.

Periodes, opleidingen, vakgroepen, cesuur, afronding, titel en toelichting staan bij Instellingen › Rapporten. De opmaak volgt het Mediacollege Design System: wit, zwart en magenta als enig accent, dunne hairlines, tekstkaders; Aeonik is een gelicentieerd lettertype, dus de pagina gebruikt de voorgeschreven vervanging Arial (en Georgia voor de titel). Het logo is ook een merkbestand en zit niet in deze repo: kies het bij Instellingen › Rapporten (blijft lokaal), of zet het officiële horizontale logo als `rapport/ma-logo.jpg` in de extensiemap; anders staat het woordmerk als tekst. Aanwezigheid staat niet in Teams en dus niet op het rapport.

Concept- en ingeplande opdrachten tellen niet mee. Vrijstelling (excused) telt niet als ontbrekend en niet mee in het gemiddelde.

## Hoe het werkt

De extensie gebruikt de **Microsoft Graph education-API** (`/education/classes/{id}/assignments` en `.../submissions?$expand=outcomes`) met het token dat de Teams-webclient zelf al heeft. In Teams zit het bruikbare token (scope `EduAssignments.ReadWrite`) alleen in het frame van de **app Opdrachten** (`assignments.edu.cloud.microsoft`); de hoofdpagina heeft geen Graph-token met education-scopes. Open daarom in Teams één keer per sessie de app Opdrachten (linkerbalk, eventueel via ⋯ Meer apps); de extensie wacht daar tot anderhalve minuut op en zegt dat ook.

Dat token mist `EduRoster`. Daarom gaat de klassenlijst niet via `/education/me/classes` (geeft 403) maar via je teams-lijst (`/me/joinedTeams`); een team is een klas als de opdrachten-aanroep slaagt. Leden komen uit de groepsleden minus eigenaren (`/groups/{id}/members` en `/owners`); namen van leerlingen die niet (meer) lid zijn via `$batch` op `/users`. Werkt `/education/me/classes` op jouw school wél, dan wordt die gewoon gebruikt.

Een script in de pagina (`content/netlog.js`) onthoudt per host het Bearer-token dat Teams meestuurt, en valt terug op de MSAL-cache van Teams in localStorage. Het token blijft in het geheugen van de pagina; de extensie krijgt alleen de antwoorden. Er worden alleen leesaanroepen gedaan (GET, en POST naar `$batch` voor namen).

De opgehaalde teams, opdrachten en inleveringen (alleen de velden die de analyse nodig heeft) blijven in `chrome.storage.local`, zodat het overzicht de volgende keer direct opent. Ze blijven ook bewaard bij een update van de extensie, zolang die in dezelfde map blijft staan (zelf bijwerken doet dat; een zip uitpakken over dezelfde map ook). Laad je de extensie uit een nieuwe map, dan ziet Chrome dat als een andere extensie met lege opslag.

Het ophalen is gebundeld: aanroepen gaan per twintig tegelijk via Graph `$batch`, en teams waarvan je geen eigenaar bent worden niet doorzocht (instelbaar), net als gearchiveerde teams. Loopt er al een ronde, dan sluit een tweede klik daarop aan in plaats van te wachten.

**Bijwerken** is incrementeel: teams, opdrachten en leden worden opnieuw opgehaald, maar inleveringen van opdrachten met een deadline ouder dan 45 dagen worden hergebruikt zolang de opdracht niet is gewijzigd. Dat scheelt het grootste deel van de aanroepen. Hoeveel schooljaren worden opgehaald staat in de instellingen (standaard dit en vorig schooljaar; minder is sneller), en de periodekeuze in het overzicht volgt dat. Grote tabellen tonen eerst 200 rijen (sorteren gebeurt over alles); met "Nog 500" of "Alles" zie je de rest. Standaard gebeurt bijwerken automatisch na 24 uur, op de achtergrond terwijl het oude overzicht in beeld blijft; eens per 7 dagen wordt toch alles opgehaald. Alt+klik op Bijwerken haalt alles direct opnieuw op; beide termijnen staan in de instellingen.

### Als de koppeling niet werkt

Het zijpaneel toont onder **Koppeling** wat er aan de hand is:

- *Nog geen Graph-token gezien* of *zonder EduAssignments*: de app Opdrachten is in deze sessie nog niet open geweest. Open die in Teams en controleer opnieuw.
- *Koppeling werkt via de teams-lijst*: normaal op scholen waar het token geen EduRoster heeft; alles werkt, alleen de klassenlijst en de leden komen uit de Teams-groepen.
- Faalt het toch, zet dan de ontwikkelaarsmodus aan, open de Inspector en kopieer het rapport (geanonimiseerd: frames, tokens met alleen host en scopes, API-calls met alleen veldnamen). In dat rapport is ook de interne API van de app Opdrachten te zien (`assignments.edu.cloud.microsoft/api/v1.0/edu/me/work`), een mogelijke tweede route.

## Privacy

- Geen eigen server, geen telemetrie. De enige externe aanroepen zijn naar Microsoft Graph (binnen je eigen sessie) en, voor updates, naar GitHub (alleen versienummer en downloadlink; uit te zetten).
- Tokens verlaten de pagina niet. Het Inspector-rapport bevat geen tokens en geen waarden: id's, GUID's en e-mailadressen worden vervangen, van API-antwoorden blijven alleen veldnamen over.
- Leerlinggegevens staan alleen in de lokale opslag van de extensie op jouw apparaat en in Excel-bestanden die je zelf exporteert.

## Ontwikkelen

- `node test/analyse.test.js` test de berekeningen (`content/analyse.js`) zonder browser.
- Nieuwe analyses: `content/analyse.js` (pure functies), weergave in `content/modules/opdrachten.js`.
- Nieuwe versie: regel toevoegen onder Wijzigingen, dan `scripts/release.sh X.Y.Z` (zet versie, tagt, pusht; GitHub Actions bouwt de zip en de release).

## Wijzigingen

- 0.5.0: rapporten per leerling (cijfers per vak per periode, studentnummer uit het e-mailadres, klas, opleiding, cohort, leerjaar, opmerkingen per periode), afdrukbaar in de huisstijl van Mediacollege Amsterdam en naar Excel; instellingen voor periodes, opleidingen, vakgroepen, cesuur, afronding, titel en toelichting; e-mailadres en naamdelen van leerlingen worden meegehaald (eerste keer bijwerken haalt leden opnieuw op).
- 0.4.0: tabblad Cijfers (gemiddelde per leerling per team, gewogen volgens de cijfercategorieën, met totaal en detail per categorie) en tabblad Rubrics (gelijke rubric-rijen over opdrachten en teams heen, met per leerling het gekozen niveau); Excel-bladen Cijfers, Cijfers per categorie, Rubric-rijen en Rubric-niveaus per leerling. Eerste keer bijwerken haalt alles opnieuw op.
- 0.3.5: klascode vóór het streepje in de teamnaam is nu het standaardpatroon (MIP26A, MR26B, KD25); teams zonder code sluiten via overlap bij elke klas aan, ook een gecodeerde; de klasnaam is de meest voorkomende code in de groep; klassen zonder leerlingen staan niet meer in de keuzelijst en het tabblad.
- 0.3.4: tabwissel en filters werkten niet meer in een Teams-tabblad dat open stond tijdens een update van de extensie (het opslaan van de keuze gooide een fout); dat opslaan blokkeert nu nooit meer, de extensie meldt dat het tabblad herladen moet worden, en een fout in het overzicht wordt getoond met kopieerbare details.
- 0.3.3: ophalen veel sneller (aanroepen gebundeld per twintig, teams zonder eigenaarschap worden overgeslagen, instelbaar); een tweede klik tijdens het ophalen sluit aan op de lopende ronde in plaats van te blijven hangen op "Verbinden"; voortgang ook in het zijpaneel; melding als opslaan mislukt.
- 0.3.2: tabwissel is weer vlot (analyse wordt één keer per filtercombinatie berekend, tabellen in porties van 200 rijen); instelling hoeveel schooljaren worden opgehaald (1 tot 4), met bijpassende periodekeuze.
- 0.3.1: opgehaalde gegevens blijven bruikbaar na een update van de extensie (aanvullen in plaats van opnieuw ophalen); Bijwerken is incrementeel (inleveringen van oude opdrachten worden hergebruikt) en gebeurt op de achtergrond terwijl het overzicht open blijft; automatisch bijwerken na 24 uur, alles opnieuw eens per 7 dagen; onbeperkte lokale opslag.
- 0.3.0: klassen afgeleid uit lidmaatschap (wie samen in dezelfde teams zit) en optioneel uit een klascode in de teamnaam; nieuw tabblad Klassen met klaslijsten, kolom Klas bij Leerlingen en Signalen, keuzelijst Klas als filter, Excel-bladen Klassen en Klasindeling; het oude tabblad Klassen heet nu Teams.
- 0.2.0: klassen automatisch bepaald (klasteam, niet gearchiveerd, eigenaar, dit schooljaar) met een klassenkiezer die de keuze onthoudt; schakelaar Alleen problemen; betere standaardvolgorde per tabblad; filters en tabblad worden onthouden; gearchiveerde teams en opdrachten van vóór het vorige schooljaar worden niet meer opgehaald.
- 0.1.1: werkt met het Graph-token van de app Opdrachten (EduAssignments zonder EduRoster): klassen via de teams-lijst, leden via groepsleden; de extensie wacht tot de app Opdrachten open is en kiest het juiste frame op scopes; status in het zijpaneel legt uit wat er mist.
- 0.1.0: eerste versie. Opdrachten over alle teams (klassen, leerlingen, opdrachten, deadlines, nakijken, signalen) via Microsoft Graph met het token van Teams zelf; Excel-export; zijpaneel met koppelingsstatus; Inspector; zelf bijwerken via GitHub Releases.
