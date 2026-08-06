/**
 * Løbskalender KBH – synkroniserer løb fra fanen "Indsendelser" til en
 * offentlig Google Kalender, OG stiller en simpel, offentlig JSON-liste til
 * rådighed (løbsnavn, dato, distance, sted, tilmeldingslink) til brug på
 * hjemmesidens "Se løbene"-liste.
 *
 * GODKENDELSESLOGIK (kolonnen "Godkendt?"):
 *   - Tom + INGEN mulig dublet (kolonne H er tom) -> offentliggøres automatisk.
 *   - Tom + MULIG DUBLET (kolonne H er flaget)    -> afventer manuel godkendelse.
 *                                                     Scriptet rører IKKE rækken,
 *                                                     før I sætter "Ja" eller "Nej".
 *   - "Ja"  -> tvinges igennem og offentliggøres, uanset dublet-flag.
 *   - "Nej" -> afvises/blokeres. Er den allerede offentliggjort, fjernes den
 *              igen fra kalenderen.
 * Allerede offentliggjorte rækker røres ikke af nye dublet-flag – kun en
 * eksplicit "Nej" kan fjerne noget, der allerede er ude.
 *
 * OPSÆTNING:
 * 1. Udfyld CALENDAR_ID nedenfor.
 * 2. Sæt scriptet ind via Extensions/Udvidelser > Apps Script i dit ark.
 * 3. Kør "createHourlyTrigger" én gang, og godkend adgangene.
 * 4. Deploy som Web App: "Deploy" > "New deployment" > vælg type "Web app" >
 *    "Execute as": Me, "Who has access": Anyone > Deploy. Kopiér URL'en –
 *    den bruges i hjemmesidens "Se løbene"-liste.
 *    (Redigerer du scriptet senere: "Manage deployments" > blyant-ikon >
 *    "New version" > Deploy, for at opdatere den offentlige liste.)
 */

// ---- KONFIGURATION ----
const CALENDAR_ID = 'INDSÆT_DIN_KALENDER_ID_HER'; // fx xxxxxxxxxx@group.calendar.google.com
const SHEET_NAME = 'Indsendelser';

// Kolonnenumre (A=1, B=2, ...) – matcher den faktiske formular/ark-rækkefølge:
// Tidsstempel, Race name, Race date, Your name and email, Link to race,
// Race distance, Race place, Mulig dublet?, Godkendt?, Synkroniseret, Event-ID
const COL = {
  TIMESTAMP: 1,
  NAME: 2,
  DATE: 3,
  SUBMITTER: 4,
  LINK: 5,
  DISTANCES: 6,
  LOCATION: 7,
  DUP_FLAG: 8,  // "Mulig dublet?" – udfyldes af en formel i arket
  APPROVED: 9,  // "Godkendt?" – blank/Ja/Nej, se logik i filens toptekst
  SYNCED: 10,
  EVENT_ID: 11,
};

/**
 * Hovedfunktion: går alle rækker igennem og offentliggør/fjerner efter
 * reglerne beskrevet i toppen af filen. Tjekker desuden selv, om kalenderen
 * allerede har en begivenhed med samme titel/dato, så intet oprettes dobbelt.
 */
function syncApprovedRaces() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Kunne ikke finde fanen "' + SHEET_NAME + '".');

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) throw new Error('Kunne ikke finde kalenderen. Tjek CALENDAR_ID.');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, Object.keys(COL).length);
  const values = range.getValues();

  let created = 0, removed = 0, pending = 0, skippedAsDuplicate = 0;
  const retainedEventIds = {}; // holder styr på alle event-ID'er, der SKAL blive stående

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowNum = i + 2;

    const name = row[COL.NAME - 1];
    const dateValue = row[COL.DATE - 1];
    const existingEventId = row[COL.EVENT_ID - 1];
    const decision = decideStatus(row);

    if (!name || !dateValue) continue; // tom række

    if (existingEventId) {
      // Allerede offentliggjort: rør kun ved den, hvis den eksplicit afvises nu.
      if (decision === 'reject') {
        try {
          const ev = calendar.getEventById(existingEventId);
          if (ev) ev.deleteEvent();
        } catch (err) {
          // Event fandtes ikke længere – ignorer
        }
        sheet.getRange(rowNum, COL.SYNCED).setValue('Afvist / fjernet');
        sheet.getRange(rowNum, COL.EVENT_ID).setValue('');
        removed++;
      } else {
        retainedEventIds[existingEventId] = true;
      }
      continue;
    }

    if (decision === 'reject') continue; // afvist, aldrig offentliggjort
    if (decision === 'pending') { pending++; continue; } // afventer manuel godkendelse

    // decision === 'approve'
    const raceDate = parseRaceDate(dateValue);
    if (!raceDate) {
      sheet.getRange(rowNum, COL.SYNCED).setValue('Fejl: ugyldig dato');
      continue;
    }

    const distances = row[COL.DISTANCES - 1] || '';
    const link = row[COL.LINK - 1] || '';

    // Titel og beskrivelse begrænset til det, der skal vises offentligt:
    // navn, distance og tilmeldingslink. Sted/indsender bruges ikke her.
    const title = distances ? (name + ' (' + distances + ')') : name;
    const description = link ? ('Tilmelding: ' + link) : '';

    const existing = findExistingEvent(calendar, raceDate, title);

    let eventId;
    if (existing) {
      eventId = existing.getId();
      skippedAsDuplicate++;
    } else {
      const event = calendar.createAllDayEvent(title, raceDate, { description: description });
      eventId = event.getId();
      created++;
    }

    sheet.getRange(rowNum, COL.SYNCED).setValue(new Date());
    sheet.getRange(rowNum, COL.EVENT_ID).setValue(eventId);
    retainedEventIds[eventId] = true;
  }

  // Selvhelende oprydning: fjern alle kalender-events, der ikke længere har
  // en tilhørende, gyldig række i arket (fx fordi en række blev SLETTET i
  // stedet for afvist med "Nej" – så bliver eventet ikke stående for evigt).
  const orphansRemoved = removeOrphanedEvents(calendar, retainedEventIds);

  // Opdater den cachede offentlige liste, så hjemmesiden altid får et hurtigt
  // svar uden selv at skulle scanne hele arket ved hvert besøg.
  refreshRacesCache();

  Logger.log('Oprettet: ' + created + ' | Fjernet/afvist: ' + removed +
    ' | Afventer manuel godkendelse: ' + pending +
    ' | Sprunget over som dublet i kalenderen: ' + skippedAsDuplicate +
    ' | Forældreløse events ryddet: ' + orphansRemoved);
}

/**
 * Sletter alle kalenderbegivenheder, hvis event-ID ikke findes i
 * `retainedEventIds` – dvs. begivenheder, hvis tilhørende række er blevet
 * slettet (eller afvist) i arket, uden at begivenheden selv blev fjernet.
 */
function removeOrphanedEvents(calendar, retainedEventIds) {
  const start = new Date();
  start.setMonth(start.getMonth() - 1);
  const end = new Date();
  end.setFullYear(end.getFullYear() + 2);

  const events = calendar.getEvents(start, end);
  let removedCount = 0;
  events.forEach(function (ev) {
    if (!retainedEventIds[ev.getId()]) {
      ev.deleteEvent();
      removedCount++;
    }
  });
  return removedCount;
}

/**
 * Afgør status for en række:
 *  - "reject"  -> "Godkendt?" er sat til "Nej".
 *  - "approve" -> "Godkendt?" er "Ja", ELLER blank uden dublet-flag.
 *  - "pending" -> blank, MEN med et dublet-flag i kolonne H – afventer et menneske.
 */
function decideStatus(row) {
  const approvedRaw = String(row[COL.APPROVED - 1] || '').trim().toLowerCase();
  if (approvedRaw === 'nej') return 'reject';
  if (approvedRaw === 'ja') return 'approve';

  const hasDupFlag = String(row[COL.DUP_FLAG - 1] || '').trim() !== '';
  return hasDupFlag ? 'pending' : 'approve';
}

/** Leder efter en eksisterende begivenhed med (nær-)samme titel på samme dato. */
function findExistingEvent(calendar, date, title) {
  const dayEvents = calendar.getEventsForDay(date);
  const normalizedTitle = title.trim().toLowerCase();
  for (let j = 0; j < dayEvents.length; j++) {
    if (dayEvents[j].getTitle().trim().toLowerCase() === normalizedTitle) {
      return dayEvents[j];
    }
  }
  return null;
}

/** Konverterer celleværdien (Date-objekt eller tekst) til en gyldig Date. */
function parseRaceDate(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value;
  }
  const parsed = new Date(value);
  return isNaN(parsed) ? null : parsed;
}

/**
 * ENGANGS-OPRYDNING: tømmer kalenderen fuldstændigt (alle events, fx gamle
 * test-events der ikke længere matcher en række i arket) og bygger den op
 * igen udelukkende fra det, der står i arket LIGE NU. Ryd op i arket først
 * (slet testrækker, sæt Godkendt?/data korrekt), og kør så denne – ikke
 * omvendt. Kør den fra Apps Script-editoren (vælg funktionen i dropdown'en
 * ved "Kør"-knappen).
 *
 * ADVARSEL: sletter ALT i den kalender, CALENDAR_ID peger på. Brug kun på en
 * kalender, der er dedikeret til løbskalenderen.
 */
function resetAndRebuildCalendar() {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) throw new Error('Kunne ikke finde kalenderen. Tjek CALENDAR_ID.');

  const start = new Date();
  start.setFullYear(start.getFullYear() - 1);
  const end = new Date();
  end.setFullYear(end.getFullYear() + 3);

  const events = calendar.getEvents(start, end);
  events.forEach(function (ev) { ev.deleteEvent(); });
  Logger.log('Slettede ' + events.length + ' events fra kalenderen.');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, COL.SYNCED, lastRow - 1, 1).clearContent();
    sheet.getRange(2, COL.EVENT_ID, lastRow - 1, 1).clearContent();
  }

  syncApprovedRaces();
  Logger.log('Kalenderen er genopbygget fra arkets nuværende indhold.');
}

/**
 * Kør denne funktion ÉN gang (eller igen, hvis I ændrer intervallet) for at
 * sætte et automatisk hjul op, der kalder syncApprovedRaces hvert 10. minut.
 * Har I allerede kørt en ældre version af denne funktion (fx den gamle
 * "createHourlyTrigger"), sletter denne det gamle hjul først, så der ikke
 * kører to trigger-hjul samtidig.
 */
function createHourlyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === 'syncApprovedRaces') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('syncApprovedRaces')
    .timeBased()
    .everyMinutes(10)
    .create();

  Logger.log('Hjul oprettet: syncApprovedRaces kører nu hvert 10. minut.');
}

// Cache-nøgle og -varighed for den offentlige liste. 15 minutter er lidt
// længere end synk-intervallet (10 min), så cachen aldrig når at udløbe
// mellem to automatiske kørsler af syncApprovedRaces.
const RACES_CACHE_KEY = 'races_json_v1';
const RACES_CACHE_SECONDS = 900;

/**
 * Offentlig, read-only JSON-liste over offentliggjorte løb – kun de fire
 * felter, listen på hjemmesiden skal vise: navn, dato, distance, link.
 * Kaldes af hjemmesiden via fetch() til denne Web Apps URL.
 *
 * Svaret hentes fra en cache (opdateret af syncApprovedRaces ved hvert kør),
 * så hjemmesiden ikke skal vente på, at hele arket bliver scannet igen ved
 * hvert besøg – det er det, der gjorde listen langsom at indlæse.
 */
function doGet(e) {
  const cache = CacheService.getScriptCache();
  let json = cache.get(RACES_CACHE_KEY);

  if (!json) {
    json = buildRacesJson();
    cache.put(RACES_CACHE_KEY, json, RACES_CACHE_SECONDS);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/** Genopbygger cachen manuelt – kaldes automatisk fra syncApprovedRaces. */
function refreshRacesCache() {
  const json = buildRacesJson();
  CacheService.getScriptCache().put(RACES_CACHE_KEY, json, RACES_CACHE_SECONDS);
}

/** Scanner arket og bygger selve JSON-strengen med de offentliggjorte løb. */
function buildRacesJson() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const lastRow = sheet ? sheet.getLastRow() : 0;

  const races = [];
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, Object.keys(COL).length).getValues();
    values.forEach(function (row) {
      const name = row[COL.NAME - 1];
      const dateValue = row[COL.DATE - 1];
      const synced = row[COL.SYNCED - 1];
      const approvedRaw = String(row[COL.APPROVED - 1] || '').trim().toLowerCase();

      if (!name || !dateValue) return;
      if (approvedRaw === 'nej') return; // afvist, selv hvis den var offentliggjort tidligere
      if (!synced || !(synced instanceof Date)) return; // ikke offentliggjort / afventer / fejlet

      const raceDate = parseRaceDate(dateValue);
      if (!raceDate) return;

      races.push({
        navn: name,
        dato: Utilities.formatDate(raceDate, 'Europe/Copenhagen', 'yyyy-MM-dd'),
        distance: row[COL.DISTANCES - 1] || '',
        sted: row[COL.LOCATION - 1] || '',
        link: row[COL.LINK - 1] || '',
      });
    });
  }

  races.sort(function (a, b) { return a.dato < b.dato ? -1 : a.dato > b.dato ? 1 : 0; });
  return JSON.stringify(races);
}
