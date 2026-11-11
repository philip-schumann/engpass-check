/**
 * Engpass-Check – Backend (Google Apps Script)
 * Eine Zeile pro Lead (über "sid"), wird bei jedem Schritt aktualisiert.
 * E-Mail an Philip, sobald ein Lead seinen Plan hat oder auf Calendly klickt.
 */
const MAIL_TO = "p.schumann.business@gmail.com";
const SHEET_NAME = "Leads";
const SHEET_ID = "1PlRtPbThS2-yVjYgWn9SIMl0NUQoRRBmPKExmnuQ4Qw";
const COLS = ["sid","zeit","status","name","quelle","heiss","typ","ziel","dauer","probiert","essen","wasessen",
  "training","kipppunkt","planung","warum","wichtig","hilfe","geschlecht","alter","groesse","gewicht","kcal","protein",
  "ernaehrung","routine","kopf","ausschluss","seit","bedeutung","anders"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(COLS);
      sh.setFrozenRows(1);
    }
    // Kopfzeile aktuell halten (neue Spalten ergänzen)
    if (sh.getRange(1, COLS.length).getValue() !== COLS[COLS.length - 1]) sh.getRange(1, 1, 1, COLS.length).setValues([COLS]);
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat();
    let row = ids.indexOf(d.sid) + 1;
    const prevStatus = row > 0 ? sh.getRange(row, 3).getValue() : "";
    const values = [COLS.map(c => d[c] !== undefined ? d[c] : "")];
    if (row > 0) sh.getRange(row, 1, 1, COLS.length).setValues(values);
    else { sh.appendRow(values[0]); row = sh.getLastRow(); }

    // Benachrichtigung
    const subjectBase = (d.name || "Unbekannt") + " – " + (d.typ || "");
    if (d.status === "calendly_klick" && prevStatus !== "calendly_klick") {
      MailApp.sendEmail(MAIL_TO, "📅 Calendly-Klick: " + subjectBase, summary(d));
    } else if (d.status === "plan" && d.heiss === "JA") {
      MailApp.sendEmail(MAIL_TO, "🔥 Heißer Lead: " + subjectBase, summary(d));
    }
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

function summary(d) {
  return [
    "Name: " + (d.name || "-"), "Typ: " + d.typ, "Ziel: " + d.ziel,
    "Wichtigkeit (1–10): " + d.wichtig, "Angehen: " + d.hilfe, "Warum jetzt: " + d.warum,
    "Seit wann 'so nicht weiter': " + (d.seit || "-"),
    "Was es bedeuten würde (seine Worte): " + (d.bedeutung || "-"),
    "Dieses Mal anders (seine Worte): " + (d.anders || "-"),
    "Kipppunkt: " + d.kipppunkt, "Planung: " + d.planung,
    "Plan: " + (d.kcal ? d.kcal + " kcal / " + d.protein + " g Protein" : "-"),
    "Quelle/Variante: " + (d.quelle || "-"),
  ].join("\n");
}
