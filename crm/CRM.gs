/**
 * Schumann CRM – Setting & Erstgespräche (Google Apps Script Web-App, nur für Philip)
 * Datenquelle: Tabelle "Freebie-Leads" (dieselbe wie das Freebie-Backend), Blatt "Leads".
 * Logik nach Akquise.de-Kurs "CRM + Technik Setup": Lead-Status + Smart Views (heiß → kalt).
 *
 * Einrichtung: setup() einmal ausführen (Tabellen-ID eintragen, Kalender-Abgleich stündlich),
 * dann Bereitstellen → Web-App (Ausführen als: ich, Zugriff: nur ich).
 */
const SHEET_ID = "__SHEET_ID__";
const TZ = "Europe/Berlin";
const COLS = ["token","angemeldet","status","name","email","telefon","telefon_ohne_einwilligung","quelle",
  "einwilligung_mail","einwilligung_tel","bestaetigt","nurture_schritt","letzte_mail","abgemeldet","freebie",
  "crm_status","versuche","letzter_kontakt","naechster_schritt","erstgespraech","notizen"];
const STATUS = ["Warm","Nicht erreicht","Im Gespräch","Erstgespräch vereinbart","EG No-Show/Absage","EG Follow-up",
  "Kunde","Ex-Kunde","UQ andere","UQ Bad Data"];

function setup() {
  ensureHeader();
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "syncKalender") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("syncKalender").timeBased().everyHours(1).create();
  syncKalender();
}
function sh() { return SpreadsheetApp.openById(SHEET_ID).getSheetByName("Leads"); }
function ensureHeader() { const s = sh(); COLS.forEach((c, i) => { if (s.getRange(1, i + 1).getValue() !== c) s.getRange(1, i + 1).setValue(c); }); }
function ci(n) { return COLS.indexOf(n); }
function iso(v) { return v instanceof Date ? v.toISOString() : (v || ""); }

function doGet() {
  return HtmlService.createHtmlOutput(PAGE).setTitle("Schumann CRM")
    .addMetaTag("viewport", "width=device-width, initial-scale=1").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/* ---------- Daten fürs Frontend ---------- */
function getData() {
  ensureHeader();
  const v = sh().getDataRange().getValues(); const out = [];
  for (let r = 1; r < v.length; r++) {
    const row = v[r]; if (!row[ci("token")]) continue;
    const o = {}; COLS.forEach((c, i) => o[c] = iso(row[i]));
    if (!o.crm_status) o.crm_status = "Warm";
    o.versuche = Number(o.versuche) || 0;
    out.push(o);
  }
  return {leads: out, now: new Date().toISOString()};
}

/* ---------- Aktionen ---------- */
function act(token, action, payload) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const s = sh(); const v = s.getDataRange().getValues();
    for (let r = 1; r < v.length; r++) if (v[r][ci("token")] === token) {
      const set = (c, val) => s.getRange(r + 1, ci(c) + 1).setValue(val);
      const now = new Date(); const p = payload || {};
      const note = (txt) => set("notizen", (Utilities.formatDate(now, TZ, "dd.MM. HH:mm") + " – " + txt + (p.notiz ? ": " + p.notiz : "")) + (v[r][ci("notizen")] ? "\n" + v[r][ci("notizen")] : ""));
      const tries = (Number(v[r][ci("versuche")]) || 0);
      if (action === "nichtErreicht") {
        set("versuche", tries + 1); set("letzter_kontakt", now);
        set("crm_status", tries + 1 >= 5 ? "Nicht erreicht" : "Warm");
        set("naechster_schritt", tries + 1 >= 5 ? "" : nextSlot(tries + 1));
        note("Nicht erreicht (Versuch " + (tries + 1) + ")");
      } else if (action === "imGespraech") {
        set("crm_status", "Im Gespräch"); set("letzter_kontakt", now);
        set("naechster_schritt", p.datum ? new Date(p.datum) : nextSlot(2)); note("Erreicht – im Gespräch");
      } else if (action === "egGebucht") {
        set("crm_status", "Erstgespräch vereinbart"); set("letzter_kontakt", now);
        if (p.datum) set("erstgespraech", new Date(p.datum)); set("naechster_schritt", ""); note("Erstgespräch vereinbart");
      } else if (action === "egErgebnis") {
        set("crm_status", p.status); set("letzter_kontakt", now);
        set("naechster_schritt", p.datum ? new Date(p.datum) : ""); note("Erstgespräch: " + p.status);
      } else if (action === "status") {
        set("crm_status", p.status); set("letzter_kontakt", now);
        if (p.datum !== undefined) set("naechster_schritt", p.datum ? new Date(p.datum) : ""); note("Status → " + p.status);
      } else if (action === "notiz") {
        note("Notiz");
      } else if (action === "wiedervorlage") {
        set("naechster_schritt", p.datum ? new Date(p.datum) : ""); note("Wiedervorlage " + (p.datum ? Utilities.formatDate(new Date(p.datum), TZ, "dd.MM. HH:mm") : "gelöscht"));
      }
      return getData();
    }
    throw new Error("Lead nicht gefunden");
  } finally { lock.releaseLock(); }
}
// Nächster Anrufversuch: wechselnde Zeiten im Anruffenster (Werktag 14–20, Wochenende 9–21)
function nextSlot(n) {
  const d = new Date(); d.setDate(d.getDate() + (n <= 1 ? 1 : n <= 3 ? 1 : 2));
  const dow = Number(Utilities.formatDate(d, TZ, "u")) % 7; const we = dow === 0 || dow === 6;
  const hour = we ? (n % 2 ? 11 : 17) : (n % 2 ? 17 : 14);
  const s = Utilities.formatDate(d, TZ, "yyyy-MM-dd") + "T" + String(hour).padStart(2, "0") + ":00:00";
  return new Date(Utilities.parseDate(s, TZ, "yyyy-MM-dd'T'HH:mm:ss"));
}
function addLead(o) {
  ensureHeader();
  const row = {token: Utilities.getUuid(), angemeldet: new Date(), status: "manuell", name: String(o.name || "").slice(0, 60),
    email: String(o.email || "").trim().toLowerCase(), telefon: String(o.telefon || "").replace(/[^\d+ ]/g, ""),
    quelle: String(o.quelle || "manuell"), freebie: "manuell", crm_status: "Warm", versuche: 0,
    notizen: o.notiz ? Utilities.formatDate(new Date(), TZ, "dd.MM. HH:mm") + " – " + o.notiz : ""};
  sh().appendRow(COLS.map(c => row[c] === undefined ? "" : (c === "telefon" && row[c] ? "'" + row[c] : row[c])));
  return getData();
}

/* ---------- Calendly-Abgleich über Google Kalender ---------- */
function syncKalender() {
  const s = sh(); const v = s.getDataRange().getValues();
  const byMail = {}; for (let r = 1; r < v.length; r++) { const e = String(v[r][ci("email")] || "").toLowerCase(); if (e) byMail[e] = r; }
  const now = new Date(); const from = new Date(now.getTime() - 14 * 864e5), to = new Date(now.getTime() + 30 * 864e5);
  const evs = CalendarApp.getDefaultCalendar().getEvents(from, to);
  evs.forEach(ev => {
    const mails = ev.getGuestList().map(g => g.getEmail().toLowerCase());
    (String(ev.getDescription() || "").match(/[\w.+-]+@[\w-]+\.[\w.]+/g) || []).forEach(m => mails.push(m.toLowerCase()));
    mails.forEach(m => {
      const r = byMail[m]; if (r === undefined) return;
      const st = String(v[r][ci("crm_status")] || "Warm");
      if (["Kunde","Ex-Kunde","UQ andere","UQ Bad Data"].indexOf(st) >= 0) return;
      const start = ev.getStartTime(); const cur = v[r][ci("erstgespraech")];
      if (cur instanceof Date && Math.abs(cur - start) < 6e4 && st === "Erstgespräch vereinbart") return;
      if (start < now && st !== "Warm" && st !== "Im Gespräch" && st !== "Nicht erreicht") return;
      s.getRange(r + 1, ci("erstgespraech") + 1).setValue(start);
      s.getRange(r + 1, ci("crm_status") + 1).setValue("Erstgespräch vereinbart");
      s.getRange(r + 1, ci("notizen") + 1).setValue(Utilities.formatDate(new Date(), TZ, "dd.MM. HH:mm") + " – Erstgespräch aus Kalender erkannt (" +
        Utilities.formatDate(start, TZ, "dd.MM. HH:mm") + ")" + (v[r][ci("notizen")] ? "\n" + v[r][ci("notizen")] : ""));
    });
  });
}
function syncJetzt() { syncKalender(); return getData(); }

const PAGE = String.raw`__PAGE__`;
