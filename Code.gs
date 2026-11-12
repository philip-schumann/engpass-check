/**
 * Engpass-Check – Backend (Google Apps Script)
 * Eine Zeile pro Lead (über "sid"), wird bei jedem Schritt aktualisiert.
 * E-Mails an Philip: heißer Lead (nach Eintragen der Kontaktdaten) und Calendly-Klick.
 * E-Mails an den Lead: Ergebnis (nach Eintragen) und Ernährungsplan (sobald erstellt).
 */
const MAIL_TO = "p.schumann.business@gmail.com";
const SHEET_NAME = "Leads";
const SHEET_ID = "1PlRtPbThS2-yVjYgWn9SIMl0NUQoRRBmPKExmnuQ4Qw";
const CHECK_URL = "https://philip-schumann.github.io/engpass-check/";
const CALENDLY = "https://calendly.com/philip-schumann/erstgespraech";
const COLS = ["sid","zeit","status","name","quelle","heiss","typ","ziel","dauer","probiert","essen","wasessen",
  "training","kipppunkt","planung","warum","wichtig","hilfe","geschlecht","alter","groesse","gewicht","kcal","protein",
  "ernaehrung","routine","kopf","ausschluss","seit","bedeutung","anders","email","telefon","einwilligung",
  "aktivitaet","ernaehrungsform","mail_ergebnis","mail_plan"];

// Texte nur serverseitig – der Inhalt der Mails kommt nie ungeprüft aus dem Formular.
const TYPES = {
  ernaehrung: {name: "Der Ernährungs-Blindflieger",
    text: "Du bist motiviert, aber dir fehlt die Klarheit, was und wie viel du konkret essen sollst. Ohne diese Basis ist jeder Tag ein Ratespiel.",
    hebel: "Feste Kalorien- und Proteinzahl plus 3–4 Standard-Mahlzeiten, die du immer wieder isst.",
    task: "Wähle 2 Frühstücke und 2 Mittagessen und iss diese Woche nur diese. Fokus: jede Mahlzeit mit Protein-Quelle."},
  routine: {name: "Der Struktur-Sucher",
    text: "Wissen ist nicht dein Problem – Planung ist es. Sobald der Alltag stressig wird, fällt alles weg, weil nichts vorbereitet ist.",
    hebel: "Ein fester Wochenrhythmus: Planungs- und Einkaufstag, feste Trainingstage im Kalender und ein Notfall-Essen für stressige Tage.",
    task: "Nimm dir diesen Sonntag 20 Minuten: Trainingstage eintragen, Einkaufsliste schreiben, 2 Mahlzeiten vorkochen."},
  training: {name: "Der Trainings-Zufall",
    text: "Du bewegst dich zu wenig oder ohne roten Faden. Dein Körper bekommt keinen klaren Reiz – deshalb tut sich trotz Mühe zu wenig.",
    hebel: "2–3 feste Ganzkörper-Einheiten pro Woche mit Fortschritt.",
    task: "Leg diese Woche 2 feste Trainingstermine in den Kalender und notiere bei jeder Übung Gewicht und Wiederholungen."},
  kopf: {name: "Der Alles-oder-Nichts-Typ",
    text: "Du kannst es eigentlich – aber ein schlechter Abend kippt gleich die ganze Woche. Das Problem ist nicht die Disziplin, sondern der Anspruch, perfekt zu sein.",
    hebel: "Die 80-%-Regel: Ein Ausrutscher ist eine Mahlzeit, kein Tag. Plus ein fester Abend-Snack im Plan.",
    task: "Plane diese Woche bewusst einen Abend-Snack mit Protein ein und schreib jeden Abend eine Sache auf, die gut lief."}
};

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
    if (sh.getMaxColumns() < COLS.length) sh.insertColumnsAfter(sh.getMaxColumns(), COLS.length - sh.getMaxColumns());
    if (sh.getRange(1, COLS.length).getValue() !== COLS[COLS.length - 1]) sh.getRange(1, 1, 1, COLS.length).setValues([COLS]);
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat();
    let row = ids.indexOf(d.sid) + 1;
    let prev = {};
    if (row > 0) {
      const old = sh.getRange(row, 1, 1, COLS.length).getValues()[0];
      COLS.forEach((c, k) => prev[c] = old[k]);
    }
    // Versand-Vermerke aus der bestehenden Zeile behalten
    d.mail_ergebnis = prev.mail_ergebnis || "";
    d.mail_plan = prev.mail_plan || "";

    // E-Mails an den Lead (je Schritt höchstens einmal)
    if (d.email && validMail(d.email)) {
      if (d.status === "lead" && !d.mail_ergebnis && mailBudget(d.email)) {
        sendResultMail(d); d.mail_ergebnis = new Date().toISOString();
      }
      if (d.status === "plan" && !d.mail_plan && mailBudget(d.email)) {
        if (sendPlanMail(d)) d.mail_plan = new Date().toISOString();
      }
    }

    const values = [COLS.map(c => d[c] !== undefined ? d[c] : "")];
    if (row > 0) sh.getRange(row, 1, 1, COLS.length).setValues(values);
    else { sh.appendRow(values[0]); row = sh.getLastRow(); }

    // Benachrichtigung an Philip
    const subjectBase = (d.name || "Unbekannt") + " – " + (d.typ || "");
    if (d.status === "calendly_klick" && prev.status !== "calendly_klick") {
      MailApp.sendEmail(MAIL_TO, "📅 Calendly-Klick: " + subjectBase, summary(d));
    } else if (d.status === "lead" && d.heiss === "JA" && prev.status !== "lead") {
      MailApp.sendEmail(MAIL_TO, "🔥 Heißer Lead: " + subjectBase, summary(d));
    }
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Mails an den Lead ---------- */
function validMail(m) { return /^[^\s@<>"]+@[^\s@<>"]+\.[a-z]{2,}$/i.test(String(m)) && String(m).length <= 80; }
// Schutz gegen Missbrauch: höchstens 4 Mails pro Adresse in 6 Stunden
function mailBudget(m) {
  const c = CacheService.getScriptCache(), k = "mb_" + Utilities.base64EncodeWebSafe(String(m).toLowerCase()).slice(0, 200);
  const n = Number(c.get(k) || 0);
  if (n >= 4 || MailApp.getRemainingDailyQuota() < 5) return false;
  c.put(k, String(n + 1), 21600);
  return true;
}
function firstName(d) { return String(d.name || "").replace(/[^\p{L}\- ]/gu, "").trim().split(" ")[0].slice(0, 20); }
function h(s) { return String(s).replace(/[&<>"]/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])); }

function shell(inner) {
  return '<div style="background:#EDF5EE;padding:24px 12px;font-family:Arial,Helvetica,sans-serif;color:#13241A">' +
    '<div style="max-width:560px;margin:0 auto;background:#fff;border-radius:16px;overflow:hidden">' +
    '<div style="background:#1B5A2F;background:linear-gradient(160deg,#103A20,#2E8B3D);padding:20px 24px;color:#fff">' +
    '<img src="' + CHECK_URL + 'logo-weiss.png" width="28" alt="" style="vertical-align:middle;margin-right:10px">' +
    '<span style="font-weight:bold;letter-spacing:.06em;vertical-align:middle">SCHUMANN CONSULTING</span></div>' +
    '<div style="padding:24px;font-size:16px;line-height:1.55">' + inner +
    '<table style="margin-top:24px;border-top:1px solid #D3E0D6;padding-top:16px;width:100%"><tr>' +
    '<td width="56"><img src="' + CHECK_URL + 'philip.jpg" width="48" height="48" alt="" style="border-radius:50%"></td>' +
    '<td style="font-size:14px;color:#526457"><b style="color:#13241A">Philip Schumann</b><br>Online-Fitnesscoach · Instagram @philip.schumann</td></tr></table>' +
    '</div></div>' +
    '<p style="max-width:560px;margin:12px auto 0;font-size:12px;color:#526457;text-align:center">Du bekommst diese Mail, weil du den Fitness-Engpass-Check gemacht hast. ' +
    'Keine weiteren Mails gewünscht? Antworte einfach mit „Stopp“. · <a href="' + CHECK_URL + 'datenschutz.html" style="color:#526457">Datenschutz & Impressum</a></p></div>';
}
function button(href, label) {
  return '<p style="margin:22px 0"><a href="' + href + '" style="background:#267A34;color:#fff;text-decoration:none;font-weight:bold;padding:14px 22px;border-radius:12px;display:inline-block">' + label + '</a></p>';
}

function sendResultMail(d) {
  const t = TYPES[d.typkey]; if (!t) return;
  const n = firstName(d);
  const inner =
    '<p>Hey' + (n ? ' ' + h(n) : '') + ' 👋</p>' +
    '<p>danke, dass du den Fitness-Engpass-Check gemacht hast! Hier ist dein Ergebnis zum Nachlesen:</p>' +
    '<p style="font-size:13px;letter-spacing:.1em;color:#1F6B2D;font-weight:bold;margin-bottom:4px">DEIN ENGPASS-TYP</p>' +
    '<h2 style="margin:0 0 8px;font-size:24px">' + h(t.name) + '</h2><p>' + h(t.text) + '</p>' +
    '<p><b>Dein größter Hebel:</b> ' + h(t.hebel) + '</p>' +
    '<p><b>Deine Aufgabe für die nächsten 7 Tage:</b> ' + h(t.task) + '</p>' +
    '<p>Wenn du im Check noch deinen Ernährungsplan erstellst, schicke ich dir den mit Rezepten und Einkaufsliste auch per Mail.</p>' +
    '<p>Fragen zu deinem Ergebnis? Antworte einfach auf diese Mail oder schreib mir auf Instagram.</p>';
  MailApp.sendEmail({to: d.email, subject: "Dein Ergebnis: " + t.name, htmlBody: shell(inner),
    name: "Philip Schumann", replyTo: MAIL_TO});
}

function planLink(d) {
  const g = d.geschlecht === "w" || d.geschlecht === "m" ? d.geschlecht : "";
  const a = Number(d.alter), hh = Number(d.groesse), w = Number(d.gewicht), x = Number(d.aktivitaet);
  if (!g || !TYPES[d.typkey] || !(a >= 18 && a <= 75 && hh >= 140 && hh <= 215 && w >= 40 && w <= 200) || [1.3, 1.45, 1.6].indexOf(x) < 0) return "";
  const z = ["abnehmen", "muskeln", "straffer", "zunehmen", "fit"].indexOf(d.ziel) >= 0 ? d.ziel : "abnehmen";
  const dd = ["veg", "nofish", "lf"].indexOf(d.ernaehrungsform) >= 0 ? d.ernaehrungsform : "";
  const r = ["0", "unregel", "1-2", "3+"].indexOf(String(d.training)) >= 0 ? String(d.training) : "";
  const parts = ["g" + g, "a" + a, "h" + hh, "w" + w, "x" + x, "d" + dd, "z" + z, "t" + d.typkey, "r" + encodeURIComponent(r)];
  const n = firstName(d);
  return CHECK_URL + "?p=" + parts.join("_") + (n ? "&name=" + encodeURIComponent(n) : "") + "&v=MAIL";
}

function sendPlanMail(d) {
  const link = planLink(d); if (!link) return false;
  const t = TYPES[d.typkey];
  const n = firstName(d);
  const kcal = Number(d.kcal), prot = Number(d.protein);
  const inner =
    '<p>Hey' + (n ? ' ' + h(n) : '') + ' 👋</p>' +
    '<p>hier ist dein persönlicher Startplan aus dem Fitness-Engpass-Check:</p>' +
    '<table style="width:100%;border-collapse:separate;border-spacing:8px 0;margin:8px 0 4px"><tr>' +
    '<td style="background:#EDF5EE;border-radius:12px;padding:14px;text-align:center"><b style="font-size:26px;color:#1F6B2D">' + kcal + '</b><br><span style="font-size:13px;color:#526457">kcal pro Tag</span></td>' +
    '<td style="background:#EDF5EE;border-radius:12px;padding:14px;text-align:center"><b style="font-size:26px;color:#1F6B2D">' + prot + ' g</b><br><span style="font-size:13px;color:#526457">Protein pro Tag</span></td>' +
    '</tr></table>' +
    '<p>Über den Button kommst du jederzeit zu deinem kompletten Plan: <b>Rezepte mit Mengen für dich, Einkaufsliste für 5 Tage und deine Wochenstruktur.</b> Speicher dir die Mail am besten ab.</p>' +
    button(link, "Meinen Ernährungsplan öffnen") +
    (t ? '<p><b>Dein größter Hebel</b> (' + h(t.name.replace(/^Der /, "")) + '): ' + h(t.hebel) + '</p>' : '') +
    '<p>Wenn du den Plan nicht nur hast, sondern im Alltag auch durchziehen willst, lass uns kurz sprechen – kostenlos und unverbindlich:</p>' +
    '<p><a href="' + CALENDLY + '" style="color:#1F6B2D;font-weight:bold">→ Kostenloses Erstgespräch buchen</a></p>';
  MailApp.sendEmail({to: d.email, subject: "Dein Ernährungsplan ist da" + (n ? ", " + n : ""), htmlBody: shell(inner),
    name: "Philip Schumann", replyTo: MAIL_TO});
  return true;
}

/* ---------- Zusammenfassung für Philip ---------- */
function summary(d) {
  return [
    "Name: " + (d.name || "-"), "E-Mail: " + (d.email || "-"), "Telefon: " + (d.telefon || "-"), "Typ: " + d.typ, "Ziel: " + d.ziel,
    "Wichtigkeit (1–10): " + d.wichtig, "Angehen: " + d.hilfe, "Warum jetzt: " + d.warum,
    "Seit wann 'so nicht weiter': " + (d.seit || "-"),
    "Was es bedeuten würde (seine Worte): " + (d.bedeutung || "-"),
    "Dieses Mal anders (seine Worte): " + (d.anders || "-"),
    "Kipppunkt: " + d.kipppunkt, "Planung: " + d.planung,
    "Plan: " + (d.kcal ? d.kcal + " kcal / " + d.protein + " g Protein" : "-"),
    "Quelle/Variante: " + (d.quelle || "-"),
  ].join("\n");
}
