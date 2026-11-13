/**
 * Kühlschrank-Methode – Backend (Google Apps Script, eigenes Projekt)
 * 1. setup() einmal ausführen: legt die Tabelle "Freebie-Leads" an und den stündlichen Trigger für die Nachfass-Mails.
 * 2. Bereitstellen → Web-App (Ausführen als: ich, Zugriff: jeder) → URL in kuehlschrank/index.html bei ENDPOINT eintragen.
 *
 * Ablauf: Anmeldung → Bestätigungsmail (Double-Opt-in) + Sofort-Mail an Philip
 *         → Klick auf Bestätigen → Guide-Seite + Guide-Mail → Nachfass-Mails an Tag 1, 2, 4, 6, 8.
 */
const MAIL_TO = "p.schumann.business@gmail.com";
const SENDER = "Philip Schumann";
const BASE = "https://philip-schumann.github.io/engpass-check/kuehlschrank/";
const PDF = BASE + "Kuehlschrank-Methode.pdf";
const CAL = "https://calendly.com/philip-schumann/erstgespraech?utm_source=freebie&utm_medium=mail&utm_campaign=kuehlschrank";
const COLS = ["token","angemeldet","status","name","email","telefon","telefon_ohne_einwilligung","quelle",
  "einwilligung_mail","einwilligung_tel","bestaetigt","nurture_schritt","letzte_mail","abgemeldet"];

/* ---------- Einrichtung ---------- */
function setup() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty("SHEET_ID")) {
    const ss = SpreadsheetApp.create("Freebie-Leads – Kühlschrank-Methode");
    const sh = ss.getSheets()[0]; sh.setName("Leads"); sh.appendRow(COLS); sh.setFrozenRows(1);
    p.setProperty("SHEET_ID", ss.getId());
  }
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "nurture") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("nurture").timeBased().everyHours(1).create();
  Logger.log("Tabelle: https://docs.google.com/spreadsheets/d/" + p.getProperty("SHEET_ID"));
}
function sheet() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty("SHEET_ID")).getSheetByName("Leads"); }
function rows() { const sh = sheet(); const v = sh.getDataRange().getValues(); return {sh, v}; }
function col(n) { return COLS.indexOf(n); }

/* ---------- Anmeldung ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.a !== "signup") return out("ok");
    const email = String(d.email || "").trim().toLowerCase();
    if (!/^[^\s@<>"]+@[^\s@<>"]+\.[a-z]{2,}$/i.test(email) || email.length > 80) return out("bad");
    const name = clean(d.name, 40), tel = clean(d.telefon, 25).replace(/[^\d+ ]/g, "");
    const {sh, v} = rows();
    // Doppelte Anmeldung: bestehende Zeile nutzen, Bestätigungsmail höchstens erneut senden
    for (let r = 1; r < v.length; r++) if (v[r][col("email")] === email) {
      if (!v[r][col("bestaetigt")] && !v[r][col("abgemeldet")]) sendConfirm(name, email, v[r][col("token")]);
      return out("dup");
    }
    if (MailApp.getRemainingDailyQuota() < 5) return out("quota");
    const token = Utilities.getUuid();
    const row = {token, angemeldet: new Date(), status: "unbestätigt", name, email, telefon: tel,
      telefon_ohne_einwilligung: clean(d.telefon_ohne_einwilligung, 4), quelle: clean(d.quelle, 120),
      einwilligung_mail: clean(d.einwilligung_mail, 60), einwilligung_tel: tel ? clean(d.einwilligung_tel, 60) : "",
      bestaetigt: "", nurture_schritt: 0, letzte_mail: "", abgemeldet: ""};
    sh.appendRow(COLS.map(c => row[c]));
    sendConfirm(name, email, token);
    alertPhilip(row);
    return out("ok");
  } finally { lock.releaseLock(); }
}

/* ---------- Bestätigen / Abmelden ---------- */
function doGet(e) {
  const a = e.parameter.a, t = e.parameter.t;
  if (!t) return page("Hoppla", "<p>Dieser Link ist unvollständig.</p>");
  const {sh, v} = rows();
  for (let r = 1; r < v.length; r++) if (v[r][col("token")] === t) {
    const name = v[r][col("name")], email = v[r][col("email")];
    if (a === "unsub") {
      sh.getRange(r + 1, col("status") + 1).setValue("abgemeldet");
      sh.getRange(r + 1, col("abgemeldet") + 1).setValue(new Date());
      return page("Du bist abgemeldet", "<p>Du bekommst keine weiteren E-Mails mehr von mir. Alles Gute, " + h(name) + "!</p>");
    }
    if (a === "confirm") {
      if (!v[r][col("bestaetigt")]) {
        sh.getRange(r + 1, col("status") + 1).setValue("bestätigt");
        sh.getRange(r + 1, col("bestaetigt") + 1).setValue(new Date());
        sendGuide(name, email, t);
      }
      return page("Bestätigt – hier ist dein Guide, " + h(name) + "!",
        "<p>Danke! Den Guide habe ich dir zusätzlich per Mail geschickt.</p>" +
        btn(PDF, "Kühlschrank-Methode herunterladen") +
        "<h2 style='margin-top:28px'>Und wenn du willst, dass es diesmal wirklich klappt</h2>" +
        "<p>Im kostenlosen Erstgespräch schauen wir uns an, wo es bei dir hakt, und du gehst mit einem klaren nächsten Schritt raus.</p>" +
        btn(CAL, "Kostenloses Erstgespräch buchen", true));
    }
  }
  return page("Link nicht gefunden", "<p>Dieser Link ist abgelaufen oder ungültig. Schreib mir einfach an " + MAIL_TO + ".</p>");
}

/* ---------- Mails ---------- */
function sendConfirm(name, email, token) {
  const link = ScriptApp.getService().getUrl() + "?a=confirm&t=" + token;
  mail(email, "Bitte kurz bestätigen – dann kommt dein Guide",
    "<p>Hi " + h(name) + ",</p><p>danke für dein Interesse an der Kühlschrank-Methode! Klick einmal auf den Button, dann schicke ich dir den Guide sofort.</p>" +
    btn(link, "Ja, Guide schicken") +
    "<p style='font-size:13px;color:#526457'>Damit bestätigst du, dass ich dir den Guide und weitere Tipps per E-Mail schicken darf. Abmelden geht jederzeit mit einem Klick. Wenn du dich nicht angemeldet hast, ignorier diese Mail einfach.</p>", null);
}
function sendGuide(name, email, token) {
  mail(email, "Hier ist deine Kühlschrank-Methode",
    "<p>Hi " + h(name) + ",</p><p>hier ist dein Guide. Mein Tipp: Lies Seite 2 und 3 und nimm dir dann direkt 20 Minuten. Je schneller du startest, desto eher merkst du, wie viel leichter der Abend wird.</p>" +
    btn(PDF, "Guide herunterladen") +
    "<p>Wenn du magst, antworte mir kurz auf diese Mail: Was ist bei dir gerade die größte Baustelle beim Essen? Ich lese jede Antwort selbst.</p><p>Philip</p>", token);
}
function alertPhilip(r) {
  const wa = r.telefon ? "https://wa.me/" + r.telefon.replace(/[^\d]/g, "").replace(/^0/, "49") : "";
  MailApp.sendEmail({to: MAIL_TO, name: "Freebie-Funnel",
    subject: (r.telefon ? "🔥 Neuer Lead MIT Nummer: " : "📩 Neuer Lead: ") + r.name,
    htmlBody: "<p><b>" + h(r.name) + "</b> hat die Kühlschrank-Methode angefordert.</p>" +
      "<p>E-Mail: " + h(r.email) + "<br>Telefon: " + (r.telefon ? h(r.telefon) + " (Einwilligung Anruf/WhatsApp: ja)" : "–") +
      "<br>Quelle: " + h(r.quelle) + "</p>" +
      (r.telefon ? "<p><b>Jetzt innerhalb von 15 Minuten melden:</b><br><a href='" + wa + "'>WhatsApp öffnen</a> · <a href='tel:" + h(r.telefon) + "'>Anrufen</a></p>" +
        "<p style='color:#526457'>Vorschlag: „Hi " + h(r.name) + ", hier ist Philip 👋 Dein Kühlschrank-Guide ist unterwegs. Kurze Frage: Was ist bei dir gerade die größte Baustelle – eher der Abend, das Wochenende oder etwas ganz anderes?“</p>" : "")});
}

/* ---------- Nachfass-Mails (Tag nach Bestätigung) ---------- */
const NURTURE = [
  {tag: 1, betreff: "Hast du schon umgeräumt?", text:
    "<p>Hi {name},</p><p>kurze Frage: Hast du die 20 Minuten schon investiert? Falls nicht, kein Stress – aber mach es am besten heute, bevor der Guide in deinem Download-Ordner verschwindet.</p>" +
    "<p>Der häufigste Fehler, den ich sehe: Am Sonntag ist die Griffzone voll, am Mittwoch ist sie leer. Und dann greifst du wieder zu dem, was gerade da ist.</p>" +
    "<p>Die Lösung ist simpel: Leg dir einen festen 10-Minuten-Termin pro Woche ins Handy, an dem du die Griffzone auffüllst. Bei mir ist das Sonntagabend. Ein Termin schlägt jeden guten Vorsatz.</p><p>Philip</p>"},
  {tag: 2, betreff: "Warum es abends immer schiefgeht", text:
    "<p>Hi {name},</p><p>tagsüber klappt es oft gut. Und dann kommt der Abend.</p>" +
    "<p>Das liegt nicht an dir. Abends bist du müde, hast den ganzen Tag Entscheidungen getroffen und willst einfach nur runterkommen. Genau dann auf Willenskraft zu setzen, ist ein Plan, der fast immer scheitert.</p>" +
    "<p>Was stattdessen funktioniert: Den Abend vorher entscheiden. Plan dir einen festen Abend-Snack ein, zum Beispiel den Schoko-Quark aus dem Guide. Dann ist die Frage „Was esse ich jetzt?“ schon beantwortet, bevor sie auftaucht.</p><p>Philip</p>"},
  {tag: 4, betreff: "Das Wochenende ist kein Problem (wenn du das weißt)", text:
    "<p>Hi {name},</p><p>viele Frauen, mit denen ich spreche, schaffen unter der Woche richtig viel – und am Wochenende ist alles wieder weg. Am Montag kommt dann der Gedanke: „Jetzt ist es eh egal.“</p>" +
    "<p>Hier hilft eine Regel: Ein Ausrutscher ist eine Mahlzeit, kein Tag und erst recht keine Woche. Die Pizza am Samstag ist eine Mahlzeit. Danach geht es ganz normal weiter, ohne Ausgleich, ohne schlechtes Gewissen.</p>" +
    "<p>Wenn du merkst, dass es bei dir genau daran hängt, lass uns sprechen. Im kostenlosen Erstgespräch schauen wir uns an, wo es bei dir hakt.</p>{cal}<p>Philip</p>"},
  {tag: 6, betreff: "So läuft ein Erstgespräch bei mir ab", text:
    "<p>Hi {name},</p><p>vielleicht hast du schon überlegt, ein Erstgespräch zu buchen, und weißt nicht genau, was dich erwartet. Deshalb ganz transparent:</p>" +
    "<p><b>1.</b> Wir schauen uns an, wo du gerade stehst und was du schon probiert hast.<br><b>2.</b> Wir finden heraus, woran es bei dir wirklich hängt.<br><b>3.</b> Du gehst mit einem klaren nächsten Schritt raus.</p>" +
    "<p>Wenn es passt, erzähle ich dir am Ende, wie eine Zusammenarbeit aussehen würde. Wenn nicht, ist das völlig okay – du hast trotzdem deinen Plan.</p>{cal}<p>Philip</p>"},
  {tag: 8, betreff: "Letzte Mail dieser Reihe", text:
    "<p>Hi {name},</p><p>das ist meine letzte Mail zur Kühlschrank-Methode. Eine Frage habe ich noch an dich:</p>" +
    "<p><b>Woran hängt es bei dir gerade am meisten?</b> Antworte mir einfach mit einem Satz. Ich lese jede Antwort selbst und melde mich mit einem ehrlichen Tipp.</p>" +
    "<p>Und wenn du das Thema dieses Mal richtig angehen willst, such dir hier einen Termin aus:</p>{cal}<p>Philip</p>"}
];
function nurture() {
  const {sh, v} = rows(); const now = new Date(); let sent = 0;
  for (let r = 1; r < v.length && sent < 40; r++) {
    const b = v[r][col("bestaetigt")]; if (!b || v[r][col("abgemeldet")]) continue;
    const step = Number(v[r][col("nurture_schritt")]) || 0; const m = NURTURE[step]; if (!m) continue;
    const days = (now - new Date(b)) / 864e5; if (days < m.tag) continue;
    if (MailApp.getRemainingDailyQuota() < 10) break;
    const body = m.text.replace(/\{name\}/g, h(v[r][col("name")])).replace(/\{cal\}/g, btn(CAL, "Kostenloses Erstgespräch buchen"));
    mail(v[r][col("email")], m.betreff, body, v[r][col("token")]);
    sh.getRange(r + 1, col("nurture_schritt") + 1).setValue(step + 1);
    sh.getRange(r + 1, col("letzte_mail") + 1).setValue(now);
    sent++;
  }
}

/* ---------- Helfer ---------- */
function mail(to, subject, inner, token) {
  const unsub = token ? "<p style='font-size:12px;color:#8a9a90;margin-top:28px'>Du bekommst diese Mail, weil du die Kühlschrank-Methode angefordert hast. <a style='color:#8a9a90' href='" +
    ScriptApp.getService().getUrl() + "?a=unsub&t=" + token + "'>Abmelden</a> · Philip Schumann, Judengasse 43, 96450 Coburg</p>" : "";
  MailApp.sendEmail({to, subject, name: SENDER, replyTo: MAIL_TO,
    htmlBody: "<div style='font-family:Arial,sans-serif;font-size:16px;line-height:1.55;color:#13241A;max-width:560px'>" + inner + unsub + "</div>"});
}
function btn(href, label, ghost) {
  return "<p style='margin:22px 0'><a target='_blank' href='" + href + "' style='display:inline-block;background:" + (ghost ? "#fff" : "#267A34") + ";color:" + (ghost ? "#1F6B2D" : "#fff") +
    ";border:2px solid #267A34;font-weight:bold;text-decoration:none;padding:14px 22px;border-radius:10px'>" + label + " →</a></p>";
}
function page(title, inner) {
  return HtmlService.createHtmlOutput("<div style='font-family:Arial,sans-serif;max-width:560px;margin:40px auto;padding:0 16px;line-height:1.55;color:#13241A'><h1 style='color:#1F6B2D'>" +
    title + "</h1>" + inner + "</div>").setTitle(title.replace(/<[^>]*>/g, "")).addMetaTag("viewport", "width=device-width, initial-scale=1");
}
function out(s) { return ContentService.createTextOutput(s); }
function clean(s, n) { return String(s || "").replace(/[<>"]/g, "").trim().slice(0, n); }
function h(s) { return String(s).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }
