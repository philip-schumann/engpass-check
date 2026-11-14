/**
 * Freebie-Backend Schumann Coaching (Google Apps Script) – mehrere Freebies, eine Tabelle.
 * Freebies: kuehlschrank, lebensmittel, weihnachten. Spalte "freebie" in der Tabelle.
 * Ablauf: Anmeldung → Bestätigungsmail (Double-Opt-in) + Sofort-Mail an Philip
 *         → bestaetigt.html meldet Bestätigung → Guide-Mail → 5 Nachfass-Mails (Tag 1, 2, 4, 6, 8).
 */
const MAIL_TO = "p.schumann.business@gmail.com";
const SENDER = "Philip Schumann";
const ROOT = "https://philip-schumann.github.io/engpass-check/";
const COLS = ["token","angemeldet","status","name","email","telefon","telefon_ohne_einwilligung","quelle",
  "einwilligung_mail","einwilligung_tel","bestaetigt","nurture_schritt","letzte_mail","abgemeldet","freebie"];

const FB = {
  kuehlschrank: {titel: "die Kühlschrank-Methode", name: "Kühlschrank-Methode", path: "kuehlschrank/", pdf: "Kuehlschrank-Methode.pdf",
    guideBetreff: "Hier ist deine Kühlschrank-Methode",
    guideTipp: "Mein Tipp: Lies Seite 2 und 3 und nimm dir dann direkt 20 Minuten. Je schneller du startest, desto eher merkst du, wie viel leichter der Abend wird.",
    waFrage: "Dein Kühlschrank-Guide ist unterwegs. Kurze Frage: Was ist bei dir gerade die größte Baustelle – eher der Abend, das Wochenende oder etwas ganz anderes?"},
  lebensmittel: {titel: "den Lebensmittel-Kompass", name: "Lebensmittel-Kompass", path: "lebensmittel/", pdf: "Lebensmittel-Kompass.pdf",
    guideBetreff: "Hier ist dein Lebensmittel-Kompass",
    guideTipp: "Mein Tipp: Lies Seite 2 und 3 und mach dir dann ein Foto von der Einkaufsliste auf Seite 4. Beim nächsten Einkauf hast du sie direkt dabei.",
    waFrage: "Dein Lebensmittel-Kompass ist unterwegs. Kurze Frage: Was ist bei dir gerade das größte Fragezeichen beim Essen – was du einkaufen sollst, wie viel oder wann?"},
  weihnachten: {titel: "die Weihnachts-Formel", name: "Weihnachts-Formel", path: "weihnachten/", pdf: "Weihnachts-Formel.pdf",
    guideBetreff: "Hier ist deine Weihnachts-Formel",
    guideTipp: "Mein Tipp: Nimm dir jetzt kurz deinen Kalender und markier die echten Festtage (Seite 3). Das dauert 2 Minuten und nimmt dir die ganze Adventszeit Druck raus.",
    waFrage: "Deine Weihnachts-Formel ist unterwegs. Kurze Frage: Was ist für dich in der Weihnachtszeit die größte Herausforderung – Weihnachtsmarkt, Plätzchen oder die Feiertage selbst?"}
};
function base(fb) { return ROOT + FB[fb || "kuehlschrank"].path; }
function cal(fb) { return "https://calendly.com/philip-schumann/erstgespraech?utm_source=freebie&utm_medium=mail&utm_campaign=" + (fb || "kuehlschrank"); }

/* ---------- Einrichtung ---------- */
function setup() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty("SHEET_ID")) {
    const ss = SpreadsheetApp.create("Freebie-Leads – Schumann Coaching");
    const sh = ss.getSheets()[0]; sh.setName("Leads"); sh.appendRow(COLS); sh.setFrozenRows(1);
    p.setProperty("SHEET_ID", ss.getId());
  }
  ensureHeader();
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "nurture") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("nurture").timeBased().everyHours(1).create();
}
function sheet() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty("SHEET_ID")).getSheetByName("Leads"); }
function ensureHeader() { const sh = sheet(); COLS.forEach((c, i) => { if (sh.getRange(1, i + 1).getValue() !== c) sh.getRange(1, i + 1).setValue(c); }); }
function rows() { const sh = sheet(); const v = sh.getDataRange().getValues(); return {sh, v}; }
function col(n) { return COLS.indexOf(n); }
function fbOf(row) { const f = row[col("freebie")]; return FB[f] ? f : "kuehlschrank"; }

/* ---------- Anmeldung ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.a === "confirm" || d.a === "unsub") return out(handleToken(d.a, String(d.t || "")));
    if (d.a !== "signup") return out("ok");
    const fb = FB[d.freebie] ? d.freebie : "kuehlschrank";
    const email = String(d.email || "").trim().toLowerCase();
    if (!/^[^\s@<>"]+@[^\s@<>"]+\.[a-z]{2,}$/i.test(email) || email.length > 80) return out("bad");
    const name = clean(d.name, 40), tel = clean(d.telefon, 25).replace(/[^\d+ ]/g, "");
    if (sheet().getLastColumn() < COLS.length) ensureHeader();
    const {sh, v} = rows();
    for (let r = 1; r < v.length; r++) if (v[r][col("email")] === email && fbOf(v[r]) === fb) {
      if (!v[r][col("bestaetigt")] && !v[r][col("abgemeldet")]) sendConfirm(name, email, v[r][col("token")], fb);
      return out("dup");
    }
    if (MailApp.getRemainingDailyQuota() < 5) return out("quota");
    const token = Utilities.getUuid();
    const row = {token, angemeldet: new Date(), status: "unbestätigt", name, email, telefon: tel,
      telefon_ohne_einwilligung: clean(d.telefon_ohne_einwilligung, 4), quelle: clean(d.quelle, 120),
      einwilligung_mail: clean(d.einwilligung_mail, 60), einwilligung_tel: tel ? clean(d.einwilligung_tel, 60) : "",
      bestaetigt: "", nurture_schritt: 0, letzte_mail: "", abgemeldet: "", freebie: fb};
    sh.appendRow(COLS.map(c => row[c]));
    sendConfirm(name, email, token, fb);
    alertPhilip(row, fb);
    return out("ok");
  } finally { lock.releaseLock(); }
}

/* ---------- Bestätigen / Abmelden ---------- */
function handleToken(a, t) {
  if (!t) return "bad";
  const {sh, v} = rows();
  for (let r = 1; r < v.length; r++) if (v[r][col("token")] === t) {
    if (a === "unsub") {
      sh.getRange(r + 1, col("status") + 1).setValue("abgemeldet");
      sh.getRange(r + 1, col("abgemeldet") + 1).setValue(new Date());
      return "unsub";
    }
    if (!v[r][col("bestaetigt")] && !v[r][col("abgemeldet")]) {
      sh.getRange(r + 1, col("status") + 1).setValue("bestätigt");
      sh.getRange(r + 1, col("bestaetigt") + 1).setValue(new Date());
      sendGuide(v[r][col("name")], v[r][col("email")], t, fbOf(v[r]));
    }
    return "confirmed";
  }
  return "notfound";
}
function doGet(e) {
  const a = e.parameter.a, t = e.parameter.t || "";
  if (a === "confirm" || a === "unsub") handleToken(a, t);
  const url = base("kuehlschrank") + (a === "unsub" ? "abmelden.html" : "bestaetigt.html");
  return HtmlService.createHtmlOutput("<p style='font-family:Arial;padding:24px'><a target='_top' href='" + url + "'>Weiter →</a></p>");
}

/* ---------- Mails ---------- */
function sendConfirm(name, email, token, fb) {
  const link = base(fb) + "bestaetigt.html?t=" + token + "&n=" + encodeURIComponent(name);
  mail(email, "Bitte kurz bestätigen – dann kommt dein Guide",
    "<p>Hi " + h(name) + ",</p><p>schön, dass du " + FB[fb].titel + " angefordert hast! Klick einmal auf den Button, dann schicke ich dir den Guide sofort.</p>" +
    btn(link, "Ja, Guide schicken") +
    "<p style='font-size:13px;color:#526457'>Damit bestätigst du, dass ich dir den Guide und weitere Tipps per E-Mail schicken darf. Abmelden geht jederzeit mit einem Klick. Wenn du dich nicht angemeldet hast, ignorier diese Mail einfach.</p>", null, fb);
}
function sendGuide(name, email, token, fb) {
  mail(email, FB[fb].guideBetreff,
    "<p>Hi " + h(name) + ",</p><p>hier ist dein Guide. " + FB[fb].guideTipp + "</p>" +
    btn(base(fb) + FB[fb].pdf, "Guide herunterladen") +
    "<p>Wenn du magst, antworte mir kurz auf diese Mail: Was ist bei dir gerade die größte Baustelle beim Essen? Ich lese jede Antwort selbst.</p><p>Philip</p>", token, fb);
}
function alertPhilip(r, fb) {
  const wa = r.telefon ? "https://wa.me/" + r.telefon.replace(/[^\d]/g, "").replace(/^0/, "49") : "";
  MailApp.sendEmail({to: MAIL_TO, name: "Freebie-Funnel",
    subject: (r.telefon ? "🔥 Neuer Lead MIT Nummer: " : "📩 Neuer Lead: ") + r.name + " (" + FB[fb].name + ")",
    htmlBody: "<p><b>" + h(r.name) + "</b> hat " + FB[fb].titel + " angefordert.</p>" +
      "<p>E-Mail: " + h(r.email) + "<br>Telefon: " + (r.telefon ? h(r.telefon) + " (Einwilligung Anruf/WhatsApp: ja)" : "–") +
      "<br>Quelle: " + h(r.quelle) + "</p>" +
      (r.telefon ? "<p><b>Jetzt innerhalb von 15 Minuten melden:</b><br><a href='" + wa + "'>WhatsApp öffnen</a> · <a href='tel:" + h(r.telefon) + "'>Anrufen</a></p>" +
        "<p style='color:#526457'>Vorschlag: „Hi " + h(r.name) + ", hier ist Philip 👋 " + FB[fb].waFrage + "“</p>" : "")});
}

/* ---------- Nachfass-Mails ---------- */
const NURTURE_K = [
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

const NURTURE_L = [
  {tag: 1, betreff: "Welche zwei Tauschpartner nimmst du?", text:
    "<p>Hi {name},</p><p>auf Seite 6 im Kompass stehen 10 Tauschpartner. Mein Tipp: Such dir genau zwei aus und setz sie diese Woche um. Nicht alle zehn.</p>" +
    "<p>Zwei kleine Wechsel, die du durchziehst, bringen mehr als zehn, die du nach drei Tagen wieder vergisst. Und nächste Woche kommen die nächsten zwei dazu.</p><p>Schreib mir gern, welche zwei du nimmst – ich bin gespannt.</p><p>Philip</p>"},
  {tag: 2, betreff: "Warum Protein zuerst kommt", text:
    "<p>Hi {name},</p><p>wenn du dir aus dem Kompass nur eine Sache merkst, dann diese: Norden zuerst. Also bei jeder Mahlzeit erst die Proteinquelle festlegen, dann den Rest drumherum bauen.</p>" +
    "<p>Protein macht lange satt und hilft dir, Muskeln zu halten, während du abnimmst. Und es nimmt dir die Frage „Was esse ich heute?“ ab: Du fragst nur noch „Welches Protein?“ – Eier, Skyr, Hähnchen, Linsen.</p><p>Philip</p>"},
  {tag: 4, betreff: "Wenig Zeit? Dann wiederhol dich.", text:
    "<p>Hi {name},</p><p>viele denken, gesund essen heißt jeden Tag etwas Neues kochen. Das Gegenteil ist der Fall: Die Leute, die es langfristig schaffen, haben feste Standards.</p>" +
    "<p>Zwei Frühstücke, drei Hauptmahlzeiten, fertig. Das spart dir Einkauf, Planung und Nerven. Abwechslung kommt am Wochenende.</p>" +
    "<p>Wenn du dir dabei Unterstützung wünschst und wissen willst, wie deine Standards konkret aussehen könnten, lass uns sprechen.</p>{cal}<p>Philip</p>"},
  {tag: 6, betreff: "So läuft ein Erstgespräch bei mir ab", text:
    "<p>Hi {name},</p><p>vielleicht hast du schon überlegt, ein Erstgespräch zu buchen, und weißt nicht genau, was dich erwartet. Deshalb ganz transparent:</p>" +
    "<p><b>1.</b> Wir schauen uns an, wo du gerade stehst und was du schon probiert hast.<br><b>2.</b> Wir finden heraus, woran es bei dir wirklich hängt.<br><b>3.</b> Du gehst mit einem klaren nächsten Schritt raus.</p>" +
    "<p>Wenn es passt, erzähle ich dir am Ende, wie eine Zusammenarbeit aussehen würde. Wenn nicht, ist das völlig okay – du hast trotzdem deinen Plan.</p>{cal}<p>Philip</p>"},
  {tag: 8, betreff: "Letzte Mail dieser Reihe", text:
    "<p>Hi {name},</p><p>das ist meine letzte Mail zum Lebensmittel-Kompass. Eine Frage habe ich noch an dich:</p>" +
    "<p><b>Was ist beim Essen gerade dein größtes Fragezeichen?</b> Antworte mir einfach mit einem Satz. Ich lese jede Antwort selbst und melde mich mit einem ehrlichen Tipp.</p>" +
    "<p>Und wenn du das Thema dieses Mal richtig angehen willst, such dir hier einen Termin aus:</p>{cal}<p>Philip</p>"}
];
const NURTURE_W = [
  {tag: 1, betreff: "Hast du deine Festtage schon markiert?", text:
    "<p>Hi {name},</p><p>die wichtigste Übung aus der Weihnachts-Formel dauert zwei Minuten: Kalender auf und die echten Festtage markieren. Weihnachtsfeier, Heiligabend, die Feiertage, Silvester.</p>" +
    "<p>Alles andere sind ganz normale Tage. Klingt banal, verändert aber die ganze Adventszeit – weil aus „Ist doch eh Weihnachten“ plötzlich wieder ein ganz normaler Dienstag wird.</p><p>Philip</p>"},
  {tag: 2, betreff: "Der Trick fürs Festessen", text:
    "<p>Hi {name},</p><p>beim Festessen musst du nichts weglassen. Mach nur eine Sache anders: Fang mit Protein und Gemüse an – Pute, Fisch, Rosenkohl, Karotten. Danach den Rest.</p>" +
    "<p>Du wirst merken, dass du automatisch entspannter bist und nicht mit Heißhunger durch das Buffet gehst. Und danach: 15 Minuten raus an die frische Luft.</p><p>Philip</p>"},
  {tag: 4, betreff: "Warum der Januar-Neustart meistens scheitert", text:
    "<p>Hi {name},</p><p>jedes Jahr dasselbe: Im Dezember „lohnt es sich nicht mehr“, im Januar wird alles auf einmal neu gestartet. Und im Februar ist es wieder vorbei.</p>" +
    "<p>Der Grund: Ein kompletter Neustart braucht viel Energie. Wer im Dezember seine Standards hält, muss im Januar gar nicht neu starten – nur weitermachen.</p>" +
    "<p>Wenn du dieses Jahr schon vor den Feiertagen anfangen willst, lass uns sprechen.</p>{cal}<p>Philip</p>"},
  {tag: 6, betreff: "So läuft ein Erstgespräch bei mir ab", text:
    "<p>Hi {name},</p><p>vielleicht hast du schon überlegt, ein Erstgespräch zu buchen, und weißt nicht genau, was dich erwartet. Deshalb ganz transparent:</p>" +
    "<p><b>1.</b> Wir schauen uns an, wo du gerade stehst und was du schon probiert hast.<br><b>2.</b> Wir finden heraus, woran es bei dir wirklich hängt.<br><b>3.</b> Du gehst mit einem klaren nächsten Schritt raus.</p>" +
    "<p>Wenn es passt, erzähle ich dir am Ende, wie eine Zusammenarbeit aussehen würde. Wenn nicht, ist das völlig okay – du hast trotzdem deinen Plan.</p>{cal}<p>Philip</p>"},
  {tag: 8, betreff: "Letzte Mail dieser Reihe", text:
    "<p>Hi {name},</p><p>das ist meine letzte Mail zur Weihnachts-Formel. Eine Frage habe ich noch an dich:</p>" +
    "<p><b>Was ist für dich die größte Herausforderung in der Weihnachtszeit?</b> Antworte mir einfach mit einem Satz. Ich lese jede Antwort selbst und melde mich mit einem ehrlichen Tipp.</p>" +
    "<p>Und wenn du im Januar nicht wieder bei null anfangen willst, such dir hier einen Termin aus:</p>{cal}<p>Philip</p>"}
];
const NURTURES = {kuehlschrank: NURTURE_K, lebensmittel: NURTURE_L, weihnachten: NURTURE_W};

function nurture() {
  const {sh, v} = rows(); const now = new Date(); let sent = 0;
  for (let r = 1; r < v.length && sent < 40; r++) {
    const b = v[r][col("bestaetigt")]; if (!b || v[r][col("abgemeldet")]) continue;
    const fb = fbOf(v[r]);
    const step = Number(v[r][col("nurture_schritt")]) || 0; const m = NURTURES[fb][step]; if (!m) continue;
    const days = (now - new Date(b)) / 864e5; if (days < m.tag) continue;
    if (MailApp.getRemainingDailyQuota() < 10) break;
    const body = m.text.replace(/\{name\}/g, h(v[r][col("name")])).replace(/\{cal\}/g, btn(cal(fb), "Kostenloses Erstgespräch buchen"));
    mail(v[r][col("email")], m.betreff, body, v[r][col("token")], fb);
    sh.getRange(r + 1, col("nurture_schritt") + 1).setValue(step + 1);
    sh.getRange(r + 1, col("letzte_mail") + 1).setValue(now);
    sent++;
  }
}

/* ---------- Helfer ---------- */
function mail(to, subject, inner, token, fb) {
  const unsub = token ? "Du bekommst diese Mail, weil du " + FB[fb || "kuehlschrank"].titel + " angefordert hast. <a style='color:#8FA597' href='" +
    base(fb) + "abmelden.html?t=" + token + "'>Abmelden</a><br>" : "";
  const html =
    "<div style='background:#EDF5EE;padding:24px 12px;font-family:Helvetica,Arial,sans-serif'>" +
    "<div style='max-width:560px;margin:0 auto'>" +
    "<div style='background:#123D22;border-radius:16px 16px 0 0;padding:20px 26px'>" +
    "<img src='https://philip-schumann.github.io/engpass-check/logo-weiss.png' width='26' style='vertical-align:middle' alt=''>" +
    "<span style='color:#C9EECF;font-size:12px;font-weight:bold;letter-spacing:2px;text-transform:uppercase;vertical-align:middle;margin-left:10px'>Schumann Consulting</span></div>" +
    "<div style='background:#ffffff;border-radius:0 0 16px 16px;padding:28px 26px;font-size:16px;line-height:1.6;color:#13241A'>" + inner +
    "<table role='presentation' style='margin-top:26px;border-top:1px solid #D3E0D6;padding-top:18px;width:100%'><tr>" +
    "<td width='56'><img src='" + "https://philip-schumann.github.io/engpass-check/philip.jpg' width='48' height='48' style='border-radius:50%;border:2px solid #2E8B3D;object-fit:cover' alt=''></td>" +
    "<td style='font-size:14px;color:#526457'><b style='color:#13241A'>Philip Schumann</b><br>Online-Fitnesscoach · @philip.schumann</td></tr></table></div>" +
    "<p style='font-size:12px;color:#8FA597;text-align:center;margin:16px 8px 0;line-height:1.5'>" + unsub + "Philip Schumann · Schumann Consulting · Judengasse 43, 96450 Coburg</p>" +
    "</div></div>";
  MailApp.sendEmail({to, subject, name: SENDER, replyTo: MAIL_TO, htmlBody: html});
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
