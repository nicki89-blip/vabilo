/**
 * Nino 40: zapis potrditev v Google tabelo.
 *
 * Koda je vezana na tabelo "Nino 40 potrditve" (Razširitve > Apps Script).
 * Prvič zaženi setup(), nato uvedi kot spletno aplikacijo (Izvedi kot: jaz, Dostop: Vsi).
 * Pri popravkih: Uvedi > Upravljaj uvedbe > uredi obstoječo > Nova različica (URL ostane enak).
 */

var SHEET_ODGOVORI = "Odgovori";
var SHEET_POVZETEK = "Povzetek";
var SHEET_POVABLJENI = "Povabljeni";
var HEADER = ["Čas prejema", "Ime", "Udeležba", "Število oseb", "Prenočišče", "Opombe", "Čas v brskalniku", "ID"];

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Besedilo obreže, omeji dolžino in prepreči, da bi ga tabela brala kot formulo.
function clean_(value, max) {
  var s = String(value == null ? "" : value).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").trim();
  if (s.length > max) s = s.slice(0, max);
  if (/^[=+@\-]/.test(s)) s = "'" + s;
  return s;
}

function yesNo_(value) {
  return String(value).toLowerCase() === "da" ? "da" : "ne";
}

function doGet() {
  return json_({ ok: true, service: "nino40" });
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: "bad_json" });
  }
  if (!data || typeof data !== "object") return json_({ ok: false, error: "bad_json" });

  // Past za bote: človek skritega polja ne vidi in ga ne izpolni.
  if (String(data.website || "").trim() !== "") return json_({ ok: true });

  var ime = clean_(data.ime, 80);
  if (!ime) return json_({ ok: false, error: "missing_name" });

  var udelezba = yesNo_(data.udelezba);
  var stevilo = parseInt(data.stevilo, 10);
  if (isNaN(stevilo)) stevilo = udelezba === "da" ? 1 : 0;
  stevilo = Math.max(0, Math.min(10, stevilo));
  if (udelezba === "ne") stevilo = 0;
  if (udelezba === "da" && stevilo === 0) stevilo = 1;
  var prenocisce = udelezba === "da" ? yesNo_(data.prenocisce) : "ne";
  var opombe = clean_(data.opombe, 500);
  var poslano = clean_(data.poslano, 40);
  // Naključna oznaka brskalnika: popravek odgovora iz istega brskalnika zamenja prejšnjega, tudi če se ime spremeni.
  var id = /^[A-Za-z0-9-]{8,40}$/.test(String(data.id || "")) ? String(data.id) : "";

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (err) {
    return json_({ ok: false, error: "busy" });
  }
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ODGOVORI);
    if (!sheet) return json_({ ok: false, error: "no_sheet" });
    sheet.appendRow([new Date(), ime, udelezba, stevilo, prenocisce, opombe, poslano, id]);
  } catch (err) {
    return json_({ ok: false, error: "write_failed" });
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

// Tabele z evropskimi nastavitvami (npr. slovenskimi) v formulah zahtevajo podpičje namesto vejice.
// Preizkusimo na pomožni celici, nato po potrebi zamenjamo vejice zunaj narekovajev.
function separator_(sheet) {
  var cell = sheet.getRange("Z1");
  cell.setFormula("=SUM(1,2)");
  SpreadsheetApp.flush();
  var ok = cell.getValue() === 3;
  cell.clearContent();
  return ok ? "," : ";";
}

function formula_(f, sep) {
  if (sep === ",") return f;
  var out = "", inStr = false;
  for (var i = 0; i < f.length; i++) {
    var c = f.charAt(i);
    if (c === '"') inStr = !inStr;
    out += (c === "," && !inStr) ? sep : c;
  }
  return out;
}

/**
 * Zaženi enkrat. Ustvari zavihke, glave in formule. Ponovni zagon ne briše odgovorov.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone("Europe/Ljubljana");
  // Jezikovnih nastavitev tabele namenoma ne spreminjamo: formule so zapisane z vejicami.

  // Odgovori
  var odg = ss.getSheetByName(SHEET_ODGOVORI) || ss.insertSheet(SHEET_ODGOVORI, 0);
  odg.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight("bold").setBackground("#d4e6e8");
  odg.setFrozenRows(1);
  odg.getRange("A:A").setNumberFormat("d. m. yyyy hh:mm");
  odg.getRange("B:B").setNumberFormat("@");
  odg.getRange("F:H").setNumberFormat("@");
  odg.setColumnWidth(1, 140);
  odg.setColumnWidth(2, 180);
  odg.setColumnWidth(6, 320);

  // Povzetek
  var pov = ss.getSheetByName(SHEET_POVZETEK) || ss.insertSheet(SHEET_POVZETEK, 1);
  pov.clear();
  var sep = separator_(pov);
  pov.getRange("A1").setValue("Povzetek potrditev (velja zadnji odgovor vsakega para)").setFontWeight("bold").setFontSize(13);
  pov.getRange("A3:A7").setValues([
    ["Skupaj oseb, ki pridejo"],
    ["Gospodinjstva, ki pridejo"],
    ["Gospodinjstva, ki so odpovedala"],
    ["Potrditve s prenočiščem"],
    ["Zadnji odgovor"]
  ]);
  pov.getRange("B3").setFormula(formula_('=SUMIFS(D11:D, C11:C, "da")', sep));
  pov.getRange("B4").setFormula(formula_('=COUNTIFS(C11:C, "da")', sep));
  pov.getRange("B5").setFormula(formula_('=COUNTIFS(C11:C, "ne")', sep));
  pov.getRange("B6").setFormula(formula_('=COUNTIFS(C11:C, "da", E11:E, "da")', sep));
  pov.getRange("B7").setFormula(formula_('=IF(COUNT(Odgovori!A2:A) = 0, "", MAX(Odgovori!A2:A))', sep)).setNumberFormat("d. m. yyyy hh:mm");
  pov.getRange("B3:B7").setFontWeight("bold").setHorizontalAlignment("left");

  pov.getRange("A9").setValue("Zadnji odgovori, ena vrstica na ime").setFontWeight("bold");
  pov.getRange(10, 1, 1, 6).setValues([HEADER.slice(0, 6)]).setFontWeight("bold").setBackground("#d4e6e8");
  // Odgovore združi po oznaki brskalnika (stolpec ID); če je ni, po normaliziranem imenu (male črke, brez presledkov na robovih).
  // Za vsako skupino vzame zadnjo vrstico (odgovori se dodajajo po vrsti) in jih razvrsti po imenu.
  pov.getRange("A11").setFormula(formula_(
    '=IFERROR(LET(kk, ARRAYFORMULA(IF(Odgovori!H2:H <> "", Odgovori!H2:H, LOWER(TRIM(Odgovori!B2:B)))), ' +
    'k, UNIQUE(FILTER(kk, Odgovori!B2:B <> "")), ' +
    'r, MAP(k, LAMBDA(x, XMATCH(x, kk, 0, -1))), ' +
    'SORT(FILTER(Odgovori!A2:F, ARRAYFORMULA(ISNUMBER(MATCH(SEQUENCE(ROWS(kk)), r, 0)))), 2, TRUE)), "")', sep));
  pov.getRange("A11:A").setNumberFormat("d. m. yyyy hh:mm");
  pov.setColumnWidth(1, 220);
  pov.setColumnWidth(2, 180);
  pov.setColumnWidth(6, 320);
  pov.setFrozenRows(10);

  // Povabljeni
  var pv = ss.getSheetByName(SHEET_POVABLJENI) || ss.insertSheet(SHEET_POVABLJENI, 2);
  pv.getRange("A1:B1").setValues([["Ime", "Status"]]).setFontWeight("bold").setBackground("#d4e6e8");
  pv.setFrozenRows(1);
  pv.getRange("B2:B").clearContent();
  // Status: najprej po osebni povezavi (ID "gost-..." iz imena, enako kot slug() na strani),
  // sicer po imenu v zadnjem odgovoru in nato po najnovejšem odgovoru iste skupine.
  pv.getRange("B2").setFormula(formula_(
    '=LET(n, ARRAYFORMULA(LOWER(TRIM(Odgovori!B2:B))), ' +
    'kk, ARRAYFORMULA(IF(Odgovori!H2:H <> "", Odgovori!H2:H, n)), ' +
    'MAP(A2:A, LAMBDA(x, IF(TRIM(x) = "", "", LET(' +
    'kljuc, "gost-" & REGEXREPLACE(LEFT(REGEXREPLACE(REGEXREPLACE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(' +
    'LOWER(TRIM(x)), "č", "c"), "š", "s"), "ž", "z"), "ć", "c"), "đ", "d"), "[^a-z0-9]+", "-"), "^-+|-+$", ""), 35), "-+$", ""), ' +
    'poimenu, XLOOKUP(LOWER(TRIM(x)), n, kk, "", 0, -1), ' +
    'g, IF(COUNTIF(Odgovori!H2:H, kljuc) > 0, kljuc, poimenu), ' +
    'u, IF(g = "", "", XLOOKUP(g, kk, Odgovori!C2:C, "", 0, -1)), ' +
    's, IF(g = "", 0, XLOOKUP(g, kk, Odgovori!D2:D, 0, 0, -1)), ' +
    'IF(u = "da", "pride (" & s & ")", IF(u = "ne", "ne pride", "ni odgovora")))))))', sep));
  // Osebna povezava za vsak par: pošlji jo obema partnerjema.
  pv.getRange("C1").setValue("Osebna povezava").setFontWeight("bold").setBackground("#d4e6e8");
  pv.getRange("C2:C").clearContent();
  pv.getRange("C2").setFormula(formula_(
    '=MAP(A2:A, LAMBDA(x, IF(TRIM(x) = "", "", "https://nicki89-blip.github.io/vabilo/?za=" & ENCODEURL(TRIM(x)))))', sep));
  pv.setColumnWidth(3, 420);
  // Besedilo vabila za WhatsApp z osebno povezavo. V stolpec E (Jezik) vpiši "en" za angleško različico.
  pv.getRange("D1:E1").setValues([["Vabilo", "Jezik"]]).setFontWeight("bold").setBackground("#d4e6e8");
  pv.getRange("D2:D").clearContent();
  pv.getRange("D2").setFormula(formula_(
    '=MAP(A2:A, E2:E, LAMBDA(x, j, IF(TRIM(x) = "", "", LET(' +
    'p, "https://nicki89-blip.github.io/vabilo/?za=" & ENCODEURL(TRIM(x)), ' +
    'IF(LOWER(TRIM(j)) = "en", ' +
    '"Hi!" & CHAR(10) & CHAR(10) & "I\'m celebrating my 40th and I\'d love for you to join the celebration. All the details are in the invitation below." & CHAR(10) & p & CHAR(10) & CHAR(10) & "See you there!" & CHAR(10) & "Nino", ' +
    '"Živjo! 14.11 ob 19.00 praznujem svojih 40 v Štumfabriki na Polzeli in res bi bil vesel, če prideta! 🥂" & CHAR(10) & ' +
    '"Tokrat brez otrok – čaka nas večer za odrasle. 😉" & CHAR(10) & ' +
    '"Vabilo in potrditev prosim do 31. 10.:" & CHAR(10) & p & CHAR(10) & ' +
    '"Vsak gost ima svojo povezavo – tale je vajina. 😊")))))', sep));
  pv.setColumnWidth(4, 360);
  pv.setColumnWidth(5, 70);
  pv.setColumnWidth(1, 220);
  pv.setColumnWidth(2, 160);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("ni odgovora")
    .setBackground("#fff2a8")
    .setRanges([pv.getRange("B2:B")])
    .build();
  pv.setConditionalFormatRules([rule]);
  pv.getRange("F1").setValue("Vsakemu paru pošlji njegovo osebno povezavo (stolpec C). Ime v stolpcu A po pošiljanju ne spreminjaj, sicer se povezava ne ujema več.").setFontStyle("italic").setFontColor("#5c7276");

  var def = ss.getSheetByName("Sheet1") || ss.getSheetByName("List1");
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 3) ss.deleteSheet(def);
}

/**
 * Za preizkus formul: doda testne odgovore (ime se začne s "TEST ").
 * Pričakovano v Povzetku: skupaj oseb 5, gospodinjstva pridejo 2, odpovedala 1, s prenočiščem 1.
 */
function dodajTestnePodatke() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ODGOVORI);
  var t = new Date().getTime();
  var rows = [
    [new Date(t - 50000), "TEST Ana Novak", "da", 2, "ne", "prvi odgovor", ""],
    [new Date(t - 40000), "TEST Boris", "da", 1, "ne", "", ""],
    [new Date(t - 30000), "test ana novak ", "da", 3, "da", "premislila se je", ""],
    [new Date(t - 20000), "TEST Cene", "da", 2, "ne", "", ""],
    [new Date(t - 10000), "TEST Boris", "ne", 0, "ne", "vseeno ne more", ""]
  ];
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/** Pobriše vse vrstice, katerih ime se začne s "test " (ne glede na velike črke). */
function pobrisiTestnePodatke() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ODGOVORI);
  var last = sheet.getLastRow();
  if (last < 2) return;
  var names = sheet.getRange(2, 2, last - 1, 1).getValues();
  for (var i = names.length - 1; i >= 0; i--) {
    if (/^test /i.test(String(names[i][0]).trim())) sheet.deleteRow(i + 2);
  }
}
