/**
 * Uteč z AI laboratoře – příjem výsledků do Google Tabulky
 * ---------------------------------------------------------
 * Vložit do: Tabulka „VÝSLEDKY Uteč z AI laboratoře“ → Rozšíření → Apps Script
 * Nasadit:   Nasadit → Nové nasazení → Webová aplikace
 *            Spustit jako: Já · Kdo má přístup: Kdokoli
 * Hra posílá ID token Google; skript ho ověří u Googlu (aud = CLIENT_ID, případně doména HD),
 * takže zapisovat může jen hráč skutečně přihlášený přes Google (libovolný účet).
 */
const CLIENT_ID = '212607744830-77cn0l9ct636g6858d608vst8l1n0gi3.apps.googleusercontent.com'; // stejný jako v index.html
const SHEET_ID = '1IbybqFHcVX_0YDP9Gq34xVDh_WBCo4A1X20SSF80nA0'; // Tabulka VÝSLEDKY na disku Aktivní školy
const HD = '';   // prázdné = libovolný účet Google; 'jak6.cz' = jen školní účty

const LOG = 'Log';
const SUM = 'Přehled';
const SUM_HEAD = ['Začátek', 'Jméno', 'E-mail', 'Stav', 'Zámek', 'Celkový čas', 'Penalizace', 'Nápovědy', '⭐',
  'Z1 čas', 'Z1 chyby', 'Z1 nápov.', 'Z2 čas', 'Z2 chyby', 'Z2 nápov.', 'Z3 čas', 'Z3 chyby', 'Z3 nápov.',
  'Z4 čas', 'Z4 nápov.', 'Z5 čas', 'Z5 pokusy', 'Z5 nápov.', 'Poslední změna', 'Běh'];
const LOG_HEAD = ['Čas', 'Jméno', 'E-mail', 'Událost', 'Zámek', 'Detail', 'Herní čas', 'Běh'];

function doGet() {
  return out_({ ok: true, app: 'Uteč z AI laboratoře' });
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    const user = verify_(p.token);
    if (!user) return out_({ ok: false, error: 'auth' });
    if (!p.runId || !p.summary) return out_({ ok: false, error: 'data' });
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      writeLog_(ss, user, p);
      upsertSummary_(ss, user, p);
      ensureStats_(ss);
    } finally {
      lock.releaseLock();
    }
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

/** Ověří ID token u Googlu; výsledek drží 30 min v cache. */
function verify_(token) {
  if (!token) return null;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token);
  const key = 't' + Utilities.base64EncodeWebSafe(digest).slice(0, 40);
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token),
    { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return null; // neplatný nebo prošlý token
  const t = JSON.parse(r.getContentText());
  if (t.aud !== CLIENT_ID) return null;
  if (String(t.email_verified) !== 'true') return null;
  if (HD && t.hd !== HD) return null;
  const u = { email: t.email, name: t.name || t.email };
  cache.put(key, JSON.stringify(u), 1800);
  return u;
}

function sheet_(ss, name, head) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(head);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#e8f0fe');
    sh.setFrozenRows(1);
  }
  return sh;
}

const dur_ = sec => (Number(sec) || 0) / 86400; // sekundy → hodnota pro formát [m]:ss

function writeLog_(ss, user, p) {
  if (p.type === 'sync') return;
  const sh = sheet_(ss, LOG, LOG_HEAD);
  sh.appendRow([new Date(), user.name, user.email, p.type, p.room || '', String(p.detail || '').slice(0, 60),
    dur_(p.elapsed), p.runId]);
  sh.getRange(sh.getLastRow(), 7).setNumberFormat('[m]:ss');
}

function upsertSummary_(ss, user, p) {
  const sh = sheet_(ss, SUM, SUM_HEAD);
  const s = p.summary, R = s.rooms || {};
  const g = (i, k) => (R[i] && R[i][k]) || 0;
  const row = [
    s.t0 ? new Date(s.t0) : new Date(), user.name, user.email,
    s.status === 'utekl' ? '✅ utekl' : '⏳ hraje', Math.min(s.room || 0, 6),
    dur_(s.total), dur_(s.pen), s.hints || 0, s.star ? '⭐' : '',
    dur_(g(1, 'sec')), g(1, 'wrong'), g(1, 'hints'),
    dur_(g(2, 'sec')), g(2, 'wrong'), g(2, 'hints'),
    dur_(g(3, 'sec')), g(3, 'wrong'), g(3, 'hints'),
    dur_(g(4, 'sec')), g(4, 'hints'),
    dur_(g(5, 'sec')), g(5, 'wrong'), g(5, 'hints'),
    new Date(), p.runId
  ];
  const runCol = SUM_HEAD.length;
  let r = 0;
  if (sh.getLastRow() > 1) {
    const f = sh.getRange(2, runCol, sh.getLastRow() - 1, 1).createTextFinder(p.runId).matchEntireCell(true).findNext();
    if (f) r = f.getRow();
  }
  if (!r) r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, row.length).setValues([row]);
  sh.getRange(r, 1).setNumberFormat('d.M. H:mm');
  [6, 7, 10, 13, 16, 19, 21].forEach(c => sh.getRange(r, c).setNumberFormat('[m]:ss'));
  sh.getRange(r, 24).setNumberFormat('H:mm:ss');
}

/** List Statistika: jak dlouho trvají jednotlivé zámky (vzorce se přepočítávají samy). */
function ensureStats_(ss) {
  if (ss.getSheetByName('Statistika')) return;
  const sh = ss.insertSheet('Statistika');
  const P = "'" + SUM + "'!";
  const locks = [
    ['1 🔐 Šifra', 'J', 'K', 'L'], ['2 💡 Binární čísla', 'M', 'N', 'O'], ['3 🤖 Prompt', 'P', 'Q', 'R'],
    ['4 🛡️ Heslo', 'S', null, 'T'], ['5 🗺️ Bludiště', 'U', 'V', 'W']];
  const rows = [['Zámek', 'Hráčů', 'Průměrný čas', 'Medián času', 'Průměr chyb', 'Průměr nápověd']];
  locks.forEach(([name, t, w, h]) => {
    const T = P + t + '2:' + t;
    rows.push([name,
      '=COUNTIF(' + T + ',">0")',
      '=IFERROR(AVERAGEIF(' + T + ',">0"),"")',
      '=IFERROR(MEDIAN(FILTER(' + T + ',' + T + '>0)),"")',
      w ? '=IFERROR(AVERAGEIF(' + T + ',">0",' + P + w + '2:' + w + '),"")' : '–',
      '=IFERROR(AVERAGEIF(' + T + ',">0",' + P + h + '2:' + h + '),"")']);
  });
  const D = P + 'D2:D', F = P + 'F2:F';
  rows.push(['🏁 Celá hra (utekli)',
    '=COUNTIF(' + D + ',"✅ utekl")',
    '=IFERROR(AVERAGEIF(' + D + ',"✅ utekl",' + F + '),"")',
    '=IFERROR(MEDIAN(FILTER(' + F + ',' + D + '="✅ utekl")),"")',
    '', '=IFERROR(AVERAGEIF(' + D + ',"✅ utekl",' + P + 'H2:H),"")']);
  // Tabulka v češtině (cs_CZ) chce středník jako oddělovač argumentů
  const semi = !/^en/.test(ss.getSpreadsheetLocale());
  rows.forEach((row, i) => row.forEach((v, j) => {
    const c = sh.getRange(i + 1, j + 1);
    if (typeof v === 'string' && v.charAt(0) === '=') c.setFormula(semi ? v.replace(/,/g, ';') : v); else c.setValue(v);
  }));
  sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange(2, 3, rows.length - 1, 2).setNumberFormat('[m]:ss');
  sh.getRange(2, 5, rows.length - 1, 2).setNumberFormat('0.0');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 6);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Ruční údržba: znovu vytvoří list Statistika (spustit z editoru). */
function obnovitStatistiku() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const st = ss.getSheetByName('Statistika');
  if (st) ss.deleteSheet(st);
  ensureStats_(ss);
}
