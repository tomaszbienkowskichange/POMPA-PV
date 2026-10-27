/**
 * Monitor PC+PV — HUB danych (Warmlink + SolaX) w Google Apps Script.
 * Działa w chmurze Google (komputer niepotrzebny), dane trzyma w tym Arkuszu, a aplikacja
 * pobiera je jednym adresem (Web App). Instrukcja: INSTRUKCJA_HUB.md
 */
var TZ = 'Europe/Warsaw';
var VERSION = '1.8';
var WL_BASE = 'https://cloud.linked-go.com:449/crmservice/api/app';   // Warmlink 3.x (port 449)
var WL_APP_ID = '16';
var WL_CODES = ['Power', 'Mode', 'R01', 'R02', 'R03', 'R08', 'R09', 'T01', 'T02', 'T03', 'T04', 'T05', 'T06', 'T07',
                'T08', 'T09', 'T10', 'T11', 'T12', 'T30', 'T32', 'T34', 'T35', 'T39', 'O1', 'O01', 'O10'];   // O1 i O01: w Warmlinku kod sprężarki bywa zapisany bez zera — pytamy o oba warianty
var WL_CODES_SAFE = WL_CODES.filter(function (c) { return !/^O/.test(c); });   // lista awaryjna, gdyby Warmlink odrzucał nieznane kody
var WL_COLS = ['T01', 'T02', 'T04', 'T30', 'T34', 'T35', 'T39', 'T08', 'Mode', 'Power', 'T32', 'T03', 'T05', 'T06', 'T07', 'T09', 'T10', 'T11', 'T12', 'O01', 'O10'];   // O01 = sprężarka ON/OFF, O10 = zawór CO/CWU (otwarty = CWU); nowe kolumny na końcu; T32 na końcu: stare wiersze zostają czytelne
var OA_BASE = 'https://openapi-eu.solaxcloud.com';                    // SolaX Developer Platform (UE)
var PV_ON_W = 50;                                                      // moc [W], od której uznajemy, że PV pracuje
var KEEP_WL_DAYS = 365, KEEP_SX_DAYS = 900;
var WL_RUN_MIN = 10;                                                   // Warmlink: zapis co tyle minut, gdy sprężarka pracuje (stabilna praca)
var WL_FAST_MIN = 5;                                                   // …a co tyle: pierwsze 30 min pracy, po zmianie (start/stop/skok Hz/odszranianie) i przy mrozie
var WL_IDLE_MIN = 5;                                                   // postój: tylko „czy ruszyła?” co tyle minut — NIC nie zapisujemy do arkusza
var WL_IDLE_STORE_MIN = 60;                                            // postój: jeden wiersz „tętna” na godzinę (temp. zewnętrzna itp.)
var WL_COLD_C = 3;                                                     // T04 ≤ tyle °C → częste odszranianie → gęstsze próbkowanie podczas pracy
var SX_EVERY_MIN = 60;                                                 // falownik SolaX: odczyt na żywo co tyle minut (domyślnie co godzinę; Warmlink dalej co 5 min)
var DAY_HDR = ['data', 'produkcja_kWh', 'szczyt_W', 'pierwsza_praca', 'ostatnia_praca', 'probki', 'import_kWh', 'eksport_kWh', 'zuzycie_kWh'];
var BF_MAX_DAYS = 360;                                                 // SolaX: history_data tylko do roku wstecz

/* ---------- drobiazgi ---------- */
function P_() { return PropertiesService.getScriptProperties(); }
function prop_(k) { return P_().getProperty(k); }
function nowStr_(d) { return Utilities.formatDate(d || new Date(), TZ, "yyyy-MM-dd'T'HH:mm"); }
function num_(x) { if (x === null || x === undefined || x === '') return null; var n = parseFloat(String(x).replace(',', '.')); return isNaN(n) ? null : n; }
function hex_(bytes) { return bytes.map(function (b) { var v = b < 0 ? b + 256 : b; return (v < 16 ? '0' : '') + v.toString(16); }).join(''); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function ss_() { return SpreadsheetApp.openById(prop_('SHEET_ID')); }
function sheet_(name, hdr) {
  var s = ss_(), sh = s.getSheetByName(name);
  if (!sh) { sh = s.insertSheet(name); sh.appendRow(hdr); sh.getRange('A:A').setNumberFormat('@'); sh.setFrozenRows(1); }
  else if (!sh.getRange(2, 1).getNumberFormat || sh.getRange(2, 1).getNumberFormat() !== '@') { try { sh.getRange('A:A').setNumberFormat('@'); } catch (e) {} }   // nowe wiersze zawsze jako tekst
  return sh;
}
function onoff_(x) { if (x === undefined || x === null || x === '') return ''; var t = String(x).trim().toLowerCase(); if (/^(1|on|open|opened|opening|true|run|running|start|started|enable|enabled|work|working)$/.test(t)) return 1; if (/^(0|off|close|closed|closing|false|stop|stopped|disable|disabled|idle|standby)$/.test(t)) return 0; var n = num_(x); return n === null ? '' : (n > 0 ? 1 : 0); }   // ON/OFF, otwarty/zamknięty → 1/0
function wnum_(x) { if (x === undefined || x === null || x === '') return ''; var n = num_(x); return n === null ? x : n; }   // zapisuj LICZBY: Arkusz (PL) zamienia tekst „3.5” na datę 3 maja
function wlSheet_() { var hdr = ['ts'].concat(WL_COLS), sh = sheet_('WL', hdr); try { if (sh.getLastColumn() < hdr.length || String(sh.getRange(1, hdr.length).getValue()) !== hdr[hdr.length - 1]) sh.getRange(1, 1, 1, hdr.length).setValues([hdr]); } catch (e) {} return sh; }
function sxSheet_() { return sheet_('SX', ['ts', 'pw_W', 'yield_today_kWh', 'yield_total_kWh']); }
function daySheet_() {
  var sh = sheet_('DNI', DAY_HDR);
  if (prop_('FMT_DNI') !== '1') { try { sh.getRange('D:E').setNumberFormat('@'); P_().setProperty('FMT_DNI', '1'); } catch (e) {} }   // godziny jako tekst (inaczej Arkusz robi z nich „1899-12-30”)
  return sh;
}
function migrate_() {                                  // jednorazowo po aktualizacji skryptu: nagłówki arkuszy WL i DNI
  var s = ss_(), a = s.getSheetByName('WL'), d = s.getSheetByName('DNI');
  if (a) a.getRange(1, 1, 1, WL_COLS.length + 1).setValues([['ts'].concat(WL_COLS)]);
  if (d) d.getRange(1, 1, 1, DAY_HDR.length).setValues([DAY_HDR]);
  try { repairDates_(); } catch (e) { errLog_('FIX', e.message); }
  P_().setProperty('HDR_V', VERSION);
}
function repairDates_() {                              // kolumna A jako tekst + zamiana dat na tekst, scalenie duplikatów dni w arkuszu DNI
  var s = ss_(), rep = {};
  ['WL', 'SX', 'DNI'].forEach(function (nm) {
    var sh = s.getSheetByName(nm); if (!sh) return;
    var n = sh.getLastRow() - 1; if (n < 1) return;
    var col = sh.getRange(2, 1, n, 1), v = col.getValues(), ch = 0;
    col.setNumberFormat('@');
    var out = v.map(function (r) { var x = r[0]; if (x instanceof Date) { ch++; return [nm === 'DNI' ? dk_(x) : tsk_(x)]; } return [nm === 'DNI' ? dk_(x) : String(x)]; });
    if (ch || nm === 'DNI') col.setValues(out);
    rep[nm] = ch;
  });
  var dz = s.getSheetByName('DNI');                     // godziny pracy PV (kolumny D:E) → tekst „HH:mm” (naprawia „1899-12-30”)
  if (dz && dz.getLastRow() > 1) {
    var n2 = dz.getLastRow() - 1, rg = dz.getRange(2, 4, n2, 2), vv = rg.getValues(), fixed = 0;
    var tv = vv.map(function (r) { if (r[0] instanceof Date || r[1] instanceof Date || typeof r[0] === 'number' || typeof r[1] === 'number') fixed++; return [hm_(r[0]), hm_(r[1])]; });
    rg.setNumberFormat('@'); rg.setValues(tv); rep.hm = fixed; P_().setProperty('FMT_DNI', '1');
  }
  var d = s.getSheetByName('DNI');                      // duplikaty dni (powstawały, gdy data była datą, a nie tekstem) → jeden wiersz na dzień
  if (d && d.getLastRow() > 2) {
    var m = {}, rows = d.getRange(2, 1, d.getLastRow() - 1, DAY_HDR.length).getValues(), dup = 0;
    rows.forEach(function (r) { var k = dk_(r[0]); if (!k) return; var c = m[k];
      if (!c) { m[k] = r.slice(); m[k][0] = k; return; } dup++;
      for (var i = 1; i < DAY_HDR.length; i++) { var a = num_(c[i]), b = num_(r[i]); if (i === 1 || i === 2 || i === 5) c[i] = Math.max(a || 0, b || 0) || ''; else if (c[i] === '' && r[i] !== '') c[i] = r[i]; else if (i === 3 && r[i] && (!c[i] || r[i] < c[i])) c[i] = r[i]; else if (i === 4 && r[i] && r[i] > c[i]) c[i] = r[i]; } });
    if (dup) { var keys = Object.keys(m).sort(), vals = keys.map(function (k) { return m[k]; });
      d.getRange(2, 1, d.getLastRow() - 1, DAY_HDR.length).clearContent();
      d.getRange(2, 1, vals.length, DAY_HDR.length).setValues(vals); rep.dup = dup; }
  }
  return rep;
}
function naprawDaty() {                                // menu: „5. Napraw daty w arkuszach”
  var r = repairDates_(), ui = SpreadsheetApp.getUi();
  ui.alert('Naprawa dat', 'Zamienione daty na tekst — WL: ' + (r.WL || 0) + ', SX: ' + (r.SX || 0) + ', DNI: ' + (r.DNI || 0) + '. Naprawione godziny pracy (1899-12-30): ' + (r.hm || 0) + '. Scalone duplikaty dni: ' + (r.dup || 0) + '.', ui.ButtonSet.OK);
}
function errLog_(k, msg) { P_().setProperty('ERR_' + k, nowStr_() + ' ' + String(msg).slice(0, 300)); }

/* ---------- Warmlink (nieoficjalne API chmury Linked-Go) ---------- */
function wlCall_(path, body, token) {
  var res = UrlFetchApp.fetch(WL_BASE + '/' + path + '?lang=en', {
    method: 'post', contentType: 'application/json; charset=utf-8', payload: JSON.stringify(body),
    headers: token ? { 'x-token': token } : {}, muteHttpExceptions: true,
    validateHttpsCertificates: false          // chmura Warmlink używa certyfikatu, którego standardowa weryfikacja nie przyjmuje
  });
  var txt = res.getContentText();
  try { return JSON.parse(txt); } catch (e) { return { _http: res.getResponseCode(), _body: String(txt).slice(0, 200) }; }
}
function wlLogin_() {
  var user = prop_('WL_USER'), pw = prop_('WL_PASS'), last = null;
  if (!user || !pw) throw new Error('Brak loginu/hasła Warmlink (menu: 1. Konfiguracja).');
  var md5 = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, pw, Utilities.Charset.UTF_8));
  var variants = [md5, pw];
  for (var i = 0; i < variants.length; i++) {
    var r = wlCall_('user/login', { userName: user, password: variants[i], loginSource: 'IOS', areaCode: 'en', appId: WL_APP_ID, type: '2' });
    last = r;
    var tok = r && r.objectResult && typeof r.objectResult === 'object' ? r.objectResult['x-token'] : null;
    if (tok) { P_().setProperty('WL_TOKEN', tok); return tok; }
  }
  throw new Error('Logowanie Warmlink nie powiodło się: ' + JSON.stringify(last).split(user).join('***').slice(0, 200));
}
function wlOk_(r) { return r && !r._http && String(r.error_code) === '0' && Array.isArray(r.objectResult); }
function wlValues_() {
  var code = prop_('WL_DEVICE'); if (!code) throw new Error('Brak kodu urządzenia Warmlink (menu: 1. Konfiguracja).');
  var tok = prop_('WL_TOKEN') || wlLogin_(), res = null, tries = [WL_CODES, WL_CODES, WL_CODES_SAFE], safe = false;
  for (var a = 0; a < tries.length; a++) {
    res = wlCall_('device/getDataByCode', { appId: WL_APP_ID, deviceCode: code, protocalCodes: tries[a] }, tok);
    if (wlOk_(res)) { safe = (a === 2); break; }
    if (a < 2) tok = wlLogin_();                  // token wygasł albo aplikacja wylogowała to konto; potem próba z listą bez kodów O*
  }
  if (!wlOk_(res)) throw new Error('Odczyt Warmlink nieudany: ' + JSON.stringify(res).slice(0, 200));
  var v = {}, oSeen = [];
  res.objectResult.forEach(function (r) {
    if (!r || !r.code) return;
    var raw = String(r.code).trim(), cn = raw, mm = cn.match(/^[oO]0*(\d{1,2})$/);
    if (mm) { cn = 'O' + ('0' + mm[1]).slice(-2); oSeen.push(raw + '=' + JSON.stringify(r.value)); }   // „O1” / „o01” → „O01”
    if (v[cn] === undefined || v[cn] === null || v[cn] === '') v[cn] = r.value;
  });
  try { P_().setProperty('WL_O_RAW', nowStr_() + ' · O01=' + JSON.stringify(v.O01) + ' O10=' + JSON.stringify(v.O10) + ' · Warmlink zwrócił ' + res.objectResult.length + ' kodów' + (safe ? ' (lista bez O*: Warmlink odrzucił kody O)' : '') + (oSeen.length ? ' · kody O: ' + oSeen.join(', ').slice(0, 200) : ' · brak kodów O1/O01/O10 w odpowiedzi')); } catch (e) {}   // diagnostyka widoczna w aplikacji (Na żywo)
  ['O01', 'O10'].forEach(function (c) { if (v[c] !== undefined) v[c] = onoff_(v[c]); });
  return v;
}

/* ---------- SolaX Developer Platform (OAuth2, client credentials) ---------- */
function sxToken_(force) {
  var exp = +prop_('SX_EXP') || 0;
  if (!force && prop_('SX_TOK') && Date.now() < exp - 3600000) return prop_('SX_TOK');
  var res = UrlFetchApp.fetch(OA_BASE + '/openapi/auth/oauth/token', {
    method: 'post', payload: { client_id: prop_('SX_ID'), client_secret: prop_('SX_SECRET'), grant_type: 'client_credentials' }, muteHttpExceptions: true
  });
  var j = JSON.parse(res.getContentText()), r = j.result || {};
  if ((j.code !== 0 && j.code !== 10000) || !r.access_token) throw new Error('Solax token: ' + JSON.stringify(j).slice(0, 200));
  P_().setProperty('SX_TOK', r.access_token); P_().setProperty('SX_EXP', String(Date.now() + (parseFloat(r.expires_in) || 3600) * 1000));
  return r.access_token;
}
function sxCall_(path, body) {
  for (var a = 0; a < 2; a++) {
    var opt = { headers: { Authorization: 'bearer ' + sxToken_(a === 1), Accept: '*/*' }, muteHttpExceptions: true };
    if (body) { opt.method = 'post'; opt.contentType = 'application/json'; opt.payload = JSON.stringify(body); }
    var res = null, lastE = null;
    for (var t = 0; t < 3 && !res; t++) {                 // błędy sieciowe Google („Niedostępny adres”, timeout) bywają chwilowe → 3 próby
      try { res = UrlFetchApp.fetch(OA_BASE + path, opt); } catch (ne) { lastE = ne; Utilities.sleep(1500 * (t + 1)); }
    }
    if (!res) { var ne2 = new Error('SolaX niedostępny z Google (' + String(lastE && lastE.message || lastE).slice(0, 160) + ')'); ne2.net = true; throw ne2; }
    var j;
    try { j = JSON.parse(res.getContentText()); } catch (e) { throw new Error('Solax HTTP ' + res.getResponseCode() + ': ' + String(res.getContentText()).slice(0, 120)); }
    if (j.code === 10402 && a === 0) continue;
    if (j.code === 10405 || j.code === 10406) { var er = new Error('Solax: limit zapytań (' + j.code + ') — spróbuję później'); er.limit = true; throw er; }
    if (j.code !== 0 && j.code !== 10000) { var e2 = new Error('Solax: ' + JSON.stringify(j).slice(0, 250)); if (/past year|within the past|too early|out of range/i.test(String(j.message))) e2.range = true; throw e2; }
    return j;
  }
}
function findKey_(o, keys) {                       // pierwsza wartość pod jednym z kluczy (rekurencyjnie, bez rozróżniania wielkości liter)
  var set = {}; keys.forEach(function (k) { set[k.toLowerCase()] = 1; });
  function walk(x) {
    if (x && typeof x === 'object') {
      if (!Array.isArray(x)) for (var k in x) { if (set[k.toLowerCase()] && x[k] !== null && typeof x[k] !== 'object') return x[k]; }
      for (var k2 in x) { var r = walk(x[k2]); if (r !== undefined && r !== null) return r; }
    }
    return null;
  }
  return walk(o);
}
function sxFindDevice_() {                              // plantId, businessType i numer falownika (jednorazowo, zapamiętane)
  if (prop_('SX_PLANT') && prop_('SX_DEVSN')) return;
  for (var bt = 1; bt <= 4; bt += 3) {                  // 1 = Residential, 4 = C&I
    var lst = sxCall_('/openapi/v2/device/page_device_info?deviceType=1&businessType=' + bt);
    var recs = (lst.result && lst.result.records) || [], want = prop_('SX_SN');
    if (want) recs = recs.filter(function (r) { return r.deviceSn === want; }).concat(recs);
    if (recs.length && recs[0].deviceSn) {
      P_().setProperty('SX_PLANT', String(recs[0].plantId || '')); P_().setProperty('SX_BT', String(bt)); P_().setProperty('SX_DEVSN', String(recs[0].deviceSn)); return;
    }
  }
  throw new Error('Solax: nie znaleziono falownika (page_device_info).');
}
function sxUnit_() { return prop_('SX_BT') === '4' ? 1000 : 1; }          // businessType 4: moce w kW, 1: w W → zawsze zwracamy W
function sumKeys_(o, re) { var s = 0, any = false; if (o) for (var k in o) { if (re.test(k) && num_(o[k]) !== null) { s += num_(o[k]); any = true; } } return any ? s : null; }
function sxPvW_(rec) {                                  // moc PV [W] z rekordu falownika (realtime lub history_data)
  var u = sxUnit_(), p = num_(rec.MPPTTotalInputPower);
  if (p === null) p = sumKeys_(rec.pvMap, /^pv\d+Power$/i);
  if (p === null) p = sumKeys_(rec.mpptMap, /^mppt\d+Power$/i);
  if (p === null) { var a = [rec.acPower1, rec.acPower2, rec.acPower3].map(num_).filter(function (x) { return x !== null; }); if (a.length) p = a.reduce(function (x, y) { return x + y; }, 0); }
  return p === null ? null : Math.round(p * u);
}
function sxValues_() {
  sxFindDevice_();
  var res = { yt: null, tot: null, pw: null, up: null };
  var dev = sxCall_('/openapi/v2/device/realtime_data?snList=' + encodeURIComponent(prop_('SX_DEVSN')) + '&deviceType=1&businessType=' + prop_('SX_BT')).result;
  var rec = Array.isArray(dev) ? dev[0] : dev;
  if (rec) { res.pw = sxPvW_(rec); res.yt = num_(rec.dailyYield); res.tot = num_(rec.totalYield); res.up = rec.plantLocalTime || null; }
  if ((res.yt === null || res.tot === null) && prop_('SX_PLANT')) {                       // produkcja z poziomu instalacji
    try { var plant = sxCall_('/openapi/v2/plant/realtime_data?plantId=' + prop_('SX_PLANT') + '&businessType=' + prop_('SX_BT')).result || {};
      if (res.yt === null) res.yt = num_(plant.dailyYield); if (res.tot === null) res.tot = num_(plant.totalYield); } catch (e) { errLog_('SXPL', e.message); }
  }
  if (res.yt === null && res.tot === null) throw new Error('Solax: odpowiedź bez produkcji (sprawdź numer falownika i typ instalacji).');
  return res;
}

/* ---------- SolaX: dane historyczne (jak w aplikacji mobilnej) ---------- */
function dayStartMs_(day) { return Utilities.parseDate(day + ' 00:00:00', TZ, 'yyyy-MM-dd HH:mm:ss').getTime(); }
function addDays_(day, n) { var d = new Date(dayStartMs_(day) + n * 864e5 + 12 * 3600000); return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function sxHistDay_(day) {                              // próbki co 5 min z jednego dnia → [[ts, pw_W, yield_today, yield_total], ...]
  sxFindDevice_();
  var t0 = dayStartMs_(day), out = [];
  var nowMs = Date.now();
  for (var h = 0; h < 2; h++) {                         // history_data: okno maks. 12 h
    var a = t0 + h * 12 * 3600000, b = h === 0 ? a + 12 * 3600000 - 1 : t0 + 24 * 3600000 - 1;
    if (a > nowMs) break; if (b > nowMs) b = nowMs;                // dzień bieżący: tylko do teraz
    var j = sxCall_('/openapi/v2/device/history_data?snList=' + encodeURIComponent(prop_('SX_DEVSN')) + '&deviceType=1&businessType=' + prop_('SX_BT') + '&startTime=' + a + '&endTime=' + b + '&timeInterval=5');
    (j.result || []).forEach(function (r) {
      var ts = String(r.plantLocalTime || '').replace(' ', 'T').slice(0, 16); if (ts.length < 16 || ts.slice(0, 10) !== day) return;
      out.push([ts, sxPvW_(r), num_(r.dailyYield), num_(r.totalYield)]);
    });
  }
  out.sort(function (x, y) { return x[0] < y[0] ? -1 : x[0] > y[0] ? 1 : 0; });
  return out;
}
function sxStatMonth_(ym) {                             // statystyki dzienne miesiąca (produkcja, import, eksport, zużycie)
  sxFindDevice_();
  if (!prop_('SX_PLANT')) throw new Error('Solax: brak plantId — statystyki dzienne niedostępne.');
  var j = sxCall_('/openapi/v2/plant/energy/get_stat_data', { plantId: prop_('SX_PLANT'), dateType: '2', date: ym, businessType: prop_('SX_BT') });
  return ((j.result && j.result.plantEnergyStatDataList) || []).map(function (r) {
    return { d: String(r.date).slice(0, 10), y: num_(r.pvGeneration), im: num_(r.importEnergy), ex: num_(r.exportEnergy), ld: num_(r.loadConsumption) };
  });
}
function withLock_(ms, fn) { var lock = LockService.getScriptLock(); if (!lock.tryLock(ms)) throw new Error('hub zajęty — spróbuj za chwilę'); try { return fn(); } finally { lock.releaseLock(); } }
function sortSheet_(sh) { var n = sh.getLastRow() - 1; if (n > 1) sh.getRange(2, 1, n, sh.getLastColumn()).sort({ column: 1, ascending: true }); }
function sxDayCounts_() {                               // ile próbek ma każdy dzień w arkuszu SX
  var col = allCol_(sxSheet_()), c = {}; col.forEach(function (r) { var d = String(r[0]).slice(0, 10); c[d] = (c[d] || 0) + 1; }); return c;
}
/* Dzienny licznik falownika (dailyYield) rośnie w ciągu dnia, ale po północy chmura SolaX potrafi aż do wschodu słońca zwracać jeszcze WCZORAJSZĄ wartość końcową.
   Taki „stary” odczyt zawyżał sumę nowego dnia (max() nigdy się nie poprawiał). Gdy licznik później spada (reset), liczy się tylko odcinek po ostatnim spadku. */
function yieldOfRows_(rows) { var seg = 0, prev = null; rows.forEach(function (r) { var v = num_(r[2]); if (v === null) return; if (prev !== null && prev - v > 0.3) seg = 0; if (v > seg) seg = v; prev = v; }); return seg; }
function writeSxDay_(day, rows) {                       // dopisz brakujące próbki dnia i uzupełnij arkusz DNI
  if (!rows.length) return 0;
  var sh = sxSheet_(), have = {}, add = [];
  allCol_(sh).forEach(function (r) { var t = String(r[0]); if (t.slice(0, 10) === day) have[t] = 1; });
  rows.forEach(function (r) { if (!have[r[0]]) add.push([r[0], r[1] === null ? '' : r[1], r[2] === null ? '' : r[2], r[3] === null ? '' : r[3]]); });
  if (add.length) { sh.getRange(sh.getLastRow() + 1, 1, add.length, 4).setValues(add); sortSheet_(sh); }
  var y = 0, pk = 0, h0 = '', h1 = '', prev = null;
  rows.forEach(function (r) { var on = r[1] !== null ? r[1] >= PV_ON_W : (prev !== null && r[2] !== null && r[2] > prev + 0.0001); prev = r[2];
    if (r[1] !== null) pk = Math.max(pk, r[1]); if (on) { var t = r[0].slice(11); if (!h0) h0 = t; h1 = t; } });
  y = yieldOfRows_(rows);
  setDay_(day, { y: y, pk: pk, h0: h0, h1: h1, n: rows.length, exact: true });
  return add.length;
}
function setDay_(day, o) {                              // wstaw/uzupełnij wiersz dnia w arkuszu DNI (nie psuje danych z próbek)
  var sh = daySheet_(), col = allCol_(sh), idx = -1;
  for (var i = col.length - 1; i >= 0; i--) { if (dk_(col[i][0]) === day) { idx = i + 2; break; } }
  var cur = idx > 0 ? sh.getRange(idx, 1, 1, DAY_HDR.length).getValues()[0] : [day, '', '', '', '', '', '', '', ''];
  cur[3] = hm_(cur[3]); cur[4] = hm_(cur[4]);
  function pick(a, b) { return (b !== undefined && b !== null && b !== '') ? b : a; }
  var y = (o.exact && num_(o.y) !== null) ? num_(o.y) : Math.max(num_(cur[1]) || 0, num_(o.y) || 0);      // krzywa 5 min jest dokładna: nadpisuje także zawyżoną sumę
  var pk0 = Math.max(num_(cur[2]) || 0, num_(o.pk) || 0);
  var row = [day, y || '', pk0 || '', o.exact ? pick(cur[3], o.h0) : pick(cur[3], o.h0), o.exact ? pick(cur[4], o.h1) : pick(cur[4], o.h1), (+cur[5] || 0) > (o.n || 0) ? cur[5] : pick(cur[5], o.n), pick(cur[6], o.im), pick(cur[7], o.ex), pick(cur[8], o.ld)];
  if (o.exact && o.h0) { row[3] = (cur[3] && cur[3] < o.h0) ? cur[3] : o.h0; row[4] = (cur[4] && cur[4] > o.h1) ? cur[4] : o.h1; }   // krzywa 5 min jest dokładniejsza niż próbki godzinowe
  if (idx > 0) sh.getRange(idx, 1, 1, DAY_HDR.length).setValues([row]);
  else { sh.appendRow(row); sortSheet_(sh); }
}
function setDaysBulk_(list, today) {                   // hurtowe uzupełnienie arkusza DNI statystykami dziennymi (jeden odczyt, jeden zapis)
  var sh = daySheet_(), n = sh.getLastRow() - 1, map = {};
  if (n > 0) sh.getRange(2, 1, n, DAY_HDR.length).getValues().forEach(function (r) { var k = dk_(r[0]); r[0] = k; r[3] = hm_(r[3]); r[4] = hm_(r[4]); map[k] = r; });
  list.forEach(function (o) {
    if (o.d > today || (o.y === null && o.im === null)) return;
    var cur = map[o.d] || [o.d, '', '', '', '', '', '', '', ''], y = (num_(o.y) > 0) ? num_(o.y) : (num_(cur[1]) || 0);      // statystyka dzienna SolaX jest dokładna — zastępuje wcześniejszą (np. zawyżoną) sumę
    map[o.d] = [o.d, y || '', cur[2], cur[3], cur[4], cur[5], o.im !== null ? o.im : cur[6], o.ex !== null ? o.ex : cur[7], o.ld !== null ? o.ld : cur[8]];
  });
  var rows = Object.keys(map).sort().map(function (k) { return map[k]; });
  if (rows.length) sh.getRange(2, 1, rows.length, DAY_HDR.length).setValues(rows);
}
function sxEnsureDay_(day) {                            // na żądanie aplikacji: dociągnij jeden dzień z SolaX, jeśli brakuje próbek
  var today = nowStr_().slice(0, 10);
  if (day > today) return 0;
  if (day === today) { var tf = prop_('SX_TFILL'); if (tf && minDiff_(tf, nowStr_()) < 8) return 0; }          // dziś: dociągaj krzywą 5 min najwyżej co 8 min
  else if ((sxDayCounts_()[day] || 0) >= 150) return 0;
  var rows = sxHistDay_(day);
  if (day === today) P_().setProperty('SX_TFILL', nowStr_());
  return withLock_(25000, function () { return writeSxDay_(day, rows); });
}
function bfState_() {
  var st = prop_('BF_STAT') || '', dy = prop_('BF_DAY') || '', mY = prop_('BF_MINYM') || '', mD = prop_('BF_MINDAY') || '';
  return { stat: st, day: dy, done: prop_('BF_DONE') === '1', err: prop_('ERR_BF') || null, statDone: !!mY && st < mY, dayDone: !!mD && dy < mD, minDay: mD, skipped: prop_('BF_SKIPPED') || '' };
}
function dataStats_() {                                 // co faktycznie leży w arkuszach (diagnostyka widoczna w aplikacji)
  var out = {};
  try { var d = daySheet_(), n = d.getLastRow() - 1, k = [], withY = 0;
    if (n > 0) d.getRange(2, 1, n, 2).getValues().forEach(function (r) { k.push(dk_(r[0])); if (num_(r[1]) > 0) withY++; });
    k.sort(); out.dni = { n: k.length, from: k[0] || '', to: k[k.length - 1] || '', withY: withY };
  } catch (e) { out.dni = { err: e.message }; }
  try { var x = sxSheet_(), m = x.getLastRow() - 1;
    out.sx = m > 0 ? { n: m, from: dk_(x.getRange(2, 1).getValue()), to: dk_(x.getRange(x.getLastRow(), 1).getValue()) } : { n: 0 };
  } catch (e2) { out.sx = { err: e2.message }; }
  return out;
}
function backfillRun_(maxMs, maxCalls) {                // porcja pobierania historii; wznawia od zapamiętanego miejsca
  var t0 = Date.now(), calls = 0, today = nowStr_().slice(0, 10), msg = '', errs = [];
  function budget() { return Date.now() - t0 < maxMs && calls < maxCalls; }
  function phase(name, fn) {                            // błąd zakresu dat kończy fazę; limit zapytań przerywa całość; inny błąd nie blokuje drugiej fazy
    try { fn(); } catch (e) {
      if (e.limit) throw e;
      if (e.range) { P_().setProperty(name === 'stat' ? 'BF_STAT' : 'BF_DAY', name === 'stat' ? '0000-00' : '0000-00-00'); }
      else errs.push(name + ': ' + e.message);
    }
  }
  try {
    sxFindDevice_();
    if (!prop_('BF_STAT')) P_().setProperty('BF_STAT', today.slice(0, 7));
    if (!prop_('BF_DAY')) P_().setProperty('BF_DAY', addDays_(today, -1));
    var created = prop_('SX_CREATED');
    if (!created) { try { var pl = sxCall_('/openapi/v2/plant/page_plant_info?businessType=' + prop_('SX_BT')).result; var ct = pl && pl.records && pl.records[0] && pl.records[0].createTime;
        created = ct ? String(ct).slice(0, 10) : '2000-01-01'; } catch (e) { created = '2000-01-01'; } P_().setProperty('SX_CREATED', created); }
    var yearAgo = addDays_(today, -BF_MAX_DAYS), minDay = yearAgo < created ? created : yearAgo;
    var minYm = (created > yearAgo ? created : yearAgo).slice(0, 7);                // API statystyk: tylko ostatni rok
    P_().setProperty('BF_MINYM', minYm); P_().setProperty('BF_MINDAY', minDay);
    phase('stat', function () {                                                     // 1) statystyki dzienne miesiąc po miesiącu (wstecz)
      while (budget() && prop_('BF_STAT') >= minYm) {
        var ym = prop_('BF_STAT'); calls++;
        var st = sxStatMonth_(ym);
        withLock_(25000, function () { setDaysBulk_(st, today); });
        var y = +ym.slice(0, 4), m = +ym.slice(5, 7) - 1; if (m === 0) { y--; m = 12; }
        P_().setProperty('BF_STAT', y + '-' + (m < 10 ? '0' : '') + m);
      }
    });
    phase('day', function () {                                                      // 2) krzywa mocy co 5 min dzień po dniu (wstecz)
      var cnt = null;
      while (budget() && prop_('BF_DAY') >= minDay) {
        var day = prop_('BF_DAY'); if (!cnt) cnt = sxDayCounts_();
        if ((cnt[day] || 0) < 150) { calls += 2;
          try { var rows = sxHistDay_(day); withLock_(25000, function () { writeSxDay_(day, rows); }); P_().deleteProperty('BF_FAILN'); }
          catch (de) {
            if (de.limit || de.range) throw de;
            var fd = prop_('BF_FAILDAY') || '', fn = fd === day ? (parseInt(prop_('BF_FAILN'), 10) || 0) + 1 : 1;
            P_().setProperty('BF_FAILDAY', day); P_().setProperty('BF_FAILN', String(fn));
            if (fn >= 3) { skipDay_(day); P_().deleteProperty('BF_FAILN'); errs.push('pominięto ' + day + ' po 3 nieudanych próbach'); continue; }
            throw de;
          } }
        P_().setProperty('BF_DAY', addDays_(day, -1));
      }
    });
    if (prop_('BF_STAT') < minYm && prop_('BF_DAY') < minDay) { P_().setProperty('BF_DONE', '1'); msg = 'gotowe'; }
    if (errs.length) { P_().setProperty('ERR_BF', nowStr_() + ' ' + errs.join(' | ').slice(0, 250)); msg = errs.join(' | '); } else P_().deleteProperty('ERR_BF');
  } catch (e) { P_().setProperty('ERR_BF', nowStr_() + ' ' + String(e.message).slice(0, 200)); msg = e.message; }
  return { done: prop_('BF_DONE') === '1', calls: calls, msg: msg, state: bfState_() };
}
function skipDay_(day) {                                // pomiń dzień w pobieraniu krzywej 5 min i zapamiętaj to
  var sk = (prop_('BF_SKIPPED') || '').split(',').filter(Boolean); if (sk.indexOf(day) < 0) sk.push(day);
  P_().setProperty('BF_SKIPPED', sk.slice(-20).join(','));
  if (prop_('BF_DAY') === day) P_().setProperty('BF_DAY', addDays_(day, -1));
}
function backfillTick() {                               // wyzwalacz co 10 min, aż do końca pobierania historii
  var r = backfillRun_(240000, 120);
  if (r.done) ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'backfillTick') ScriptApp.deleteTrigger(t); });
}
function startBackfill() {                              // menu: „4. Pobierz historię SolaX”
  ['BF_STAT', 'BF_DAY', 'BF_DONE', 'ERR_BF'].forEach(function (k) { P_().deleteProperty(k); });
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'backfillTick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('backfillTick').timeBased().everyMinutes(10).create();
  var r = backfillRun_(200000, 80);
  try { SpreadsheetApp.getUi().alert('Historia SolaX', 'Pierwsza porcja pobrana (' + r.calls + ' zapytań). Reszta dociąga się sama co 10 minut (do roku wstecz dla krzywej mocy co 5 min i do początku instalacji dla sum dziennych).\n\nStan: statystyki do ' + (r.state.stat || '—') + ', krzywa do ' + (r.state.day || '—') + (r.msg ? '\n' + r.msg : ''), SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) {}
}

/* ---------- zapis próbek ---------- *//* ---------- zapis próbek ---------- */
function lastRows_(sh, n) {
  var last = sh.getLastRow(); if (last < 2) return [];
  var from = Math.max(2, last - n + 1);
  return sh.getRange(from, 1, last - from + 1, sh.getLastColumn()).getValues();
}
function wlEvery_(key, def) { var n = parseInt(prop_(key), 10); return n > 0 ? n : def; }
function wlState_() { try { return JSON.parse(prop_('WL_STATE') || '{}') || {}; } catch (e) { return {}; } }
function wlDue_(force) {                               // czy teraz zapytać Warmlink (wyzwalacz tyka co 5 min, decyzja jest tu)
  if (force === true) return true;
  var last = null; try { last = JSON.parse(prop_('WL_LAST') || 'null'); } catch (e) { last = null; }
  if (!last || !last.ts) return true;
  var now = nowStr_(), age = minDiff_(last.ts, now), st = wlState_(), need;
  if (st.run) {
    var fast = (st.since && minDiff_(st.since, now) <= 30) || (st.fastUntil && st.fastUntil > now) || (st.t04 !== null && st.t04 !== undefined && st.t04 <= WL_COLD_C);
    need = fast ? WL_FAST_MIN : wlEvery_('WL_RUN', WL_RUN_MIN);
  } else need = wlEvery_('WL_IDLE', WL_IDLE_MIN);
  return age >= need - 2;
}
function pollWl_() {
  var v = wlValues_(), ts = nowStr_(), st = wlState_();
  var hz = num_(v.T30), t01 = num_(v.T01), t02 = num_(v.T02), t04 = num_(v.T04), o01 = num_(v.O01);
  var run = (hz !== null && hz > 0) || o01 === 1, prevRun = !!st.run, def = run && t01 !== null && t02 !== null && t02 < t01 - 1;
  var store;
  if (run) store = !st.lastStore || minDiff_(st.lastStore, ts) >= 3;          // praca: zapisuj (odstęp pilnuje wlDue_)
  else if (prevRun) store = true;                                              // właśnie się zatrzymała: zapisz moment stopu
  else store = !st.lastStore || minDiff_(st.lastStore, ts) >= WL_IDLE_STORE_MIN;   // postój: tylko „tętno” raz na godzinę
  var row = [ts]; WL_COLS.forEach(function (c) { row.push(wnum_(v[c])); });
  if (store) wlSheet_().appendRow(row);
  var event = run !== prevRun || def || (run && prevRun && st.hz !== null && st.hz !== undefined && hz !== null && Math.abs(hz - st.hz) >= 20);
  P_().setProperty('WL_STATE', JSON.stringify({ run: run, since: run ? (prevRun && st.since ? st.since : ts) : '', lastStore: store ? ts : (st.lastStore || ''), hz: hz, t04: t04,
    fastUntil: event ? nowStr_(new Date(Date.now() + 15 * 60000)) : (st.fastUntil || '') }));
  P_().setProperty('WL_LAST', JSON.stringify({ ts: ts, v: v }));
  return { ts: ts, v: v };
}
function pollSx_() {
  var r = sxValues_(), ts = nowStr_();
  sxSheet_().appendRow([ts, r.pw === null ? '' : r.pw, r.yt === null ? '' : r.yt, r.tot === null ? '' : r.tot]);
  var prev = null; try { prev = JSON.parse(prop_('SX_LAST') || 'null'); } catch (e) { prev = null; }
  var on = r.pw !== null ? r.pw >= PV_ON_W : (prev && prev.yt !== null && r.yt !== null && prev.ts.slice(0, 10) === ts.slice(0, 10) ? r.yt > prev.yt + 0.0001 : false);
  P_().setProperty('SX_LAST', JSON.stringify({ ts: ts, pw: r.pw, yt: r.yt, tot: r.tot, on: on }));
  updateDay_(ts, r.pw, r.yt, on);
  return { ts: ts, pw: r.pw, yt: r.yt, tot: r.tot, on: on };
}
function updateDay_(ts, pw, yt, on) {
  var sh = daySheet_(), day = ts.slice(0, 10), last = sh.getLastRow(), t = ts.slice(11);
  if (last >= 2 && dk_(sh.getRange(last, 1).getValue()) === day) {
    var r = sh.getRange(last, 1, 1, 6).getValues()[0]; r[3] = hm_(r[3]); r[4] = hm_(r[4]);
    var y0 = num_(r[1]) || 0, y = (yt !== null && y0 - yt > 0.3) ? yt : Math.max(y0, yt || 0), pk = Math.max(num_(r[2]) || 0, pw || 0);      // licznik dnia spadł = wcześniejsza wyższa wartość była przeniesiona z wczoraj
    sh.getRange(last, 2, 1, 5).setValues([[y, pk, r[3] || (on ? t.slice(0, 5) : ''), on ? t.slice(0, 5) : r[4], (+r[5] || 0) + 1]]);
  } else {
    var prevY = last >= 2 ? num_(sh.getRange(last, 2).getValue()) : null, y1 = yt || 0;
    if (y1 > 0 && prevY !== null && Math.abs(y1 - prevY) < 0.05 && !(pw !== null && pw >= PV_ON_W)) y1 = 0;      // pierwszy odczyt nowego dnia = końcowa wartość wczoraj, a falownik jeszcze nie produkuje → to nie jest produkcja nowego dnia
    sh.appendRow([day, y1, pw || 0, on ? t.slice(0, 5) : '', on ? t.slice(0, 5) : '', 1]);
  }
}
function sxEveryMin_() { var n = parseInt(prop_('SX_EVERY'), 10); return n > 0 ? n : SX_EVERY_MIN; }
function sxDue_(force) {                                    // czy minął interwał od ostatniego odczytu SolaX (-2 min zapasu na rozrzut wyzwalacza)
  if (force === true) return true;
  var last = null; try { last = JSON.parse(prop_('SX_LAST') || 'null'); } catch (e) { last = null; }
  return !last || !last.ts || minDiff_(last.ts, nowStr_()) >= sxEveryMin_() - 2 || last.ts.slice(0, 10) !== nowStr_().slice(0, 10) && minDiff_(last.ts, nowStr_()) >= 5;
}
function poll(force) {                                // wyzwalacz co 5 min (jego argument to obiekt zdarzenia, więc „force” musi być dokładnie true)
  var lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) return;
  try {
    if (prop_('SHEET_ID') && prop_('HDR_V') !== VERSION) { try { migrate_(); } catch (e) { errLog_('MIG', e.message); } }
    if (prop_('WL_USER') && wlDue_(force)) { try { pollWl_(); P_().deleteProperty('ERR_WL'); } catch (e) { errLog_('WL', e.message); } }
    if (prop_('SX_ID') && sxDue_(force)) { try { pollSx_(); P_().deleteProperty('ERR_SX'); } catch (e) { errLog_('SX', e.message); } }      // falownik rzadziej niż pompa (domyślnie co godzinę)
  } finally { lock.releaseLock(); }
}
function prune() {
  if (prop_('SX_ID')) { try { sxEnsureDay_(addDays_(nowStr_().slice(0, 10), -1)); } catch (e) { errLog_('SXFILL', e.message); } }   // nocą: domknij wczorajszy dzień krzywą 5 min
  [['WL', KEEP_WL_DAYS], ['SX', KEEP_SX_DAYS]].forEach(function (x) {
    var sh = ss_().getSheetByName(x[0]); if (!sh || sh.getLastRow() < 3) return;
    var cut = nowStr_(new Date(Date.now() - x[1] * 864e5)), col = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues(), k = 0;
    while (k < col.length && String(col[k][0]) < cut) k++;
    if (k > 0) sh.deleteRows(2, k);
  });
}

/* ---------- zapytania (Web App) ---------- */
function lowerIdx_(col, ts) { var lo = 0, hi = col.length; while (lo < hi) { var m = (lo + hi) >> 1; if (String(col[m][0]) < ts) lo = m + 1; else hi = m; } return lo; }
function minDiff_(a, b) { return Math.abs(Date.parse(a + ':00Z') - Date.parse(b + ':00Z')) / 60000; }
var SSTZ_ = null;
function ssTz_() { if (!SSTZ_) { try { SSTZ_ = ss_().getSpreadsheetTimeZone() || TZ; } catch (e) { SSTZ_ = TZ; } } return SSTZ_; }
function hm_(v) {                                      // godzina „HH:mm” z komórki (tekst, czas z Arkusza = Date z rokiem 1899, albo ułamek doby)
  if (v === null || v === undefined || v === '') return '';
  var h, m;
  if (v instanceof Date) { var t = Utilities.formatDate(v, ssTz_(), 'HH:mm'); return t === '00:00' ? '' : t; }
  if (typeof v === 'number') { var mins = Math.round((v % 1) * 1440); if (!mins) return ''; h = Math.floor(mins / 60) % 24; m = mins % 60; return (h < 10 ? '0' : '') + h + ':' + (m < 10 ? '0' : '') + m; }
  var x = /(\d{1,2}):(\d{2})/.exec(String(v)); if (!x) return '';
  return (x[1].length < 2 ? '0' : '') + x[1] + ':' + x[2];
}
function dk_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v === null || v === undefined ? '' : v).slice(0, 10); }       // klucz dnia z komórki (tekst albo data)
function tsk_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm") : String(v === null || v === undefined ? '' : v); }          // klucz czasu z komórki
function allCol_(sh) { return sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (r) { return [tsk_(r[0])]; }); }
function range_(sh, from, to, max) {
  var col = allCol_(sh); if (!col.length) return [];
  var i = lowerIdx_(col, from), j = lowerIdx_(col, to + '~');            // '~' > cyfry: włącznie z końcem dnia
  if (j - i > max) i = j - max;
  return j > i ? sh.getRange(i + 2, 1, j - i, sh.getLastColumn()).getValues() : [];
}
function wlObj_(row) { var o = { ts: tsk_(row[0]) }; WL_COLS.forEach(function (c, i) { var v = row[i + 1]; if (v === '' || v === null || v === undefined) { o[c] = null; return; } if (v instanceof Date) { o[c] = parseFloat(v.getDate() + '.' + (v.getMonth() + 1)); return; } var n = num_(v); o[c] = n === null ? (/^\d{4}-\d\d-\d\d/.test(String(v)) ? null : v) : n; }); return o; }   // data/tekst w kolumnie liczbowej → null (nie „2026”)
function sxOn_(pw, yt, prevYt) { return pw !== null ? pw >= PV_ON_W : (prevYt !== null && yt !== null ? yt > prevYt + 0.0001 : false); }
function sxRows_(rows) {
  var prev = null, prevDay = '';
  return rows.map(function (r) {
    var ts = String(r[0]), pw = num_(r[1]), yt = num_(r[2]), day = ts.slice(0, 10);
    var on = sxOn_(pw, yt, prevDay === day ? prev : null); prev = yt; prevDay = day;
    return [ts, pw, yt, on ? 1 : 0];
  });
}
function nearest_(sh, ts, tolMin) {
  var col = allCol_(sh); if (!col.length) return null;
  var i = lowerIdx_(col, ts), best = null, bd = 1e9;
  [i - 1, i].forEach(function (k) { if (k >= 0 && k < col.length) { var d = minDiff_(String(col[k][0]), ts); if (d < bd) { bd = d; best = k; } } });
  if (best === null || bd > tolMin) return null;
  return { row: sh.getRange(best + 2, 1, 1, sh.getLastColumn()).getValues()[0], dt: bd, idx: best + 2 };
}
function pvWindow_(f, t) {                             // produkcja PV w przedziale [f, t] oraz podczas pracy sprężarki (z krzywej 5 min SolaX + próbek Warmlink)
  if (f.length < 16 || t.length < 16 || t <= f) return { ok: false, error: 'zły przedział' };
  if (!prop_('SX_ID')) return { ok: false, error: 'SolaX nieskonfigurowany' };
  var errs = [], d = f.slice(0, 10), d1 = t.slice(0, 10), k = 0;
  while (d <= d1 && k++ < 4) { try { sxEnsureDay_(d); } catch (e) { errs.push(d + ': ' + e.message); } d = addDays_(d, 1); }
  var sx = range_(sxSheet_(), f, t, 3000).map(function (r) { return { ts: tsk_(r[0]), m: Date.parse(tsk_(r[0]) + ':00Z') / 60000, pw: num_(r[1]), yt: num_(r[2]) }; });
  var spanMin = (Date.parse(t + ':00Z') - Date.parse(f + ':00Z')) / 60000;
  if (!sx.length) return { ok: true, n: 0, errs: errs, spanMin: spanMin };
  var kInt = 0, covMin = 0, onMin = 0, pk = 0, i;
  for (i = 0; i + 1 < sx.length; i++) { var a = sx[i], b = sx[i + 1], dt = b.m - a.m;
    if (dt > 0 && dt <= 20 && a.pw !== null && b.pw !== null) { kInt += (a.pw + b.pw) / 2 * dt / 60000; covMin += dt; if (a.pw >= PV_ON_W) onMin += dt; } }
  sx.forEach(function (r) { if (r.pw !== null && r.pw > pk) pk = r.pw; });
  var kY = null, byDay = {};                                                    // różnica produkcji dziennej (licznik falownika) pierwsza→ostatnia próbka każdego dnia
  sx.forEach(function (r) { if (r.yt === null) return; var dd = r.ts.slice(0, 10), c = byDay[dd] || (byDay[dd] = { a: r.yt, b: r.yt }); if (r.yt < c.a) c.a = r.yt; if (r.yt > c.b) c.b = r.yt; });
  Object.keys(byDay).forEach(function (dd) { kY = (kY || 0) + (byDay[dd].b - byDay[dd].a); });
  var wl = range_(wlSheet_(), f, t, 3000).map(function (r) { return { ts: tsk_(r[0]), m: Date.parse(tsk_(r[0]) + ':00Z') / 60000, hz: num_(r[4]) }; });   // kolumna 4 = T30 (częstotliwość sprężarki)
  var runMin = 0, runKwh = 0, runCov = 0, j = 0;
  wl.forEach(function (w, ix) {
    var dt = Math.min(10, ix + 1 < wl.length ? wl[ix + 1].m - w.m : 5); if (!(dt > 0) || !(w.hz > 0)) return;
    runMin += dt;
    while (j + 1 < sx.length && Math.abs(sx[j + 1].m - w.m) <= Math.abs(sx[j].m - w.m)) j++;
    var s = sx[j]; if (s && s.pw !== null && Math.abs(s.m - w.m) <= 6) { runKwh += s.pw * dt / 60000; runCov += dt; }
  });
  function r2(x) { return x === null ? null : Math.round(x * 100) / 100; }
  return { ok: true, n: sx.length, spanMin: Math.round(spanMin), covPct: Math.round(100 * covMin / Math.max(1, spanMin)),
    kwh: r2(covMin / Math.max(1, spanMin) >= 0.8 || kY === null ? kInt : kY), kwhInt: r2(kInt), kwhYield: r2(kY), peakW: pk, avgW: spanMin ? Math.round(kInt * 60000 / spanMin) : 0, onMin: Math.round(onMin),
    runMin: Math.round(runMin), runKwh: runCov ? r2(runKwh) : null, runCovPct: runMin ? Math.round(100 * runCov / runMin) : 0, errs: errs };
}
function handle_(p, body) {
  if (!prop_('APP_TOKEN') || p.t !== prop_('APP_TOKEN')) return out_({ ok: false, error: 'token' });
  try {
    switch (p.a) {
      case 'ping': return out_({ ok: true, version: VERSION, rev: '2026-10-23',  sxEvery: sxEveryMin_(), wlRun: wlEvery_('WL_RUN', WL_RUN_MIN), wlIdle: wlEvery_('WL_IDLE', WL_IDLE_MIN), now: nowStr_(), wl: !!prop_('WL_USER'), sx: !!prop_('SX_ID'), errWl: prop_('ERR_WL') || null, errSx: prop_('ERR_SX') || null, bf: bfState_(), data: dataStats_() });
      case 'bf': return out_({ ok: true, bf: bfState_() });
      case 'bfskip': { var bd = prop_('BF_DAY'); if (bd) skipDay_(bd); ['ERR_BF', 'BF_FAILN'].forEach(function (k) { P_().deleteProperty(k); }); return out_({ ok: true, bf: bfState_() }); }
      case 'bfrun': { var br = backfillRun_(30000, parseInt(p.n, 10) || 12); return out_({ ok: true, done: br.done, calls: br.calls, msg: br.msg, bf: br.state }); }
      case 'sxfill': { var dd = String(p.d || '').slice(0, 10), added = 0; try { added = sxEnsureDay_(dd); } catch (e0) { return out_({ ok: false, error: e0.message }); } return out_({ ok: true, d: dd, added: added }); }
      case 'live': {
        var lock = LockService.getScriptLock(), wl = null, sx = null;
        try { wl = JSON.parse(prop_('WL_LAST') || 'null'); } catch (e) { wl = null; }
        try { sx = JSON.parse(prop_('SX_LAST') || 'null'); } catch (e) { sx = null; }
        if (p.fresh === '1' && p.src !== 'sx' && prop_('WL_USER') && (!wl || minDiff_(wl.ts, nowStr_()) >= 1) && lock.tryLock(15000)) {
          try { wl = pollWl_(); } catch (e) { errLog_('WL', e.message); } finally { lock.releaseLock(); }
        }
        if (p.fresh === '1' && p.src !== 'wl' && prop_('SX_ID') && (!sx || minDiff_(sx.ts, nowStr_()) >= 3) && lock.tryLock(15000)) {
          try { sx = pollSx_(); P_().deleteProperty('ERR_SX'); } catch (e) { errLog_('SX', e.message); } finally { lock.releaseLock(); }
        }
        return out_({ ok: true, now: nowStr_(), wl: wl, sx: sx, errWl: prop_('ERR_WL') || null, errSx: prop_('ERR_SX') || null, oRaw: prop_('WL_O_RAW') || null });
      }
      case 'near': {                                                     // próbki najbliższe podanej chwili (do podstawiania przy wpisie)
        var ts = String(p.ts || '').slice(0, 16), tol = parseFloat(p.tol) || 10, rw = nearest_(wlSheet_(), ts, tol), rs = nearest_(sxSheet_(), ts, tol), sxo = null;
        if ((!rs || rs.dt > 3) && prop_('SX_ID') && ts.slice(0, 10) <= nowStr_().slice(0, 10)) { try { sxEnsureDay_(ts.slice(0, 10)); var rs2 = nearest_(sxSheet_(), ts, tol); if (rs2) rs = rs2; } catch (e1) { errLog_('SXFILL', e1.message); } }   // brak próbki z tej chwili → dociągnij dzień z historii SolaX
        if (rs) {
          var sh = sxSheet_(), prevYt = null;
          if (rs.idx > 2) { var pr = sh.getRange(rs.idx - 1, 1, 1, 3).getValues()[0]; if (String(pr[0]).slice(0, 10) === ts.slice(0, 10)) prevYt = num_(pr[2]); }
          var pw = num_(rs.row[1]), yt = num_(rs.row[2]);
          sxo = { ts: String(rs.row[0]), dt: rs.dt, pw: pw, yt: yt, on: sxOn_(pw, yt, prevYt) };
        }
        return out_({ ok: true, ts: ts, wl: rw ? Object.assign(wlObj_(rw.row), { dt: rw.dt }) : null, sx: sxo });
      }
      case 'wl': {
        var rows = range_(wlSheet_(), String(p.from), String(p.to), 4000), step = parseInt(p.step, 10) || 0, outR = [], lastB = null;
        rows.forEach(function (r) { var b = step ? Math.floor(Date.parse(String(r[0]) + ':00Z') / 60000 / step) : null; if (!step || b !== lastB) { outR.push(wlObj_(r)); lastB = b; } });
        return out_({ ok: true, rows: outR });
      }
      case 'sx': return out_({ ok: true, rows: sxRows_(range_(sxSheet_(), String(p.from), String(p.to), 20000)) });
      case 'days': {
        var all = daySheet_(), d = all.getLastRow() < 2 ? [] : all.getRange(2, 1, all.getLastRow() - 1, DAY_HDR.length).getValues(), from = String(p.from || '0000'), to = String(p.to || '9999') + '~';
        return out_({ ok: true, rows: d.filter(function (r) { var k = dk_(r[0]); return k >= from && k <= to; }).map(function (r) { return { d: dk_(r[0]), y: num_(r[1]), pk: num_(r[2]), h0: hm_(r[3]), h1: hm_(r[4]), n: +r[5] || 0, im: num_(r[6]), ex: num_(r[7]), ld: num_(r[8]) }; }) });
      }
      case 'sxnow': { var lk2 = LockService.getScriptLock(), sn = null; if (!prop_('SX_ID')) return out_({ ok: false, error: 'SolaX nieskonfigurowany' }); if (!lk2.tryLock(20000)) return out_({ ok: false, error: 'hub zajęty' }); try { sn = pollSx_(); P_().deleteProperty('ERR_SX'); } catch (e2) { errLog_('SX', e2.message); return out_({ ok: false, error: e2.message }); } finally { lk2.releaseLock(); } return out_({ ok: true, sx: sn }); }
      case 'dayfix': {                                                   // napraw zawyżone sumy dzienne (przeniesiony odczyt z wczoraj) na podstawie próbek arkusza SX
        var nD = Math.min(parseInt(p.n, 10) || 30, 120), tdy = nowStr_().slice(0, 10), d0 = addDays_(tdy, -nD), by = {}, fixed = [];
        range_(sxSheet_(), d0 + 'T00:00', tdy + 'T23:59', 40000).forEach(function (r) { var dd = tsk_(r[0]).slice(0, 10); (by[dd] = by[dd] || []).push(r); });
        withLock_(25000, function () {
          var dsh = daySheet_(), dn = dsh.getLastRow() - 1; if (dn < 1) return;
          dsh.getRange(2, 1, dn, DAY_HDR.length).getValues().forEach(function (r, ix) {
            var k = dk_(r[0]), rr = by[k]; if (!rr || rr.length < 6) return;
            var mx = 0; rr.forEach(function (x) { var v = num_(x[2]); if (v !== null && v > mx) mx = v; });
            var y = yieldOfRows_(rr), cur = num_(r[1]) || 0;
            if (mx - y > 0.3 && cur > y + 0.3) { y = Math.round(y * 100) / 100; fixed.push({ d: k, was: cur, now: y }); dsh.getRange(ix + 2, 2).setValue(y); }       // tylko gdy w próbkach widać reset licznika (stary odczyt przed nim)
          });
        });
        return out_({ ok: true, checked: Object.keys(by).length, fixed: fixed });
      }
      case 'pvwin': return out_(pvWindow_(String(p.from || '').slice(0, 16), String(p.to || '').slice(0, 16)));
      case 'push': {                                                     // opcjonalnie: zewnętrzny zbieracz (np. skrypt na komputerze) dopisuje próbki
        var n = 0; body = body || {};
        (body.wl || []).forEach(function (x) { var row = [String(x.ts).slice(0, 16)]; WL_COLS.forEach(function (c) { row.push(x.v && x.v[c] !== undefined ? wnum_(x.v[c]) : ''); }); wlSheet_().appendRow(row); n++; });
        (body.sx || []).forEach(function (x) { sxSheet_().appendRow([String(x.ts).slice(0, 16), x.pw === undefined ? '' : x.pw, x.yt === undefined ? '' : x.yt, x.tot === undefined ? '' : x.tot]); updateDay_(String(x.ts).slice(0, 16), num_(x.pw), num_(x.yt), sxOn_(num_(x.pw), num_(x.yt), null)); n++; });
        return out_({ ok: true, added: n });
      }
      default: return out_({ ok: false, error: 'nieznana akcja' });
    }
  } catch (err) { return out_({ ok: false, error: String(err && err.message || err) }); }
}
function doGet(e) { return handle_(e.parameter || {}, null); }
function doPost(e) {
  var b = {}; try { b = JSON.parse(e.postData.contents); } catch (x) { b = {}; }
  return handle_(Object.assign({}, e.parameter || {}, b.params || {}), b);
}

/* ---------- menu w Arkuszu: konfiguracja i testy ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Monitor PC+PV')
    .addItem('1. Konfiguracja (login, hasło, kody)', 'konfiguracja')
    .addItem('2. Test połączeń', 'testPolaczen')
    .addItem('3. Pokaż link do aplikacji', 'pokazLink')
    .addItem('4. Pobierz historię SolaX (jak w aplikacji mobilnej)', 'startBackfill')
    .addSeparator().addItem('Zainstaluj harmonogram (co 5 min)', 'instalujHarmonogram')
    .addItem('5. Napraw daty w arkuszach (gdy w aplikacji brakuje dni)', 'naprawDaty')
    .addItem('Odczytaj teraz (Warmlink + SolaX)', 'pollNow')
    .addItem('6. Diagnostyka Warmlink (O1 / O10)', 'diagnostykaWarmlink').addToUi();
}
function ask_(ui, title, text, key, keepIfEmpty) {
  var cur = prop_(key), r = ui.prompt(title, text + (cur && keepIfEmpty ? '\n(zostaw puste = bez zmian)' : ''), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return false;
  var v = r.getResponseText().trim();
  if (v) P_().setProperty(key, v); else if (!keepIfEmpty) P_().deleteProperty(key);
  return true;
}
function konfiguracja() {
  var ui = SpreadsheetApp.getUi();
  P_().setProperty('SHEET_ID', SpreadsheetApp.getActive().getId());
  if (!prop_('APP_TOKEN')) P_().setProperty('APP_TOKEN', Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8));
  var steps = [
    ['Warmlink — login', 'Login (e-mail) DRUGIEGO konta Warmlink (z udostępnionym domem). Użyj osobnego konta — jedno konto = jedno logowanie.', 'WL_USER'],
    ['Warmlink — hasło', 'Hasło tego konta.', 'WL_PASS'],
    ['Warmlink — kod pompy', 'Pełny kod urządzenia (z ekranu informacji o urządzeniu w aplikacji).', 'WL_DEVICE'],
    ['SolaX — Client ID', 'Client ID aplikacji z developer.solaxcloud.com (puste = pomiń SolaX).', 'SX_ID'],
    ['SolaX — Client Secret', 'Client Secret tej aplikacji.', 'SX_SECRET'],
    ['SolaX — numer falownika', 'Numer seryjny falownika (deviceSn) — opcjonalnie, puste = pierwszy z konta.', 'SX_SN'],
    ['Warmlink — co ile minut podczas pracy', 'Co ile minut zapisywać odczyt, gdy sprężarka pracuje stabilnie (zalecane 10). W pierwszych 30 min, przy odszranianiu, skokach częstotliwości i mrozie ≤ 7 °C hub sam skraca do 5 min. Puste = 10.', 'WL_RUN'],
    ['Warmlink — sprawdzanie przy postoju', 'Co ile minut sprawdzać, czy pompa ruszyła (nic nie zapisuje, zalecane 10). Przy postoju zapisywany jest tylko jeden wiersz na godzinę. Puste = 10.', 'WL_IDLE'],
    ['SolaX — co ile minut', 'Jak często zapisywać odczyt na żywo z falownika (zalecane 60). Dokładna krzywa 5-minutowa i tak jest dociągana z historii SolaX przy wpisywaniu cyklu i nocą. Puste = 60.', 'SX_EVERY']
  ];
  for (var i = 0; i < steps.length; i++) { if (!ask_(ui, steps[i][0], steps[i][1], steps[i][2], true)) return; }
  ['WL_TOKEN', 'SX_TOK', 'SX_EXP', 'SX_PLANT', 'SX_DEVSN', 'SX_BT', 'SX_CREATED', 'ERR_BF'].forEach(function (k) { P_().deleteProperty(k); });
  try { migrate_(); } catch (e) {}
  instalujHarmonogram();
  ui.alert('Gotowe', 'Zapisano ustawienia i włączono odczyt co 5 minut.\nTeraz: menu „2. Test połączeń”. Potem wdróż jako Aplikację internetową (patrz INSTRUKCJA_HUB.md) i użyj „3. Pokaż link do aplikacji”.', ui.ButtonSet.OK);
}
function pollNow() { poll(true); }                    // menu: odczytaj od razu Warmlink i SolaX, bez czekania na interwał
function instalujHarmonogram() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('poll').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('prune').timeBased().atHour(3).everyDays(1).create();
  if (prop_('SX_ID') && prop_('BF_STAT') && prop_('BF_DONE') !== '1') ScriptApp.newTrigger('backfillTick').timeBased().everyMinutes(10).create();   // dokończ pobieranie historii
}
function pokazLink() {
  var ui = SpreadsheetApp.getUi(), url = null;
  try { url = ScriptApp.getService().getUrl(); } catch (e) { url = null; }
  if (!prop_('APP_TOKEN')) { ui.alert('Najpierw uruchom „1. Konfiguracja”.'); return; }
  if (!url) { ui.alert('Brak wdrożenia', 'Najpierw: Wdróż → Nowe wdrożenie → Aplikacja internetowa (Wykonaj jako: Ja, Kto ma dostęp: Każdy). Potem uruchom to ponownie.', ui.ButtonSet.OK); return; }
  var full = url + '?t=' + prop_('APP_TOKEN');
  var html = HtmlService.createHtmlOutput('<div style="font:14px Arial;padding:6px"><p>Skopiuj <b>cały</b> adres (przycisk poniżej) i wklej w aplikacji: zakładka „📡 Na żywo” → „Adres huba”.</p>' +
    '<textarea id="u" readonly style="width:100%;height:120px;font:12px monospace" onclick="this.select()">' + full + '</textarea>' +
    '<p><button onclick="var t=document.getElementById(\'u\');t.select();document.execCommand(\'copy\');this.textContent=\'Skopiowano ✓\'">Kopiuj adres</button></p>' +
    '<p style="color:#a00;font-size:12px">Adres zawiera klucz dostępu — nie udostępniaj go. Długość: ' + full.length + ' znaków; adres musi kończyć się na /exec?t=… i mieć ok. 100+ znaków przed „?t=”.</p></div>').setWidth(600).setHeight(330);
  ui.showModalDialog(html, 'Link do aplikacji');
}

function diagnostykaWarmlink() {                 // pokazuje, co dokładnie Warmlink zwraca dla kodów O* (sprężarka, zawór CO/CWU)
  var ui = SpreadsheetApp.getUi(), code = prop_('WL_DEVICE'), wide = ['Power', 'Mode', 'T01', 'T02', 'T04', 'T30'], i;
  if (!code) { ui.alert('Brak kodu urządzenia Warmlink (menu: 1. Konfiguracja).'); return; }
  for (i = 1; i <= 16; i++) { wide.push('O' + i); wide.push('O' + ('0' + i).slice(-2)); }
  wide = wide.filter(function (c, ix) { return wide.indexOf(c) === ix; });
  var tok = prop_('WL_TOKEN') || wlLogin_(), res = wlCall_('device/getDataByCode', { appId: WL_APP_ID, deviceCode: code, protocalCodes: wide }, tok), note = '';
  if (!wlOk_(res)) { tok = wlLogin_(); res = wlCall_('device/getDataByCode', { appId: WL_APP_ID, deviceCode: code, protocalCodes: wide }, tok); }
  if (!wlOk_(res)) { note = 'Szeroka lista została odrzucona (' + JSON.stringify(res).slice(0, 160) + ') — próbuję tylko O1, O01, O10.\n\n'; res = wlCall_('device/getDataByCode', { appId: WL_APP_ID, deviceCode: code, protocalCodes: ['O1', 'O01', 'O10'] }, tok); }
  if (!wlOk_(res)) { ui.alert('Diagnostyka Warmlink', note + 'Warmlink nie zwrócił danych: ' + JSON.stringify(res).slice(0, 300), ui.ButtonSet.OK); return; }
  var rows = res.objectResult || [], oRows = rows.filter(function (r) { return r && /^[oO]\d/.test(String(r.code)); }), L = [];
  L.push('Warmlink zwrócił ' + rows.length + ' kodów, w tym ' + oRows.length + ' kodów O*.');
  if (oRows.length) oRows.forEach(function (r) { L.push(JSON.stringify(r).slice(0, 180)); });
  else L.push('Brak O1/O01/O10 w odpowiedzi — Warmlink nie zwraca ich przez ten endpoint dla tego urządzenia.\nKody, które wróciły: ' + rows.map(function (r) { return r && r.code; }).join(', ').slice(0, 400));
  ui.alert('Diagnostyka Warmlink (O1 / O10)', note + L.join('\n'), ui.ButtonSet.OK);
}
function testPolaczen() {
  var ui = SpreadsheetApp.getUi(), L = [];
  try { var v = wlValues_(); L.push('✔ Warmlink: OK — sprężarka T30 = ' + v.T30 + ' Hz, T01 = ' + v.T01 + ', T02 = ' + v.T02 + ', T04 = ' + v.T04); }
  catch (e) { L.push('✘ Warmlink: ' + e.message + '\n   (Jeśli to „Address unavailable”/certyfikat — Google nie połączył się z portem 449; zob. INSTRUKCJA_HUB.md, sekcja „Gdy Warmlink nie działa”.)'); }
  if (prop_('SX_ID')) {
    try { var s = sxValues_(); L.push('✔ SolaX: OK — dziś ' + s.yt + ' kWh, suma ' + s.tot + ' kWh, moc ' + (s.pw === null ? 'brak w odpowiedzi (stan PV wyliczany z przyrostu produkcji)' : s.pw + ' W')); }
    catch (e2) { L.push('✘ SolaX: ' + e2.message); }
    try { var yd = addDays_(nowStr_().slice(0, 10), -1), hh = sxHistDay_(yd); L.push((hh.length ? '✔' : '✘') + ' SolaX historia: wczoraj (' + yd + ') ' + hh.length + ' próbek co 5 min' + (hh.length ? ', szczyt ' + Math.max.apply(null, hh.map(function (r) { return r[1] || 0; })) + ' W' : ' — brak danych (sprawdź numer falownika)')); }
    catch (e3) { L.push('✘ SolaX historia: ' + e3.message); }
  } else L.push('— SolaX: nie skonfigurowano');
  L.push('Wyzwalacze: ' + ScriptApp.getProjectTriggers().length + (ScriptApp.getProjectTriggers().length ? ' (OK)' : ' — uruchom „Zainstaluj harmonogram”'));
  ui.alert('Test połączeń', L.join('\n\n'), ui.ButtonSet.OK);
}
