// ── Smoke-Test ───────────────────────────────────────────────────────────────
// Bootet die App unter jsdom mit gefälschtem MSAL + Microsoft Graph und fährt
// eine echte „Neue Anmeldung" durch submitNew(). Der Mock-SharePoint kennt
// absichtlich NICHT alle Spalten (Funktion, BesucherTelefon fehlen). Sendet die
// App eine nicht existierende Spalte, antwortet der Mock – wie echtes SharePoint –
// mit 400, und der Test schlägt fehl. So wird der 400-Fehler dauerhaft verhindert.
//
// Start:  npm run smoke   (nach: npm install)

import { JSDOM } from 'jsdom';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, '..');

let failures = 0;
const ok  = (cond, msg) => { if (cond) console.log('  ✓ ' + msg); else { console.error('  ✗ ' + msg); failures++; } };
const sleep = ms => new Promise(r => setTimeout(r, ms));

// Vom Mock-SharePoint bekannte Spalten (INTERN). Bewusst OHNE Funktion/BesucherTelefon.
const COLS = ['Title','Werk','Bereich','AnsprechpartnerName','AnsprechpartnerTelefon',
  'Besuchsdatum','Ankunftszeit','Firma','BesucherEmail','Autokennzeichen','Besuchszweck',
  'PSA','SHBAkzeptiert','Signatur','Eingangszeit','Abgangszeit','Status','Bemerkungen','GruppenId','ErstellerUPN'];
const COLSET = new Set(COLS);

const posts = [];           // erfasste POST-Bodies (fields)
const patches = [];         // erfasste PATCH-Bodies (fields)
const sentMail = [];        // erfasste /me/sendMail-Bodies
let sentUnknownColumn = false;

function resp(data, status = 200) {
  const good = status >= 200 && status < 300;
  return { ok: good, status, statusText: good ? 'OK' : 'Error',
    json: async () => data, text: async () => JSON.stringify(data) };
}

async function fakeFetch(url, opts) {
  const method = (opts && opts.method) || 'GET';
  const u = String(url);

  if (method === 'POST' && u.includes('/me/sendMail')) { sentMail.push(JSON.parse(opts.body || '{}')); return resp({}, 202); }
  if (method === 'POST' && u.includes('/lists/listid/items')) {
    const body = JSON.parse(opts.body || '{}');
    const fields = body.fields || {};
    posts.push(fields);
    for (const k of Object.keys(fields)) {
      if (k.includes('@')) continue;                 // @odata.type-Annotation, keine Spalte
      if (!COLSET.has(k)) { sentUnknownColumn = true;
        return resp({ error: { code: 'invalidRequest', message: `Field '${k}' does not exist` } }, 400); }
    }
    return resp({ id: 'new-' + posts.length, fields });
  }
  if (method === 'PATCH') {
    const fields = JSON.parse(opts.body || '{}');
    patches.push(fields);
    for (const k of Object.keys(fields)) { if (k.includes('@')) continue;
      if (!COLSET.has(k)) { sentUnknownColumn = true; return resp({ error:{ message:'bad' } }, 400); } }
    return resp({ id: 'patched' });
  }
  if (u.includes('/me?$select')) return resp({ userPrincipalName:'administrator@dihag.com', mail:'administrator@dihag.com', otherMails:[], proxyAddresses:[] });
  if (u.includes('dihag.sharepoint.com:/sites/IT') && !u.includes('/lists')) return resp({ id:'siteid' });
  if (u.includes('/lists/Besucheranmeldung')) return resp({ id:'listid', displayName:'Besucheranmeldung' });
  if (u.includes('/lists/listid/columns')) return resp({ value: COLS.map(n => ({ name:n, displayName:n })) });
  if (u.includes('/lists?$select=id,displayName')) return resp({ value: [] });  // keine Config-Liste
  // items GET über zwei Seiten (nextLink) → testet Pagination.
  // Seite 1: rec1 (ohne SHB, für „SHB nachträglich"); Seite 2: rec2 (eingecheckt).
  if (u.includes('/lists/listid/items')) {
    const nowIso = new Date().toISOString();
    if (u.includes('nextpage')) return resp({ value: [
      { id:'rec2', createdDateTime:nowIso, createdBy:{ user:{ email:'administrator@dihag.com', displayName:'Administrator' } },
        fields:{ Title:'Hans Vorort', Werk:'DSO', Bereich:'Tor 1', Firma:'Gamma AG', Status:'Eingecheckt', SHBAkzeptiert:true,
          Besuchsdatum:nowIso, Eingangszeit:nowIso } }
    ] });
    return resp({ value: [
      { id:'rec1', createdDateTime:nowIso, createdBy:{ user:{ email:'administrator@dihag.com', displayName:'Administrator' } },
        fields:{ Title:'Erika Ohne SHB', Werk:'SHB', Bereich:'Tor 2', Firma:'Beta AG', Status:'Angemeldet', SHBAkzeptiert:false } }
    ], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/sites/siteid/lists/listid/items?nextpage=1' });
  }
  return resp({});
}

// MSAL-Stub: sofort als Admin fedorov@dihag.com angemeldet
class FakePCA {
  constructor(){}
  async initialize(){}
  async handleRedirectPromise(){ return null; }
  getAllAccounts(){ return [{ username:'administrator@dihag.com', name:'Administrator', idTokenClaims:{} }]; }
  async acquireTokenSilent(){ return { accessToken:'tkn' }; }
  async acquireTokenRedirect(){}
  async loginRedirect(){}
  logoutRedirect(){}
}

async function main() {
  console.log('Smoke-Test: DIHAG Besuchermanagement');

  const html = readFileSync(join(root, 'index.html'), 'utf8');
  const dom = new JSDOM(html, { url: 'http://localhost:8767/', runScripts: 'outside-only' });
  const w = dom.window;

  // Stubs injizieren
  w.msal = { PublicClientApplication: FakePCA };
  w.fetch = fakeFetch;
  w.event = undefined;
  w.confirm = () => true;
  w.print = () => {};
  w.HTMLCanvasElement.prototype.getContext = () => ({
    lineWidth:0, lineCap:'', strokeStyle:'', beginPath(){}, moveTo(){}, lineTo(){}, stroke(){}, clearRect(){}
  });
  w.HTMLCanvasElement.prototype.toDataURL = () => 'data:image/png;base64,iVBORw0KGgo=';

  // app.js in der Fensterrealität ausführen + Test-Handles exportieren
  let src = readFileSync(join(root, 'app.js'), 'utf8');
  src += `\n;window.__app = { submitNew, navigate, applyTemplate, sendSavedInvite, isOverdue, buildCsv, canEditItem, findDuplicate, setNewMode, openSHBModal, saveSHB, shbActive, inDateScope, printAttendance, dsgvoFind, openSettings, sortList, setKioskTimeout, get itemCount(){return ITEMS.length}, get kioskMin(){return Number(appSettings.kioskTimeoutMin)||0}, get HAVE(){return HAVE}, get C(){return C}, WERKE, get account(){return account}, isFull, canSeeDashboard, canSeeReports, isMine };`;
  w.eval(src);

  // Auf Boot warten (discoverSP füllt HAVE)
  for (let i = 0; i < 60 && !(w.__app && w.__app.HAVE && w.__app.HAVE.size); i++) await sleep(50);
  const bootErr = w.document.getElementById('boot-err')?.textContent || '';
  ok(w.document.getElementById('app').style.display === '', 'App gestartet (Boot erfolgreich)' + (bootErr ? ' – boot-err: '+bootErr : ''));
  ok(w.__app.account && w.__app.account.username === 'administrator@dihag.com', 'Als Admin administrator@dihag.com angemeldet');

  // Werke-Liste
  const WERKE = w.__app.WERKE;
  ['EMH','Kein','Kernwerk','OZB'].forEach(x => ok(!WERKE.includes(x), `Werk „${x}" entfernt`));
  ok(WERKE.length === 10 && WERKE.includes('SHB'), 'Werke-Liste hat 10 Einträge inkl. SHB');

  // Neue Anmeldung ausfüllen
  w.__app.navigate('new');
  await sleep(10);
  const doc = w.document;
  // Modus-Umschaltung (Voranmeldung vs. vor Ort)
  w.__app.setNewMode('voranmeldung');
  ok(doc.getElementById('ankunft-group').style.display !== 'none', 'Voranmeldung: Ankunftszeit sichtbar');
  ok(doc.getElementById('shb-section').style.display === 'none', 'Voranmeldung: SHB ausgeblendet');
  w.__app.setNewMode('vorort');
  ok(doc.getElementById('ankunft-group').style.display === 'none', 'Vor Ort: Ankunftszeit ausgeblendet');
  ok(doc.getElementById('shb-section').style.display !== 'none', 'Vor Ort: SHB sichtbar');
  const set = (id, v) => { const el = doc.getElementById(id); el.value = v; };
  set('f-werk', 'SHB'); set('f-bereich', 'Halle 3'); set('f-firma', 'ACME GmbH');
  doc.querySelector('#visitors [data-f="name"]').value = 'Max Mustermann';
  doc.querySelector('#visitors [data-f="funktion"]').value = 'Einkauf';       // Spalte fehlt → darf NICHT gesendet werden
  doc.querySelector('#visitors [data-f="tel"]').value = '0123';               // Spalte fehlt → darf NICHT gesendet werden
  doc.querySelector('#visitors [data-f="email"]').value = 'max@acme.example'; // Spalte existiert → wird gesendet
  doc.querySelector('input[name="zweck"][value="Audit"]').checked = true;     // Multi-Choice
  doc.querySelector('#visitors [data-psa][value="Schutzhelm"]').checked = true; // Multi-Choice
  doc.getElementById('f-shb').checked = true;

  // Unterschrift simulieren (setzt sigHasInk = true)
  const sig = doc.getElementById('sig');
  sig.dispatchEvent(new w.MouseEvent('mousedown', { clientX:5, clientY:5 }));
  sig.dispatchEvent(new w.MouseEvent('mousemove', { clientX:9, clientY:9 }));
  w.dispatchEvent(new w.MouseEvent('mouseup', {}));

  await w.__app.submitNew();
  await sleep(10);

  // Auswertung
  ok(posts.length === 1, 'Genau ein Datensatz gepostet');
  ok(!sentUnknownColumn, 'KEINE nicht existierende Spalte gesendet (kein 400)');
  const f = posts[0] || {};
  ok(f.Title === 'Max Mustermann', 'Title = Besuchername gesetzt');
  ok(f.Werk === 'SHB', 'Werk gesetzt');
  ok(f.Status === 'Angemeldet', 'Status = Angemeldet');
  ok(f.Firma === 'ACME GmbH', 'Firma gesetzt');
  ok(f.BesucherEmail === 'max@acme.example', 'Vorhandene Spalte BesucherEmail gesendet');
  ok(!('Funktion' in f), 'Fehlende Spalte Funktion NICHT gesendet');
  ok(!('BesucherTelefon' in f), 'Fehlende Spalte BesucherTelefon NICHT gesendet');
  ok(f.ErstellerUPN === 'administrator@dihag.com', 'ErstellerUPN beim Anlegen gesetzt');
  ok(w.__app.isMine({ creatorUPN:'administrator@dihag.com' }) === true, 'isMine über ErstellerUPN');
  ok(f.SHBAkzeptiert === true, 'Vor-Ort: SHB akzeptiert gesetzt');
  ok(typeof f.Signatur === 'string' && f.Signatur.startsWith('data:image'), 'Vor-Ort: Unterschrift gespeichert');
  ok(Object.keys(f).every(k => k.includes('@') || COLSET.has(k)), 'Alle gesendeten Felder existieren als Spalte');
  ok(Array.isArray(f.Besuchszweck) && f.Besuchszweck.join(',') === 'Audit', 'Besuchszweck als Array gesendet');
  ok(f['Besuchszweck@odata.type'] === 'Collection(Edm.String)', 'Multi-Choice Besuchszweck mit @odata.type annotiert');
  ok(f['PSA@odata.type'] === 'Collection(Edm.String)', 'Multi-Choice PSA mit @odata.type annotiert');

  // Einladung per Microsoft Graph (kein Outlook/mailto)
  await w.__app.sendSavedInvite(0);
  await sleep(5);
  ok(sentMail.length === 1, 'Einladung per Graph /me/sendMail gesendet');
  ok(sentMail[0]?.message?.toRecipients?.[0]?.emailAddress?.address === 'max@acme.example', 'Einladung an richtige Adresse');
  ok((sentMail[0]?.message?.subject||'').includes('Einladung'), 'Einladungs-Betreff gesetzt');

  // Rollen & Sichtbarkeit (Admin ist vollberechtigt)
  ok(w.__app.isFull(), 'Admin ist vollberechtigt');
  ok(w.__app.canSeeDashboard(), 'Dashboard für Vollberechtigte/Admin sichtbar');
  ok(w.__app.canSeeReports(), 'Reports für Vollberechtigte/Admin sichtbar');
  ok(doc.querySelector('.nav-item[data-view="dashboard"]').style.display !== 'none', 'Dashboard-Nav sichtbar (Admin)');
  ok(doc.querySelector('.nav-item[data-view="reports"]').style.display !== 'none', 'Reports-Nav sichtbar (Admin)');
  ok(doc.querySelector('.nav-item[data-view="records"]').textContent.includes('Eigene Datensätze'), 'Nav „Eigene Datensätze" umbenannt');
  ok(w.__app.isMine({ createdByEmail:'administrator@dihag.com' }) === true, 'Eigener Datensatz (E-Mail) erkannt');
  ok(w.__app.isMine({ createdByEmail:'fremd@dihag.com', createdBy:'Jemand' }) === false, 'Fremder Datensatz nicht als eigener erkannt');

  // Anwesenheits-Warnung (noch eingecheckt > 8 h)
  const hAgo = h => new Date(Date.now() - h*3600000).toISOString();
  ok(w.__app.isOverdue({ status:'Eingecheckt', eingang:hAgo(10) }) === true, 'Überfällig: >8 h anwesend erkannt');
  ok(w.__app.isOverdue({ status:'Eingecheckt', eingang:hAgo(1) }) === false, 'Nicht überfällig: 1 h anwesend');
  ok(w.__app.isOverdue({ status:'Geschlossen', eingang:hAgo(10) }) === false, 'Geschlossen ist nicht überfällig');

  // Weitere Optimierungen: CSV, Bearbeiten-Recht, Dubletten
  const csv = w.__app.buildCsv([{ werk:'SHB', bereich:'Halle;3', besucherName:'Max "M"', firma:'ACME', funktion:'', tel:'', email:'', kennzeichen:'', besuchsdatum:'2026-07-09', eingang:'', abgang:'', status:'Angemeldet', ansprechName:'Chef', zweck:['Audit'], psa:['Warnweste'] }]);
  ok(csv.split('\r\n')[0].startsWith('﻿Werk;Bereich;Besucher'), 'CSV: Kopfzeile mit BOM');
  ok(csv.includes('"Halle;3"') && csv.includes('"Max ""M"""'), 'CSV: Sonderzeichen korrekt maskiert');
  ok(w.__app.canEditItem({}) === true, 'Admin darf bearbeiten');
  ok(w.__app.findDuplicate('Niemand','Nirgends','2026-07-09') === undefined, 'Dublettenprüfung ohne Treffer');

  // SHB nachträglich (Schritt 2) auf einen Datensatz ohne SHB (rec1)
  w.__app.openSHBModal('rec1');
  await sleep(5);
  doc.getElementById('shb-ok').checked = true;
  const sm = doc.getElementById('sig-modal');
  sm.dispatchEvent(new w.MouseEvent('mousedown', { clientX:2, clientY:2 }));
  sm.dispatchEvent(new w.MouseEvent('mousemove', { clientX:6, clientY:6 }));
  w.dispatchEvent(new w.MouseEvent('mouseup', {}));
  const before = patches.length;
  await w.__app.saveSHB('rec1');
  await sleep(5);
  const lastPatch = patches[patches.length-1] || {};
  ok(patches.length > before && lastPatch.SHBAkzeptiert === true, 'SHB nachträglich: SHBAkzeptiert gesetzt');
  ok(typeof lastPatch.Signatur === 'string' && lastPatch.Signatur.startsWith('data:image'), 'SHB nachträglich: Unterschrift gespeichert');
  ok(w.__app.shbActive() === true, 'SHB standardmäßig aktiv');

  // Datums-Bereich
  const t = new Date().toISOString().slice(0,10);
  ok(w.__app.inDateScope(t, 'today') === true && w.__app.inDateScope('2000-01-01', 'today') === false, 'Datumsfilter „Heute"');
  ok(w.__app.inDateScope('2000-01-01', '') === true, 'Datumsfilter „Alle"');

  // Anwesenheitsliste (Evakuierung) – Druckbereich befüllt
  w.__app.printAttendance();
  ok(doc.getElementById('print-area').innerHTML.includes('Anwesenheitsliste'), 'Anwesenheitsliste erzeugt');

  // Autocomplete-Datalists aus früheren Besuchen (rec1: Erika/Beta AG)
  w.__app.navigate('new');
  await sleep(5);
  ok(!!doc.querySelector('#dl-firma option[value="Beta AG"]'), 'Autocomplete Firma aus Historie');
  ok(!!doc.querySelector('#dl-names option[value="Erika Ohne SHB"]'), 'Autocomplete Name aus Historie');

  // DSGVO-Suche (Admin)
  w.__app.openSettings();
  await sleep(5);
  doc.getElementById('dsgvo-q').value = 'erika';
  w.__app.dsgvoFind();
  ok(doc.getElementById('dsgvo-res').textContent.includes('gefunden'), 'DSGVO: Treffer gefunden');

  // Pagination über zwei Seiten (rec1 + rec2)
  ok(w.__app.itemCount === 2, 'Pagination: zweite Seite (nextLink) geladen');

  // Sortierung
  const sn = w.__app.sortList([{besucherName:'Zoe',besuchsdatum:'2026-01-01'},{besucherName:'Anna',besuchsdatum:'2026-01-02'}], 'name');
  ok(sn[0].besucherName === 'Anna', 'Sortierung nach Name (A–Z)');
  const sd = w.__app.sortList([{besuchsdatum:'2026-01-01'},{besuchsdatum:'2026-02-01'}], 'date-desc');
  ok(sd[0].besuchsdatum === '2026-02-01', 'Sortierung Datum neueste zuerst');

  // Reports mit Zeitraum/Trend
  w.__app.navigate('reports');
  await sleep(5);
  ok(doc.getElementById('reports-body').textContent.includes('Verlauf'), 'Reports: Verlauf/Trend gerendert');
  ok(doc.getElementById('reports-body').textContent.includes('Aufenthalt'), 'Reports: ø-Aufenthalt gerendert');

  // Kiosk-Auto-Logout-Einstellung
  w.__app.setKioskTimeout(10);
  ok(w.__app.kioskMin === 10, 'Kiosk-Timeout gesetzt');
  w.__app.setKioskTimeout(0);
  ok(w.__app.kioskMin === 0, 'Kiosk-Timeout aus');

  // Modal per Esc schließen
  w.__app.openSettings();
  await sleep(5);
  doc.dispatchEvent(new w.KeyboardEvent('keydown', { key:'Escape', bubbles:true }));
  ok(doc.getElementById('settings-modal').classList.contains('hidden'), 'Esc schließt das Modal');

  // Anleitung-Reiter
  w.__app.navigate('anleitung');
  await sleep(5);
  ok(doc.querySelector('.nav-item[data-view="anleitung"]')?.textContent.includes('Anleitung'), 'Nav „Anleitung" vorhanden');
  ok(doc.getElementById('anleitung-body').textContent.includes('Bedienungsanleitung'), 'Anleitung-Inhalt gerendert');

  // „Als Vorlage" – Formular aus einem Datensatz vorbefüllen
  w.__app.navigate('new');
  await sleep(10);
  w.__app.applyTemplate({ werk:'DSO', bereich:'Tor 1', ansprechName:'Chef', ansprechTel:'', firma:'Beta AG',
    zweck:['Audit'], besucherName:'Erika Beispiel', funktion:'', tel:'', email:'', kennzeichen:'', psa:['Warnweste'] });
  ok(doc.getElementById('f-werk').value === 'DSO', 'Vorlage: Werk vorbefüllt');
  ok(doc.getElementById('f-bereich').value === 'Tor 1', 'Vorlage: Bereich vorbefüllt');
  ok(doc.getElementById('f-firma').value === 'Beta AG', 'Vorlage: Firma vorbefüllt');
  ok(doc.querySelector('#visitors [data-f="name"]').value === 'Erika Beispiel', 'Vorlage: Besuchername vorbefüllt');
  ok(doc.querySelector('input[name="zweck"][value="Audit"]').checked, 'Vorlage: Besuchszweck übernommen');
  ok(doc.querySelector('#visitors [data-psa][value="Warnweste"]').checked, 'Vorlage: PSA übernommen');

  await scenarioAccessMatch();

  console.log(failures ? `\nFEHLGESCHLAGEN: ${failures} Prüfung(en)` : '\nALLE PRÜFUNGEN BESTANDEN');
  process.exit(failures ? 1 : 0);
}

// Regression: Nicht-Admin, dessen Konfig-Schlüssel in GROSSschreibung hinterlegt ist,
// muss trotzdem Zugriff bekommen (Kern des „Kein Zugriff trotz hinterlegt"-Bugs).
async function scenarioAccessMatch() {
  const html = readFileSync(join(root, 'index.html'), 'utf8');
  const fetch2 = async (url, opts) => {
    const u = String(url);
    if (u.includes('/me?$select')) return resp({ userPrincipalName:'Tester@DIHAG.com', mail:'Tester@DIHAG.com', otherMails:[], proxyAddresses:[] });
    if (u.includes('dihag.sharepoint.com:/sites/IT') && !u.includes('/lists')) return resp({ id:'siteid' });
    if (u.includes('/lists/Besucheranmeldung')) return resp({ id:'listid', displayName:'Besucheranmeldung' });
    if (u.includes('/lists/listid/columns')) return resp({ value: COLS.map(n => ({ name:n, displayName:n })) });
    if (u.includes('/lists?$select=id,displayName')) return resp({ value: [{ id:'cfg', displayName:'BESU_Konfiguration' }] });
    if (u.includes('/lists/cfg/items')) return resp({ value: [{ id:'a1', fields:{ Title:'access',
      ConfigValue: JSON.stringify({ users: { 'Tester@DIHAG.com': { role:'sekretariat', werke:['SHB'] } } }) } }] });
    if (u.includes('/lists/listid/items')) return resp({ value: [] });
    return resp({});
  };
  class PCA2 extends FakePCA { getAllAccounts(){ return [{ username:'Tester@DIHAG.com', name:'Tester', idTokenClaims:{} }]; } }
  const dom = new JSDOM(html, { url:'http://localhost:8767/', runScripts:'outside-only' });
  const w = dom.window;
  w.msal = { PublicClientApplication: PCA2 }; w.fetch = fetch2; w.event = undefined; w.confirm = () => true;
  w.HTMLCanvasElement.prototype.getContext = () => ({ lineWidth:0,lineCap:'',strokeStyle:'',beginPath(){},moveTo(){},lineTo(){},stroke(){},clearRect(){} });
  w.HTMLCanvasElement.prototype.toDataURL = () => 'data:image/png;base64,AAAA';
  let src = readFileSync(join(root, 'app.js'), 'utf8');
  src += `\n;window.__app2 = { get role(){return myRole()}, isFull, get access(){return myAccess()}, get state(){return accessLoadState} };`;
  w.eval(src);
  for (let i=0; i<60 && !(w.__app2 && w.__app2.access); i++) await sleep(50);
  ok(w.document.getElementById('app').style.display === '', 'Zugriff trotz Groß-/Kleinschreibung im Konfig-Schlüssel');
  ok(w.__app2.role === 'sekretariat', 'Rolle aus Konfig übernommen (sekretariat)');
  ok(w.__app2.isFull(), 'sekretariat ist vollberechtigt');
}

main().catch(e => { console.error('Smoke-Test-Ausnahme:', e); process.exit(1); });
