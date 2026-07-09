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
  'PSA','SHBAkzeptiert','Signatur','Eingangszeit','Abgangszeit','Status','Bemerkungen','GruppenId'];
const COLSET = new Set(COLS);

const posts = [];           // erfasste POST-Bodies (fields)
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
    for (const k of Object.keys(fields)) { if (k.includes('@')) continue;
      if (!COLSET.has(k)) { sentUnknownColumn = true; return resp({ error:{ message:'bad' } }, 400); } }
    return resp({ id: 'patched' });
  }
  if (u.includes('/me?$select')) return resp({ userPrincipalName:'administrator@dihag.com', mail:'administrator@dihag.com', otherMails:[], proxyAddresses:[] });
  if (u.includes('dihag.sharepoint.com:/sites/IT') && !u.includes('/lists')) return resp({ id:'siteid' });
  if (u.includes('/lists/Besucheranmeldung')) return resp({ id:'listid', displayName:'Besucheranmeldung' });
  if (u.includes('/lists/listid/columns')) return resp({ value: COLS.map(n => ({ name:n, displayName:n })) });
  if (u.includes('/lists?$select=id,displayName')) return resp({ value: [] });  // keine Config-Liste
  if (u.includes('/lists/listid/items')) return resp({ value: [] });            // items GET
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
  w.HTMLCanvasElement.prototype.getContext = () => ({
    lineWidth:0, lineCap:'', strokeStyle:'', beginPath(){}, moveTo(){}, lineTo(){}, stroke(){}, clearRect(){}
  });
  w.HTMLCanvasElement.prototype.toDataURL = () => 'data:image/png;base64,iVBORw0KGgo=';

  // app.js in der Fensterrealität ausführen + Test-Handles exportieren
  let src = readFileSync(join(root, 'app.js'), 'utf8');
  src += `\n;window.__app = { submitNew, navigate, applyTemplate, sendSavedInvite, isOverdue, get HAVE(){return HAVE}, get C(){return C}, WERKE, get account(){return account}, isFull, canSeeDashboard, canSeeReports, isMine };`;
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
