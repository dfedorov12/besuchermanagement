'use strict';

/* ══════════════════════════════════════════════════════════════════════════
   DIHAG Besuchermanagement
   Statische SPA · MSAL-Login · SharePoint-Liste als Backend (Microsoft Graph)
   Zugriff: Werk + Rolle (zentral in SharePoint gepflegt, nur Admin)
   Datenschutz: keine Dritt-CDNs mit PII, keine PII im localStorage,
                Aufbewahrung via SharePoint/Retention-Job (90 Tage)
   ═════════════════════════════════════════════════════════════════════════ */

// ── CONFIG ──────────────────────────────────────────────────────────────────
const CLIENT_ID = '674a4aed-2a41-4c31-9d3f-ded1a1377afa';
const TENANT_ID = 'fdb70646-023a-403b-a4b9-1f474a935123';
const SCOPES    = ['User.Read', 'Sites.ReadWrite.All'];
// SharePoint-Site mit den Listen (bei abweichendem Namen hier anpassen):
const SP_SITE   = 'dihag.sharepoint.com:/sites/IT';
const SP_LIST   = 'Besucheranmeldung';
const ACCESS_LIST_NAME  = 'BESU_Konfiguration';
const ACCESS_ITEM_TITLE = 'access';
// Admins: sehen alle Werke, dürfen Zugriffsrechte pflegen und löschen.
const ADMIN_EMAILS = ['administrator@dihag.com'];
const API = 'https://graph.microsoft.com/v1.0';

// Werke (Reihenfolge = Anzeige). Steuert den Zugriff.
const WERKE = ['DIHAG','DSO','EIS','EWA','LEG','MEG','SCH','SHB','WGC','ZAI'];
const BESUCHSZWECKE = ['Werksbesichtigung','Kundenbesuch','Audit','Lieferantenbesuch','Sonstiges','DIHAG'];
const PSA_LISTE     = ['Schutzhelm','Schutzbrille','Eigene PSA','Warnweste','Gehörschutz'];
const ROLLEN = {
  verantwortlicher:'SHB-Verantwortlicher (nur eigene Anmeldungen)',
  wachschutz:      'Wachschutz (voll)',
  sekretariat:     'Sekretariat (voll)'
};
// „Vollberechtigt": sieht Dashboard + Reports + alle Datensätze der Werke.
// (Admin ist immer vollberechtigt und darf zusätzlich Zugriffsrechte verwalten.)
const FULL_ROLES = ['wachschutz', 'sekretariat'];

// ── STATE ───────────────────────────────────────────────────────────────────
let msalApp = null, account = null;
let siteId = null, listId = null, accessListId = null, accessConfigItemId = null;
let C = {};                 // logischer Key → interner SP-Spaltenname
let HAVE = new Set();       // logische Keys, für die eine SP-Spalte wirklich existiert
let ITEMS = [];             // normalisierte Datensätze
let accessUsers = {};       // { upn: { role, werke:[] } }
let _meIds = null;
let currentView = 'dashboard';
let newVisitorSeq = 0;
let templateItem = null;   // Datensatz, dessen Werte in „Neue Anmeldung" vorbefüllt werden

// ── DOM-HELFER ──────────────────────────────────────────────────────────────
const $id = id => document.getElementById(id);
function esc(s){ return String(s==null?'':s).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m])); }
function toast(msg, type){
  const c = $id('toast-c'); if(!c) return;
  const t = document.createElement('div');
  t.className = 'toast toast-' + (type||'info');
  t.textContent = msg;
  c.appendChild(t);
  setTimeout(()=>{ t.style.opacity='0'; setTimeout(()=>t.remove(), 300); }, 3600);
}
function bootSub(s){ const e=$id('boot-sub'); if(e) e.textContent=s; }
function bootErr(s){ const e=$id('boot-err'); if(e) e.textContent=s; }

// ── MSAL / LOGIN ────────────────────────────────────────────────────────────
async function initMsal(){
  const redirectUri = location.href.split('?')[0].split('#')[0];
  msalApp = new msal.PublicClientApplication({
    auth: { clientId: CLIENT_ID, authority:`https://login.microsoftonline.com/${TENANT_ID}`, redirectUri },
    cache:{ cacheLocation:'sessionStorage' }   // PII-schonend: kein persistentes Token im localStorage
  });
  await msalApp.initialize();
  const resp = await msalApp.handleRedirectPromise();
  if (resp && resp.account) account = resp.account;
  if (!account){ const accs = msalApp.getAllAccounts(); if (accs.length) account = accs[0]; }
}
async function doLogin(){ try{ await msalApp.loginRedirect({ scopes: SCOPES }); }catch(e){ bootErr(e.message); } }
function doLogout(){ msalApp.logoutRedirect({ account }); }

async function getToken(force){
  try { return (await msalApp.acquireTokenSilent({ scopes:SCOPES, account, forceRefresh:!!force })).accessToken; }
  catch(e){ await msalApp.acquireTokenRedirect({ scopes:SCOPES }); throw e; }
}

// ── GRAPH ───────────────────────────────────────────────────────────────────
async function gFetch(path, opts){
  const token = await getToken();
  const url = path.startsWith('http') ? path : API + path;
  const r = await fetch(url, { ...opts, headers:{ Authorization:'Bearer '+token, 'Content-Type':'application/json', ...(opts&&opts.headers||{}) } });
  if (!r.ok){ const txt = await r.text(); throw new Error(`${r.status} ${r.statusText} – ${txt.slice(0,300)}`); }
  return r.status===204 ? null : r.json();
}
const gGet    = p => gFetch(p);
const gPost   = (p,b) => gFetch(p,{ method:'POST',  body:JSON.stringify(b) });
const gPatch  = (p,b) => gFetch(p,{ method:'PATCH', body:JSON.stringify(b) });
const gDelete = p => gFetch(p,{ method:'DELETE' });

// ── SHAREPOINT DISCOVERY + SPALTEN-MAPPING ─────────────────────────────────
const FIELD_DEFS = [
  { k:'Werk',            cands:['Werk'] },
  { k:'Bereich',         cands:['Bereich'] },
  { k:'AnsprechName',    cands:['AnsprechpartnerName','Ansprechpartner'] },
  { k:'AnsprechTel',     cands:['AnsprechpartnerTelefon','AnsprechTelefon'] },
  { k:'Besuchsdatum',    cands:['Besuchsdatum'] },
  { k:'Ankunftszeit',    cands:['Ankunftszeit'] },
  { k:'BesucherName',    cands:['Title'] },   // Titel = Besuchername (gut lesbare Standardansicht)
  { k:'Firma',           cands:['Firma'] },
  { k:'Funktion',        cands:['Funktion'] },
  { k:'BesucherTelefon', cands:['BesucherTelefon','Telefon'] },
  { k:'BesucherEmail',   cands:['BesucherEmail','EMail','Email'] },
  { k:'Autokennzeichen', cands:['Autokennzeichen','Kennzeichen'] },
  { k:'Besuchszweck',    cands:['Besuchszweck'] },     // Multi-Choice → Array
  { k:'PSA',             cands:['PSA'] },               // Multi-Choice → Array
  { k:'SHBAkzeptiert',   cands:['SHBAkzeptiert'] },     // Ja/Nein
  { k:'Signatur',        cands:['Signatur'] },          // Mehrzeiliger Text (Base64-PNG)
  { k:'Eingangszeit',    cands:['Eingangszeit'] },
  { k:'Abgangszeit',     cands:['Abgangszeit'] },
  { k:'Status',          cands:['Status'] },            // Choice: Angemeldet/Eingecheckt/Geschlossen
  { k:'Bemerkungen',     cands:['Bemerkungen'] },
  { k:'GruppenId',       cands:['GruppenId','Besuchsgruppe'] }
];

async function discoverSP(){
  let site;
  try { site = await gGet(`/sites/${SP_SITE}`); }
  catch(e){ if(/404|itemNotFound|could not be found/i.test(e.message)) throw new Error('SITE_NOT_FOUND'); throw e; }
  siteId = site.id;
  let list;
  try { list = await gGet(`/sites/${siteId}/lists/${encodeURIComponent(SP_LIST)}?$select=id,displayName`); }
  catch(e){ if(/404|itemNotFound|could not be found/i.test(e.message)) throw new Error('LIST_NOT_FOUND'); throw e; }
  listId = list.id;
  const cols = await gGet(`/sites/${siteId}/lists/${listId}/columns?$select=name,displayName`);
  const byName = new Map(), byDisp = new Map();
  (cols.value||[]).forEach(c=>{ byName.set((c.name||'').toLowerCase(), c.name); byDisp.set((c.displayName||'').toLowerCase(), c.name); });
  const missing = [];
  C = {}; HAVE = new Set();
  FIELD_DEFS.forEach(f=>{
    let internal = null;
    for (const cand of f.cands){
      const lc = cand.toLowerCase();
      if (byName.has(lc)) { internal = byName.get(lc); break; }
      if (byDisp.has(lc)) { internal = byDisp.get(lc); break; }
    }
    if (internal) { C[f.k] = internal; HAVE.add(f.k); }
    else { C[f.k] = f.cands[0]; if (f.k!=='GruppenId') missing.push(f.cands[0]); }
  });
  if (missing.length){
    const w = $id('col-warning');
    if (w){ w.style.display='block';
      w.innerHTML = '⚠️ Fehlende SharePoint-Spalten in der Liste „'+esc(SP_LIST)+'": <b>'+missing.map(esc).join(', ')+
        '</b><br>Bitte gemäß SETUP.md anlegen – betroffene Felder werden sonst nicht gespeichert.'; }
    console.warn('[SP] Fehlende Spalten:', missing);
  }
}

// ── ZUGRIFF: WERK + ROLLE ──────────────────────────────────────────────────
function myUPN(){ return (account?.username||'').trim().toLowerCase(); }
function isAdmin(){ return ADMIN_EMAILS.includes(myUPN()); }

async function loadMyIdentities(){
  _meIds = new Set();
  try{
    const me = await gGet('/me?$select=userPrincipalName,mail,otherMails,proxyAddresses');
    const add = v => { if(v) _meIds.add(String(v).replace(/^smtp:/i,'').trim().toLowerCase()); };
    add(me.userPrincipalName); add(me.mail);
    (me.otherMails||[]).forEach(add);
    (me.proxyAddresses||[]).forEach(add);
  }catch(e){ console.warn('[ident]', e.message); }
}
function _myIdentities(){
  const ids = new Set(_meIds||[]); const add=v=>{ if(v) ids.add(String(v).trim().toLowerCase()); };
  add(account?.username);
  const c = account?.idTokenClaims||{}; add(c.email); add(c.preferred_username); add(c.upn);
  if (Array.isArray(c.emails)) c.emails.forEach(add);
  return ids;
}
function myAccess(){
  if (isAdmin()) return { role:'admin', werke: WERKE.slice() };
  for (const id of _myIdentities()){ if (accessUsers[id]) return accessUsers[id]; }
  return null; // kein Zugriff
}
function allowedWerke(){ const a = myAccess(); return a ? a.werke.slice() : []; }
function myRole(){ const a = myAccess(); return a ? a.role : null; }
function canView(werk){ return isAdmin() || allowedWerke().includes(werk); }
function isFull(){ return isAdmin() || FULL_ROLES.includes(myRole()); }   // vollberechtigt
function canCreate(){ return myRole() != null; }        // jede Rolle darf anlegen
function canEditFull(){ return isFull(); }
function canStamp(){ return myRole() != null; }          // jede Rolle darf ein-/auschecken
function canSeeDashboard(){ return isFull(); }
function canSeeReports(){ return isFull(); }
function canManageAccess(){ return isAdmin(); }
// „Eigener" Datensatz: vom aktuellen Nutzer erstellt (E-Mail oder Anzeigename).
function isMine(i){
  const ids = _myIdentities();
  if (i.createdByEmail && ids.has(i.createdByEmail)) return true;
  if (i.createdBy && account?.name && i.createdBy === account.name) return true;
  return false;
}

function _decodeSpText(s){ if(s==null) return ''; const noTags=String(s).replace(/<[^>]*>/g,''); const ta=document.createElement('textarea'); ta.innerHTML=noTags; return ta.value; }

async function _findConfigList(){
  if (accessListId) return accessListId;
  try{
    const all = await gGet(`/sites/${siteId}/lists?$select=id,displayName&$top=500`);
    accessListId = (all.value||[]).find(l => (l.displayName||'').trim().toLowerCase()===ACCESS_LIST_NAME.toLowerCase())?.id || null;
  }catch{ accessListId=null; }
  return accessListId;
}
// Zustand der letzten Konfig-Ladung – für die „Kein Zugriff"-Diagnose:
// 'ok' | 'no-list' (Liste fehlt/nicht sichtbar) | 'read-failed' (keine Leseberechtigung) | 'parse-failed'
let accessLoadState = 'unknown';
async function loadAccessConfig(){
  accessUsers={}; accessConfigItemId=null; accessListId=null; accessLoadState='unknown';
  try{
    if(!siteId){ accessLoadState='no-list'; return; }
    const lid = await _findConfigList();
    if(!lid){ accessLoadState='no-list'; return; }
    let res;
    try{ res = await gGet(`/sites/${siteId}/lists/${lid}/items?$expand=fields&$top=200`); }
    catch(e){ accessLoadState='read-failed'; console.warn('[Zugriff] Liste nicht lesbar:', e.message); return; }
    const item = (res.value||[]).find(it => (it.fields?.Title||'')===ACCESS_ITEM_TITLE);
    if(item){
      accessConfigItemId = item.id;
      let raw = {};
      try{ raw = JSON.parse(_decodeSpText(item.fields?.ConfigValue)||'{}').users||{}; }
      catch(e){ accessLoadState='parse-failed'; console.warn('[Zugriff] JSON-Fehler:', e.message); return; }
      // Schlüssel klein schreiben (robuster Abgleich mit den Nutzer-Identitäten)
      Object.keys(raw).forEach(u=>{ const v=raw[u]||{}; accessUsers[String(u).trim().toLowerCase()]={ role:v.role||'verantwortlicher', werke:Array.isArray(v.werke)?v.werke:[] }; });
    }
    accessLoadState='ok';
  }catch(e){ accessLoadState='read-failed'; console.warn('[Zugriff]', e.message); }
}
async function saveAccessConfig(){
  if(!isAdmin()){ toast('Nur Administrator darf Zugriffsrechte ändern.','error'); return; }
  try{
    const lid = await _findConfigList();
    if(!lid){ toast("Liste '"+ACCESS_LIST_NAME+"' nicht gefunden – siehe SETUP.md.",'error'); return; }
    const fields = { Title: ACCESS_ITEM_TITLE, ConfigValue: JSON.stringify({ users: accessUsers }) };
    if (accessConfigItemId) await gPatch(`/sites/${siteId}/lists/${lid}/items/${accessConfigItemId}/fields`, fields);
    else { const cr = await gPost(`/sites/${siteId}/lists/${lid}/items`, { fields }); accessConfigItemId = cr.id; }
    toast('Zugriffsrechte gespeichert ✓','success');
  }catch(e){
    const hint = /ConfigValue/i.test(e.message) ? " – Spalte 'ConfigValue' (Mehrere Textzeilen) anlegen." : '';
    toast('Speichern fehlgeschlagen: '+e.message+hint,'error');
  }
}

// ── DATEN LADEN + NORMALISIEREN ─────────────────────────────────────────────
function _v(fields, key){ return fields ? fields[C[key]] : undefined; }
function _arr(x){ return Array.isArray(x) ? x : (x==null||x==='' ? [] : [x]); }
function normalize(it){
  const f = it.fields||{};
  return {
    id: it.id,
    werk:        _v(f,'Werk')||'',
    bereich:     _v(f,'Bereich')||'',
    ansprechName:_v(f,'AnsprechName')||'',
    ansprechTel: _v(f,'AnsprechTel')||'',
    besuchsdatum:_v(f,'Besuchsdatum')||'',
    ankunftszeit:_v(f,'Ankunftszeit')||'',
    besucherName:_v(f,'BesucherName')||'',
    firma:       _v(f,'Firma')||'',
    funktion:    _v(f,'Funktion')||'',
    tel:         _v(f,'BesucherTelefon')||'',
    email:       _v(f,'BesucherEmail')||'',
    kennzeichen: _v(f,'Autokennzeichen')||'',
    zweck:       _arr(_v(f,'Besuchszweck')),
    psa:         _arr(_v(f,'PSA')),
    shb:         !!_v(f,'SHBAkzeptiert'),
    signatur:    _decodeSpText(_v(f,'Signatur')||''),
    eingang:     _v(f,'Eingangszeit')||'',
    abgang:      _v(f,'Abgangszeit')||'',
    status:      _v(f,'Status')||'Angemeldet',
    bemerkungen: _v(f,'Bemerkungen')||'',
    gruppenId:   _v(f,'GruppenId')||'',
    created:       it.createdDateTime||'',
    createdBy:      it.createdBy?.user?.displayName||'',
    createdByEmail: (it.createdBy?.user?.email||'').toLowerCase()
  };
}
async function loadItems(force){
  if(!siteId||!listId){ try{ await discoverSP(); }catch(e){ toast('SharePoint nicht erreichbar: '+e.message,'error'); return; } }
  try{
    if(force) toast('Wird aktualisiert …','info');
    const res = await gGet(`/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=500&$orderby=createdDateTime desc`);
    let items = (res.value||[]).map(normalize);
    // Zugriffsgrenze (UX; echte Grenze = SharePoint-Berechtigungen)
    const allowed = allowedWerke();
    if(!isAdmin()) items = items.filter(i => allowed.includes(i.werk));
    ITEMS = items;
    renderCurrentView();
  }catch(e){ toast('Laden fehlgeschlagen: '+e.message,'error'); }
}

// ── NAVIGATION ──────────────────────────────────────────────────────────────
const VIEW_TITLES = { dashboard:'Dashboard', new:'Neue Anmeldung', checkin:'Empfang / Check-in', records:'Eigene Datensätze', reports:'Reports', detail:'Details' };
function navigate(view, arg){
  // Zugriffsgrenzen für gesperrte Bereiche (Nav ist ohnehin ausgeblendet)
  if (view==='dashboard' && !canSeeDashboard()) view = 'new';
  if (view==='reports'   && !canSeeReports())   view = 'new';
  currentView = view;
  document.querySelectorAll('.view').forEach(v=>v.classList.remove('active'));
  const el = $id('view-'+view); if(el) el.classList.add('active');
  document.querySelectorAll('.nav-item').forEach(n=>n.classList.toggle('active', n.dataset.view===view));
  $id('page-title').textContent = VIEW_TITLES[view]||'';
  if (view==='dashboard') renderDashboard();
  else if (view==='new') renderNewForm();
  else if (view==='checkin') renderCheckin();
  else if (view==='records') renderRecords();
  else if (view==='reports') renderReports();
  else if (view==='detail') renderDetail(arg);
}
function renderCurrentView(){
  if (currentView==='dashboard') renderDashboard();
  else if (currentView==='checkin') renderCheckin();
  else if (currentView==='records') renderRecords();
  else if (currentView==='reports') renderReports();
}
function applyNavVisibility(){
  const vis = (v,on)=>{ const el=document.querySelector(`.nav-item[data-view="${v}"]`); if(el) el.style.display = on ? '' : 'none'; };
  vis('dashboard', canSeeDashboard());
  vis('new',       canCreate());
  vis('reports',   canSeeReports());
  const role = myRole();
  const rb = $id('hdr-role');
  if (rb){ rb.style.display=''; rb.textContent = role==='admin' ? 'Administrator' : (ROLLEN[role] || 'Kein Zugriff'); }
  fillWerkFilter();
}
// Werk-Filter im Dashboard mit den freigegebenen Werken füllen
function fillWerkFilter(){
  const sel = $id('dash-werk'); if(!sel) return;
  const cur = sel.value;
  sel.innerHTML = '<option value="">Alle Werke</option>' + allowedWerke().map(w=>`<option value="${esc(w)}">${esc(w)}</option>`).join('');
  sel.value = cur;
}

// ── DASHBOARD ───────────────────────────────────────────────────────────────
function todayStr(){ return new Date().toISOString().slice(0,10); }
function isToday(iso){ return iso && iso.slice(0,10)===todayStr(); }
function fmtDate(iso){ if(!iso) return '–'; const d=new Date(iso); return isNaN(d)?esc(iso):d.toLocaleDateString('de-DE'); }
function fmtDateTime(iso){ if(!iso) return '–'; const d=new Date(iso); return isNaN(d)?esc(iso):d.toLocaleString('de-DE',{day:'2-digit',month:'2-digit',year:'2-digit',hour:'2-digit',minute:'2-digit'}); }
function fmtTime(iso){ if(!iso) return '–'; const d=new Date(iso); return isNaN(d)?esc(iso):d.toLocaleTimeString('de-DE',{hour:'2-digit',minute:'2-digit'}); }

function statusBadge(s){ const cls={'Angemeldet':'status-angemeldet','Eingecheckt':'status-eingecheckt','Geschlossen':'status-geschlossen'}[s]||'status-angemeldet'; return `<span class="status-badge ${cls}">${esc(s)}</span>`; }
function werkBadge(w){ return `<span class="werk-badge">${esc(w||'?')}</span>`; }

function renderDashboard(){
  if(!canSeeDashboard()){ $id('dash-list').innerHTML = `<div class="empty-state">Das Dashboard ist nur für vollberechtigte Rollen sichtbar.</div>`; $id('dash-stats').innerHTML=''; return; }
  const onSite = ITEMS.filter(i => i.status==='Eingecheckt');
  const todayReg = ITEMS.filter(i => isToday(i.besuchsdatum) || isToday(i.created));
  const closedToday = ITEMS.filter(i => i.status==='Geschlossen' && isToday(i.abgang));
  $id('dash-stats').innerHTML = `
    <div class="stat-card on-site"><div class="stat-num">${onSite.length}</div><div class="stat-lbl">Aktuell anwesend</div></div>
    <div class="stat-card today"><div class="stat-num">${todayReg.length}</div><div class="stat-lbl">Anmeldungen heute</div></div>
    <div class="stat-card"><div class="stat-num">${closedToday.length}</div><div class="stat-lbl">Heute abgemeldet</div></div>
    <div class="stat-card"><div class="stat-num">${ITEMS.length}</div><div class="stat-lbl">Datensätze gesamt</div></div>`;

  const q = ($id('search-dashboard')?.value||'').toLowerCase();
  const status = $id('dash-status')?.value||'';
  const werk = $id('dash-werk')?.value||'';
  const list = ITEMS.filter(i => recordMatches(i, q, status, werk));
  const body = $id('dash-list');
  body.innerHTML = list.length ? list.map(recordCard).join('') : `<div class="empty-state">Keine Datensätze für diesen Filter.</div>`;
}
// Gemeinsame Filter-/Karten-Helfer für Dashboard und Eigene Datensätze
function recordMatches(i, q, status, werk){
  if (status && i.status !== status) return false;
  if (werk && i.werk !== werk) return false;
  if (q && !((i.besucherName+' '+i.firma+' '+i.werk+' '+i.bereich+' '+i.kennzeichen+' '+i.ansprechName).toLowerCase().includes(q))) return false;
  return true;
}
function recordCard(i){
  const stamp = canStamp() ? (i.status==='Angemeldet'
      ? `<button class="btn btn-sm btn-success" onclick="event.stopPropagation();checkIn('${i.id}')">Einchecken</button>`
      : (i.status==='Eingecheckt' ? `<button class="btn btn-sm btn-outline" onclick="event.stopPropagation();checkOut('${i.id}')">Abmelden</button>` : '')) : '';
  return `<div class="visitor-card" style="cursor:pointer" onclick="navigate('detail','${i.id}')">
    <div class="vc-main">
      <div class="vc-name">${esc(i.besucherName)} ${werkBadge(i.werk)} ${statusBadge(i.status)}</div>
      <div class="vc-sub">${esc(i.firma||'–')} · ${esc(i.bereich||'')} · ${fmtDate(i.besuchsdatum)} · Ein ${fmtTime(i.eingang)} / Ab ${fmtTime(i.abgang)}</div>
    </div>
    <div class="vc-actions">${stamp}${canCreate()?`<button class="mini-btn" onclick="event.stopPropagation();useAsTemplate('${i.id}')">Vorlage</button>`:''}<span class="mini-btn">Öffnen →</span></div>
  </div>`;
}

// ── NEUE ANMELDUNG ──────────────────────────────────────────────────────────
function renderNewForm(){
  const host = $id('view-new');
  if(!canCreate()){ host.innerHTML = `<div class="empty-state">Für das Anlegen von Anmeldungen fehlt dir die Berechtigung (Rolle: SHB-Verantwortlicher oder Sekretariat).</div>`; return; }
  newVisitorSeq = 0;
  const werkOpts = allowedWerke().map(w=>`<option value="${esc(w)}">${esc(w)}</option>`).join('');
  const zweckBoxes = BESUCHSZWECKE.map(z=>`<label><input type="checkbox" name="zweck" value="${esc(z)}"> ${esc(z)}</label>`).join('');
  host.innerHTML = `
  <div class="form-card">
    <h2>Neue Besucheranmeldung</h2>
    <div class="fc-sub">Vom besuchten Bereich auszufüllen. Datensatz möglichst vorab erzeugen.</div>
    <div class="form-grid">
      <div class="form-group"><label>Werk <span class="req">*</span></label><select id="f-werk">${werkOpts}</select></div>
      <div class="form-group"><label>Bereich <span class="req">*</span></label><input id="f-bereich" type="text"></div>
      <div class="form-group"><label>Ansprechpartner im Werk <span class="req">*</span></label><input id="f-ansprech" type="text" value="${esc(account?.name||'')}"></div>
      <div class="form-group"><label>Telefon (Ansprechpartner)</label><input id="f-ansprechtel" type="text"></div>
      <div class="form-group"><label>Besuchsdatum <span class="req">*</span></label><input id="f-datum" type="date" value="${todayStr()}"></div>
      <div class="form-group"><label>Ankunftszeit (geplant)</label><input id="f-ankunft" type="time"></div>
      <div class="form-group full"><label>Firma <span class="req">*</span></label><input id="f-firma" type="text"><div class="field-sub">Gilt für alle Personen dieser Anmeldung.</div></div>
    </div>

    <div class="form-sub-h">Besuchszweck</div>
    <div class="zweck-grid">${zweckBoxes}</div>

    <div class="form-sub-h">Besucher</div>
    <div id="visitors"></div>
    <button class="btn btn-sm btn-outline" onclick="addVisitorRow()">+ weitere Person (gleiche Firma)</button>

    <div class="form-sub-h">Sicherheitsunterweisung (SHB)</div>
    <label style="display:flex;align-items:center;gap:8px;font-size:.88rem"><input type="checkbox" id="f-shb"> Sicherheitshinweise SHB akzeptiert</label>
    <div id="sig-block" style="margin-top:10px;display:none">
      <div class="field-sub" style="margin-bottom:4px">Digitale Unterschrift</div>
      <div class="sig-wrap"><canvas id="sig" class="sig-canvas" width="420" height="150"></canvas></div>
      <div class="sig-actions"><button class="mini-btn" onclick="clearSig()">Löschen</button></div>
    </div>

    <div class="form-group full" style="margin-top:14px"><label>Bemerkungen</label><textarea id="f-bemerk" rows="2"></textarea></div>

    <div class="privacy-note">
      <b>Datenschutzhinweis:</b> Die Daten werden ausschließlich zur Werks-/Besuchersicherheit und zum Nachweis der
      Sicherheitsunterweisung verarbeitet (Art. 6 (1) f, c DSGVO), nur berechtigten Werken zugänglich gemacht und nach
      <b>90 Tagen automatisch gelöscht</b>. <button class="mini-btn" onclick="showPrivacyNotice()">Vollständigen Hinweis anzeigen</button>
    </div>

    <div class="card-actions">
      <button class="btn btn-primary" onclick="submitNew()">Anmeldung speichern</button>
      <button class="btn btn-ghost" onclick="goHome()">Abbrechen</button>
    </div>
  </div>`;
  addVisitorRow(true);
  setupSignature();
  $id('f-shb').addEventListener('change', e=>{ $id('sig-block').style.display = e.target.checked ? 'block' : 'none'; });
  if (templateItem){ applyTemplate(templateItem); templateItem = null; }
}
function goHome(){ navigate(canSeeDashboard() ? 'dashboard' : 'records'); }

// „Als Vorlage": bestehenden Datensatz als Basis für eine neue Anmeldung nutzen.
function useAsTemplate(id){
  const it = ITEMS.find(x=>x.id===id);
  if(!it){ toast('Datensatz nicht gefunden.','error'); return; }
  if(!canCreate()){ toast('Keine Berechtigung zum Anlegen.','error'); return; }
  templateItem = it;
  navigate('new');
}
function applyTemplate(it){
  const setv=(id,v)=>{ const el=$id(id); if(el&&v) el.value=v; };
  if(allowedWerke().includes(it.werk)) setv('f-werk', it.werk);
  setv('f-bereich', it.bereich);
  setv('f-ansprech', it.ansprechName);
  setv('f-ansprechtel', it.ansprechTel);
  setv('f-firma', it.firma);
  (it.zweck||[]).forEach(z=>{ const cb=document.querySelector(`input[name="zweck"][value="${z}"]`); if(cb) cb.checked=true; });
  const row = document.querySelector('#visitors .visitor-row');
  if(row){
    const sv=(sel,v)=>{ const el=row.querySelector(sel); if(el&&v) el.value=v; };
    sv('[data-f="name"]', it.besucherName);
    sv('[data-f="funktion"]', it.funktion);
    sv('[data-f="tel"]', it.tel);
    sv('[data-f="email"]', it.email);
    sv('[data-f="kennzeichen"]', it.kennzeichen);
    (it.psa||[]).forEach(p=>{ const cb=row.querySelector(`[data-psa][value="${p}"]`); if(cb) cb.checked=true; });
  }
  toast('Aus Vorlage übernommen – Datum prüfen und neu unterschreiben.','info');
}
function addVisitorRow(first){
  const seq = ++newVisitorSeq;
  const psaBoxes = PSA_LISTE.map(p=>`<label><input type="checkbox" data-psa value="${esc(p)}"> ${esc(p)}</label>`).join('');
  const div = document.createElement('div');
  div.className = 'visitor-row'; div.dataset.seq = seq;
  div.innerHTML = `
    <div class="vr-head">
      <span class="vr-title">Person ${seq}</span>
      ${first?'' : `<button class="vr-del" onclick="this.closest('.visitor-row').remove()" title="Entfernen">✕</button>`}
    </div>
    <div class="form-grid">
      <div class="form-group"><label>Name, Vorname <span class="req">*</span></label><input type="text" data-f="name"></div>
      <div class="form-group"><label>Funktion</label><input type="text" data-f="funktion"></div>
      <div class="form-group"><label>Telefon</label><input type="text" data-f="tel"></div>
      <div class="form-group"><label>E-Mail</label><input type="email" data-f="email"></div>
      <div class="form-group"><label>Autokennzeichen</label><input type="text" data-f="kennzeichen"></div>
    </div>
    <div style="font-size:.8rem;font-weight:600;color:#374151;margin:8px 0 2px">Ausgabe PSA</div>
    <div class="psa-grid">${psaBoxes}</div>`;
  $id('visitors').appendChild(div);
}

// Signatur (Canvas)
let sigCtx=null, sigDrawing=false, sigHasInk=false;
function setupSignature(){
  const cv = $id('sig'); if(!cv) return;
  sigCtx = cv.getContext('2d'); sigCtx.lineWidth=2; sigCtx.lineCap='round'; sigCtx.strokeStyle='#1e2939'; sigHasInk=false;
  const pos = e => { const r=cv.getBoundingClientRect(); const p=(e.touches?e.touches[0]:e); return { x:(p.clientX-r.left)*(cv.width/r.width), y:(p.clientY-r.top)*(cv.height/r.height) }; };
  const start = e => { sigDrawing=true; const {x,y}=pos(e); sigCtx.beginPath(); sigCtx.moveTo(x,y); e.preventDefault(); };
  const move  = e => { if(!sigDrawing) return; const {x,y}=pos(e); sigCtx.lineTo(x,y); sigCtx.stroke(); sigHasInk=true; e.preventDefault(); };
  const end   = () => { sigDrawing=false; };
  cv.addEventListener('mousedown',start); cv.addEventListener('mousemove',move); window.addEventListener('mouseup',end);
  cv.addEventListener('touchstart',start,{passive:false}); cv.addEventListener('touchmove',move,{passive:false}); cv.addEventListener('touchend',end);
}
function clearSig(){ const cv=$id('sig'); if(cv&&sigCtx){ sigCtx.clearRect(0,0,cv.width,cv.height); sigHasInk=false; } }

// Setzt ein Feld nur, wenn die Spalte existiert und der Wert nicht leer ist.
// Verhindert Graph-400 „Invalid request" durch nicht vorhandene Spalten.
// Arrays (Mehrfachauswahl) werden mit @odata.type annotiert – ohne diese
// Annotation lehnt Graph Multi-Choice-/Multi-Text-Spalten mit 400 ab.
function putField(fields, key, value){
  if (!HAVE.has(key)) return;
  if (value === undefined || value === null) return;
  if (typeof value === 'string' && value.trim() === '') return;
  if (Array.isArray(value)){
    if (value.length === 0) return;
    fields[C[key]] = value;
    fields[C[key] + '@odata.type'] = 'Collection(Edm.String)';
    return;
  }
  fields[C[key]] = value;
}

// Beschriftungen + Typ-Hinweise für die Fehlerdiagnose
const FIELD_LABELS = {
  BesucherName:'Besucher (Titel)', Werk:'Werk', Bereich:'Bereich', AnsprechName:'Ansprechpartner-Name',
  AnsprechTel:'Ansprechpartner-Telefon', Besuchsdatum:'Besuchsdatum', Ankunftszeit:'Ankunftszeit',
  Firma:'Firma', Funktion:'Funktion', BesucherTelefon:'Besucher-Telefon', BesucherEmail:'E-Mail',
  Autokennzeichen:'Autokennzeichen', Besuchszweck:'Besuchszweck', PSA:'PSA', SHBAkzeptiert:'SHB akzeptiert',
  Signatur:'Signatur', Status:'Status', Bemerkungen:'Bemerkungen', GruppenId:'GruppenId'
};
function typeHint(key, value){
  if (key==='Besuchszweck'||key==='PSA') return 'Spalte muss Typ <b>Auswahl</b> mit <b>Mehrfachauswahl</b> sein.';
  if (key==='Werk'||key==='Status') return `Spalte muss Typ <b>Auswahl</b> sein und den Wert <b>„${esc(String(value))}"</b> als Option enthalten.`;
  if (/datum|zeit/i.test(key)) return 'Spalte muss Typ <b>Datum/Uhrzeit</b> sein.';
  if (key==='SHBAkzeptiert') return 'Spalte muss Typ <b>Ja/Nein</b> sein.';
  if (key==='Signatur'||key==='Bemerkungen') return 'Spalte muss Typ <b>Mehrere Textzeilen</b> sein.';
  return 'Spalte muss Typ <b>Einzelne Textzeile</b> sein.';
}

async function submitNew(){
  const werk = $id('f-werk').value;
  const bereich = $id('f-bereich').value.trim();
  const ansprech = $id('f-ansprech').value.trim();
  const datum = $id('f-datum').value;
  const firma = $id('f-firma').value.trim();
  const shb = $id('f-shb').checked;
  if(!werk || !bereich || !ansprech || !datum || !firma){ toast('Bitte Pflichtfelder (Werk, Bereich, Ansprechpartner, Datum, Firma) ausfüllen.','error'); return; }
  if(!allowedWerke().includes(werk)){ toast('Keine Berechtigung für dieses Werk.','error'); return; }
  const zweck = [...document.querySelectorAll('input[name="zweck"]:checked')].map(c=>c.value);
  if(!shb){ toast('Sicherheitsunterweisung (SHB) muss bestätigt werden.','error'); return; }
  if(shb && !sigHasInk){ toast('Bitte digital unterschreiben (SHB-Bestätigung).','error'); return; }
  const rows = [...document.querySelectorAll('#visitors .visitor-row')];
  const visitors = rows.map(r=>({
    name: r.querySelector('[data-f="name"]').value.trim(),
    funktion: r.querySelector('[data-f="funktion"]').value.trim(),
    tel: r.querySelector('[data-f="tel"]').value.trim(),
    email: r.querySelector('[data-f="email"]').value.trim(),
    kennzeichen: r.querySelector('[data-f="kennzeichen"]').value.trim(),
    psa: [...r.querySelectorAll('[data-psa]:checked')].map(c=>c.value)
  })).filter(v=>v.name);
  if(!visitors.length){ toast('Mindestens eine Person mit Namen erfassen.','error'); return; }

  const sig = sigHasInk ? $id('sig').toDataURL('image/png') : '';
  const gruppenId = 'G'+Date.now().toString(36)+Math.random().toString(36).slice(2,6);
  const ankunftIso = $id('f-ankunft').value ? new Date(datum+'T'+$id('f-ankunft').value).toISOString() : null;

  const ansprechTel = $id('f-ansprechtel').value.trim();
  const bemerk = $id('f-bemerk').value.trim();
  const common = { werk, bereich, ansprech, ansprechTel, datum, ankunftIso, firma, zweck, sig, bemerk, gruppenId };
  const btn = event?.target; if(btn){ btn.disabled=true; btn.textContent='Speichert …'; }
  const created = [];
  try{
    for (const v of visitors){
      const pairs = visitorFieldPairs(v, common);
      try{
        await gPost(`/sites/${siteId}/lists/${listId}/items`, { fields: pairsToFields(pairs) });
      }catch(e){
        if(/\b400\b/.test(e.message)){ if(btn){ btn.disabled=false; btn.textContent='Anmeldung speichern'; } await diagnoseAndReport(pairs); return; }
        throw e;
      }
      created.push(v);
    }
    toast(`${visitors.length} Anmeldung(en) gespeichert ✓`,'success');
    await loadItems();
    showPostSaveModal({ werk, bereich, ansprech, datum, firma, zweck }, created);
  }catch(e){ toast('Speichern fehlgeschlagen: '+e.message,'error'); if(btn){ btn.disabled=false; btn.textContent='Anmeldung speichern'; } }
}

// Logische (Key,Wert)-Paare eines Besuchers – Grundlage für Speichern & Diagnose.
function visitorFieldPairs(v, c){
  return [
    ['BesucherName', v.name],
    ['Werk', c.werk],
    ['Bereich', c.bereich],
    ['AnsprechName', c.ansprech],
    ['AnsprechTel', c.ansprechTel],
    ['Besuchsdatum', new Date(c.datum+'T00:00:00').toISOString()],
    ['Ankunftszeit', c.ankunftIso],
    ['Firma', c.firma],
    ['Funktion', v.funktion],
    ['BesucherTelefon', v.tel],
    ['BesucherEmail', v.email],
    ['Autokennzeichen', v.kennzeichen],
    ['Besuchszweck', c.zweck],
    ['PSA', v.psa],
    ['SHBAkzeptiert', true],
    ['Signatur', c.sig],
    ['Status', 'Angemeldet'],
    ['Bemerkungen', c.bemerk],
    ['GruppenId', c.gruppenId]
  ];
}
function pairsToFields(pairs){ const f={}; pairs.forEach(([k,v])=>putField(f,k,v)); return f; }

// Bei 400: minimalen Datensatz anlegen und jedes Feld einzeln per PATCH testen,
// um die von SharePoint abgelehnte(n) Spalte(n) exakt zu benennen. Danach aufräumen.
async function diagnoseAndReport(pairs){
  toast('Analysiere abgelehnte Felder …','info');
  const itemsUrl = `/sites/${siteId}/lists/${listId}/items`;
  const titleVal = (pairs.find(p=>p[0]==='BesucherName')?.[1]) || 'Diagnose (wird gelöscht)';
  let testId = null;
  try{
    const cr = await gPost(itemsUrl, { fields: pairsToFields([['BesucherName', titleVal]]) });
    testId = cr.id;
  }catch(e){ showDiagModal([], 'Selbst ein minimaler Datensatz (nur Titel) wird abgelehnt: '+e.message); return; }
  const bad = [];
  for (const [k,v] of pairs){
    if (k==='BesucherName') continue;
    const mini = pairsToFields([[k,v]]);
    if (!Object.keys(mini).length) continue;   // Spalte fehlt oder leer → wird ohnehin nicht gesendet
    try{ await gPatch(`${itemsUrl}/${testId}/fields`, mini); }
    catch(e){ if(/\b400\b/.test(e.message)) bad.push([k,v]); }
  }
  try{ await gDelete(`${itemsUrl}/${testId}`); }catch{}
  showDiagModal(bad);
}
function showDiagModal(bad, fatal){
  $id('modal-title').textContent = 'Speichern abgelehnt – Diagnose';
  if (fatal){
    $id('modal-body').innerHTML = `<p style="font-size:.88rem;color:#b91c1c">${esc(fatal)}</p>`;
  } else if (!bad.length){
    $id('modal-body').innerHTML = `<p style="font-size:.88rem;color:#374151">Kein einzelnes Feld ließ sich isolieren – bitte Konsole prüfen. Häufige Ursache: eine Auswahl-Spalte ohne passende Optionen.</p>`;
  } else {
    const rows = bad.map(([k,v])=>`<li style="margin-bottom:8px"><b>${esc(FIELD_LABELS[k]||k)}</b>${(v!=null&&v!=='')?` (Wert: ${esc(Array.isArray(v)?v.join(', '):String(v).slice(0,40))})`:''}<br><span class="dsgvo-hint">${typeHint(k,v)}</span></li>`).join('');
    $id('modal-body').innerHTML = `<p style="font-size:.88rem;color:#374151;margin-bottom:8px">SharePoint lehnt diese Spalte(n) ab – Typ/Optionen in der Liste <b>${esc(SP_LIST)}</b> prüfen:</p><ul style="padding-left:18px">${rows}</ul>`;
  }
  $id('modal-footer').innerHTML = `<button class="btn btn-primary" onclick="closeModal()">Verstanden</button>`;
  $id('modal-overlay').classList.remove('hidden');
}

// ── EINLADUNG PER E-MAIL (mailto – ohne Zusatzberechtigung) ──────────────────
function inviteBody(head, v){
  const L = [];
  L.push(`Guten Tag ${v.name},`);
  L.push('');
  L.push(`wir laden Sie zu einem Besuch bei DIHAG (${head.werk}) ein.`);
  L.push('');
  L.push(`Datum: ${new Date(head.datum+'T00:00:00').toLocaleDateString('de-DE')}`);
  L.push(`Bereich: ${head.bereich}`);
  L.push(`Ansprechpartner: ${head.ansprech}`);
  if (head.zweck && head.zweck.length) L.push(`Zweck: ${head.zweck.join(', ')}`);
  L.push('');
  L.push('Sicherheitshinweise (SHB): Auf dem Werksgelände gelten die Sicherheits- und PSA-Vorschriften.');
  L.push('Die Sicherheitsunterweisung ist vor Betreten des Geländes am Empfang zu bestätigen (digitale Unterschrift).');
  L.push('Bitte bringen Sie einen Lichtbildausweis mit; bei Anreise mit PKW bitte das Kennzeichen am Empfang angeben.');
  L.push('');
  L.push('Mit freundlichen Grüßen');
  L.push(head.ansprech);
  return L.join('\r\n');
}
function inviteMailto(head, v){
  const subject = `Einladung zum Besuch bei DIHAG – ${new Date(head.datum+'T00:00:00').toLocaleDateString('de-DE')}`;
  const href = `mailto:${encodeURIComponent(v.email||'')}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(inviteBody(head,v))}`;
  window.location.href = href;
}
function showPostSaveModal(head, visitors){
  $id('modal-title').textContent = 'Anmeldung gespeichert';
  const rows = visitors.map((v,i)=>`
    <div class="visitor-card" style="margin-bottom:8px">
      <div class="vc-main"><div class="vc-name">${esc(v.name)}</div>
        <div class="vc-sub">${v.email?esc(v.email):'<span style="color:#b45309">keine E-Mail hinterlegt</span>'}</div></div>
      <div class="vc-actions">${v.email
        ? `<button class="btn btn-sm btn-primary" onclick='inviteMailto(${JSON.stringify(head).replace(/'/g,"&#39;")}, ${JSON.stringify(v).replace(/'/g,"&#39;")})'>✉ Einladung</button>`
        : ''}</div>
    </div>`).join('');
  $id('modal-body').innerHTML = `
    <p style="font-size:.86rem;color:#374151;margin-bottom:10px">Einladungs-E-Mail an die Besucher senden (öffnet dein Outlook mit vorbereitetem Text inkl. Sicherheitshinweis):</p>
    ${rows}
    <p class="dsgvo-hint" style="margin-top:8px">Die digitale Unterschrift zur Sicherheitsunterweisung erfolgt beim Check-in am Empfang. Ein vorheriges digitales Signieren durch externe Besucher ist als Ausbaustufe möglich (separater Signatur-Link).</p>`;
  $id('modal-footer').innerHTML = `<button class="btn btn-primary" onclick="closeModal();navigate('checkin')">Weiter zum Empfang</button>`;
  $id('modal-overlay').classList.remove('hidden');
}

// ── CHECK-IN / CHECK-OUT ────────────────────────────────────────────────────
async function checkIn(id){
  if(!canStamp()){ toast('Keine Berechtigung zum Stempeln.','error'); return; }
  const fields={}; putField(fields,'Eingangszeit',new Date().toISOString()); putField(fields,'Status','Eingecheckt');
  try{ await gPatch(`/sites/${siteId}/lists/${listId}/items/${id}/fields`, fields);
    toast('Eingecheckt ✓','success'); await loadItems(); }
  catch(e){ toast('Fehler: '+e.message,'error'); }
}
async function checkOut(id){
  if(!canStamp()){ toast('Keine Berechtigung zum Stempeln.','error'); return; }
  const it = ITEMS.find(i=>i.id===id);
  // Schließbedingung: Pflichtfelder vorhanden
  if(it && (!it.werk || !it.bereich || !it.besucherName || !it.firma || !it.shb)){
    toast('Datensatz unvollständig – kann nicht geschlossen werden (Pflichtfelder/SHB).','error'); return;
  }
  const fields={}; putField(fields,'Abgangszeit',new Date().toISOString()); putField(fields,'Status','Geschlossen');
  try{ await gPatch(`/sites/${siteId}/lists/${listId}/items/${id}/fields`, fields);
    toast('Abgemeldet & geschlossen ✓','success'); await loadItems(); }
  catch(e){ toast('Fehler: '+e.message,'error'); }
}

function renderCheckin(){
  const q = ($id('search-checkin')?.value||'').toLowerCase();
  const rel = ITEMS.filter(i => i.status!=='Geschlossen')
    .filter(i => !q || (i.besucherName+' '+i.firma+' '+i.werk).toLowerCase().includes(q));
  const body = $id('checkin-body');
  if(!rel.length){ body.innerHTML = `<div class="empty-state">Keine offenen Anmeldungen.</div>`; return; }
  // Nach Gruppe bündeln
  const groups = {};
  rel.forEach(i=>{ (groups[i.gruppenId||i.id] ||= []).push(i); });
  body.innerHTML = Object.values(groups).map(g=>{
    const head = g[0];
    const rows = g.map(i=>{
      const act = canStamp()
        ? (i.status==='Angemeldet'
            ? `<button class="btn btn-sm btn-success" onclick="checkIn('${i.id}')">Einchecken</button>`
            : `<button class="btn btn-sm btn-outline" onclick="checkOut('${i.id}')">Abmelden</button>`)
        : '';
      return `<div class="visitor-card">
        <div class="vc-main"><div class="vc-name">${esc(i.besucherName)} ${statusBadge(i.status)}</div>
          <div class="vc-sub">${i.eingang?('Eingang '+fmtTime(i.eingang)) : 'noch nicht eingecheckt'}${i.psa.length?(' · PSA: '+i.psa.map(esc).join(', ')):''}</div></div>
        <div class="vc-actions"><button class="mini-btn" onclick="navigate('detail','${i.id}')">Details</button>${act}</div>
      </div>`;
    }).join('');
    return `<div class="form-card" style="max-width:none;padding:16px 18px">
      <div style="display:flex;justify-content:space-between;flex-wrap:wrap;gap:6px;margin-bottom:8px">
        <div><b>${esc(head.firma||'–')}</b> ${werkBadge(head.werk)} · ${esc(head.bereich||'')}</div>
        <div class="vc-sub">${fmtDate(head.besuchsdatum)} · Gastgeber: ${esc(head.ansprechName||'–')}</div>
      </div>${rows}</div>`;
  }).join('');
}

// ── DATENSÄTZE ──────────────────────────────────────────────────────────────
function renderRecords(){
  const q = ($id('search-records')?.value||'').toLowerCase();
  const sf = $id('filter-status')?.value||'';
  const list = ITEMS.filter(isMine).filter(i => recordMatches(i, q, sf, ''));
  const body = $id('records-body');
  if(!list.length){ body.innerHTML = `<div class="empty-state">Noch keine eigenen Datensätze – lege unter „Neue Anmeldung" welche an.</div>`; return; }
  body.innerHTML = list.map(recordCard).join('');
}

// ── DETAIL ──────────────────────────────────────────────────────────────────
function renderDetail(id){
  const i = ITEMS.find(x=>x.id===id);
  const host = $id('detail-content');
  if(!i){ host.innerHTML = `<div class="empty-state">Datensatz nicht gefunden.</div>`; return; }
  const row=(l,v)=>`<div class="dl">${esc(l)}</div><div class="dv">${v}</div>`;
  const stamp = canStamp() ? (i.status==='Angemeldet'
      ? `<button class="btn btn-sm btn-success" onclick="checkIn('${i.id}')">Einchecken</button>`
      : (i.status==='Eingecheckt' ? `<button class="btn btn-sm btn-outline" onclick="checkOut('${i.id}')">Abmelden</button>` : '')) : '';
  const del = isAdmin() ? `<button class="btn btn-sm btn-danger" onclick="deleteItem('${i.id}')">Löschen</button>` : '';
  const tmpl = canCreate() ? `<button class="btn btn-sm btn-outline" onclick="useAsTemplate('${i.id}')">Als Vorlage</button>` : '';
  const spUrl = `https://${SP_SITE.split(':/')[0]}/${SP_SITE.split(':/')[1]}/Lists/${encodeURIComponent(SP_LIST)}/AllItems.aspx`;
  host.innerHTML = `
    <div class="detail-head"><h2>${esc(i.besucherName)}</h2> ${werkBadge(i.werk)} ${statusBadge(i.status)}</div>
    <div class="card-actions" style="margin-top:4px">${stamp} ${tmpl} ${del}
      <a class="btn btn-sm btn-ghost" href="${spUrl}" target="_blank" rel="noopener">Versionsverlauf in SharePoint</a></div>
    <div class="detail-grid">
      ${row('Firma', esc(i.firma||'–'))}
      ${row('Funktion', esc(i.funktion||'–'))}
      ${row('Bereich', esc(i.bereich||'–'))}
      ${row('Ansprechpartner', esc(i.ansprechName||'–')+(i.ansprechTel?(' · '+esc(i.ansprechTel)):''))}
      ${row('Besuchsdatum', fmtDate(i.besuchsdatum))}
      ${row('Besuchszweck', i.zweck.length?i.zweck.map(esc).join(', '):'–')}
      ${row('PSA', i.psa.length?i.psa.map(esc).join(', '):'–')}
      ${row('Kontakt', (i.tel||i.email)?[i.tel,i.email].filter(Boolean).map(esc).join(' · '):'–')}
      ${row('Autokennzeichen', esc(i.kennzeichen||'–'))}
      ${row('Eingangszeit', fmtDateTime(i.eingang))}
      ${row('Abgangszeit', fmtDateTime(i.abgang))}
      ${row('SHB akzeptiert', i.shb?'✅ Ja':'—')}
      ${row('Bemerkungen', esc(i.bemerkungen||'–'))}
      ${row('Angelegt', fmtDateTime(i.created)+(i.createdBy?(' · '+esc(i.createdBy)):''))}
    </div>
    ${i.signatur?`<div class="field-sub">Unterschrift SHB</div><img src="${esc(i.signatur)}" alt="Unterschrift" style="border:1px solid #e5e7eb;border-radius:8px;max-width:420px;width:100%">`:''}
  `;
}
async function deleteItem(id){
  if(!isAdmin()){ toast('Nur Administrator darf löschen.','error'); return; }
  if(!confirm('Diesen Datensatz endgültig löschen?')) return;
  try{ await gDelete(`/sites/${siteId}/lists/${listId}/items/${id}`); toast('Gelöscht.','success'); await loadItems(); navigate('records'); }
  catch(e){ toast('Löschen fehlgeschlagen: '+e.message,'error'); }
}

// ── REPORTS ─────────────────────────────────────────────────────────────────
function renderReports(){
  const body = $id('reports-body');
  if(!canSeeReports()){ body.innerHTML = `<div class="empty-state">Reports sind nur für vollberechtigte Rollen sichtbar.</div>`; return; }
  const per = {}; allowedWerke().forEach(w=>per[w]={ges:0,anwesend:0});
  ITEMS.forEach(i=>{ if(!per[i.werk]) per[i.werk]={ges:0,anwesend:0}; per[i.werk].ges++; if(i.status==='Eingecheckt') per[i.werk].anwesend++; });
  const rows = Object.entries(per).sort((a,b)=>b[1].ges-a[1].ges)
    .map(([w,s])=>`<div class="visitor-card"><div class="vc-main"><div class="vc-name">${werkBadge(w)}</div>
      <div class="vc-sub">${s.ges} Datensätze (letzte 90 Tage) · aktuell anwesend: ${s.anwesend}</div></div></div>`).join('');
  body.innerHTML = `<h3 class="section-h">Übersicht nach Werk</h3>${rows||'<div class="empty-state">Keine Daten.</div>'}
    <div class="privacy-note" style="margin-top:16px">Reports zeigen nur aggregierte Zahlen deiner freigegebenen Werke. Aufbewahrung: 90 Tage, danach automatische Löschung.</div>`;
}

// ── EINSTELLUNGEN ───────────────────────────────────────────────────────────
function openSettings(){
  const role = myRole();
  const myWerke = allowedWerke();
  const meBlock = `
    <div class="settings-section-title">👤 Mein Zugriff</div>
    <div class="detail-grid" style="grid-template-columns:130px 1fr">
      <div class="dl">Rolle</div><div class="dv">${role==='admin'?'Administrator':(ROLLEN[role]||'Kein Zugriff')}</div>
      <div class="dl">Werke</div><div class="dv">${myWerke.length?myWerke.map(werkBadge).join(' '):'–'}</div>
    </div>
    <button class="btn btn-sm btn-ghost" onclick="showPrivacyNotice()">🔒 Datenschutzhinweis (Aushang Empfang)</button>`;

  // ── Rechtemodell (Illustration, für alle sichtbar) ──
  const currentKey = isAdmin() ? 'admin' : role;
  const rr = (key, name, caps) => {
    const active = key === currentKey;
    const cells = caps.map(c=>`<td style="padding:5px 8px;text-align:center">${c?'✅':'–'}</td>`).join('');
    return `<tr style="${active?'background:#eff6ff;font-weight:600':''}"><td style="padding:5px 8px;white-space:nowrap">${esc(name)}${active?' ◀ Du':''}</td>${cells}</tr>`;
  };
  const rechteBlock = `
    <hr class="modal-hr">
    <div class="settings-section-title">🧭 Rechtemodell</div>
    <div style="overflow-x:auto">
      <table style="border-collapse:collapse;font-size:.76rem;width:100%;min-width:520px">
        <thead><tr style="color:#6b7280;border-bottom:1px solid #e5e7eb">
          <th style="padding:5px 8px;text-align:left">Rolle</th>
          <th style="padding:5px 8px">Anlegen</th>
          <th style="padding:5px 8px">Ein-/Aus&shy;checken</th>
          <th style="padding:5px 8px">Eigene Datens&auml;tze</th>
          <th style="padding:5px 8px">Dashboard (alle)</th>
          <th style="padding:5px 8px">Reports</th>
          <th style="padding:5px 8px">Rechte&shy;verwaltung</th>
        </tr></thead>
        <tbody>
          ${rr('verantwortlicher','SHB-Verantwortlicher',[1,1,1,0,0,0])}
          ${rr('wachschutz','Wachschutz (voll)',[1,1,1,1,1,0])}
          ${rr('sekretariat','Sekretariat (voll)',[1,1,1,1,1,0])}
          ${rr('admin','Administrator',[1,1,1,1,1,1])}
        </tbody>
      </table>
    </div>
    <p class="dsgvo-hint" style="margin-top:6px">Zugriff zus&auml;tzlich je Werk begrenzt. „Eigene Datens&auml;tze" = nur selbst erstellte. Neue Nutzer starten als SHB-Verantwortlicher.</p>`;

  let adminBlock = '';
  if (canManageAccess()){
    const roleOpts = r => Object.entries(ROLLEN).map(([k,l])=>`<option value="${k}" ${r===k?'selected':''}>${esc(l)}</option>`).join('');
    const werkChecks = (upn, werke) => WERKE.map(w=>`<label><input type="checkbox" ${werke.includes(w)?'checked':''} onchange="toggleWerk('${esc(upn)}','${w}',this.checked)"> ${w}</label>`).join('');
    const userBlocks = Object.entries(accessUsers).sort((a,b)=>a[0].localeCompare(b[0])).map(([upn,p])=>`
      <div class="visitor-row" style="background:#fff">
        <div class="vr-head"><span class="vr-title">${esc(upn)}</span>
          <button class="vr-del" onclick="removeAccessUser('${esc(upn)}')" title="Entfernen">✕</button></div>
        <div class="form-group" style="margin-bottom:8px"><label>Rolle</label>
          <select onchange="setRole('${esc(upn)}',this.value)">${roleOpts(p.role)}</select></div>
        <div style="font-size:.78rem;font-weight:600;color:#374151;margin-bottom:4px">Freigegebene Werke</div>
        <div class="werk-picker">${werkChecks(upn,p.werke)}</div>
      </div>`).join('') || '<p class="su-empty" style="color:#9ca3af;font-size:.85rem">Noch keine Nutzer freigegeben.</p>';
    adminBlock = `
      <hr class="modal-hr">
      <div class="settings-section-title">🔐 Zugriffsverwaltung (Werk + Rolle)</div>
      <p class="field-sub" style="margin-bottom:8px">Lege pro Person Rolle und freigegebene Werke fest. Als Administrator hast du immer Zugriff auf alle Werke.</p>
      <div class="su-add" style="display:flex;gap:8px;margin-bottom:12px">
        <input type="email" id="access-new-upn" placeholder="name@dihag.com" class="su-input">
        <button class="btn btn-sm btn-primary" onclick="addAccessUser()">+ Hinzufügen</button>
      </div>
      ${userBlocks}`;
  }
  $id('settings-body').innerHTML = meBlock + rechteBlock + adminBlock;
  $id('settings-modal').classList.remove('hidden');
}
function closeSettings(){ $id('settings-modal').classList.add('hidden'); }
function addAccessUser(){
  const inp=$id('access-new-upn'); const upn=(inp?.value||'').trim().toLowerCase();
  if(!upn||!/@/.test(upn)){ toast('Bitte gültige E-Mail/UPN angeben.','error'); return; }
  if(!accessUsers[upn]) accessUsers[upn]={ role:'verantwortlicher', werke:[] };
  saveAccessConfig(); openSettings();
}
function removeAccessUser(upn){ delete accessUsers[(upn||'').trim().toLowerCase()]; saveAccessConfig(); openSettings(); }
function setRole(upn,role){ upn=(upn||'').toLowerCase(); if(accessUsers[upn]){ accessUsers[upn].role=role; saveAccessConfig(); } }
function toggleWerk(upn,werk,on){ upn=(upn||'').toLowerCase(); if(!accessUsers[upn]) return; const s=new Set(accessUsers[upn].werke); on?s.add(werk):s.delete(werk); accessUsers[upn].werke=[...s]; saveAccessConfig(); }

function showPrivacyNotice(){
  $id('modal-title').textContent = 'Datenschutzhinweis für Besucher';
  $id('modal-body').innerHTML = `<div style="font-size:.86rem;line-height:1.55;color:#374151">
    <p><b>Verantwortlicher:</b> DIHAG Deutschland GmbH / betreffende Konzerngesellschaft am Standort.</p>
    <p><b>Zweck &amp; Rechtsgrundlage:</b> Zutrittskontrolle, Werks- und Besuchersicherheit sowie Nachweis der
       Sicherheitsunterweisung (Art. 6 (1) f DSGVO – berechtigtes Interesse an der Betriebssicherheit; Art. 6 (1) c
       i.V.m. Arbeitsschutzrecht für den Unterweisungsnachweis).</p>
    <p><b>Kategorien:</b> Name, Firma, Funktion, ggf. Kontaktdaten und Kfz-Kennzeichen, Besuchszeitpunkte, ausgegebene
       PSA sowie die Bestätigung/Unterschrift zur Sicherheitsunterweisung.</p>
    <p><b>Empfänger:</b> nur berechtigte Beschäftigte des jeweiligen Werks (Empfang/Wachschutz, Sekretariat, verantwortlicher
       Bereich). Speicherung in Microsoft 365 (Auftragsverarbeiter mit AV-Vertrag), Serverstandort EU.</p>
    <p><b>Speicherdauer:</b> automatische Löschung nach <b>90 Tagen</b>.</p>
    <p><b>Ihre Rechte:</b> Auskunft, Berichtigung, Löschung, Einschränkung sowie Widerspruch. Kontakt:
       Datenschutzbeauftragte(r) des Verantwortlichen.</p>
    <p class="dsgvo-hint">Entwurf – vor Produktivnutzung durch DSB/Rechtsabteilung und Betriebsrat freigeben.</p>
  </div>`;
  $id('modal-footer').innerHTML = `<button class="btn btn-ghost" onclick="window.print()">Drucken</button><button class="btn btn-primary" onclick="closeModal()">Schließen</button>`;
  $id('modal-overlay').classList.remove('hidden');
}
function closeModal(){ $id('modal-overlay').classList.add('hidden'); }

// ── BOOT ────────────────────────────────────────────────────────────────────
document.querySelectorAll('.nav-item').forEach(n=>n.addEventListener('click', e=>{ e.preventDefault(); navigate(n.dataset.view); }));
$id('menu-toggle')?.addEventListener('click', ()=> $id('sidebar').classList.toggle('open'));

async function boot(){
  try{
    bootSub('Anmeldung läuft …');
    await initMsal();
    if(!account){ $id('boot-spinner').style.display='none'; $id('boot-btn').style.display='inline-block'; bootSub('Bitte anmelden.'); return; }
    bootSub('Lade Berechtigungen …');
    await loadMyIdentities();
    await discoverSP();
    await loadAccessConfig();

    // Kein Zugriff → Diagnose, keine Daten laden
    if(!myAccess()){
      $id('boot-spinner').style.display='none';
      bootSub('');
      const ids = [..._myIdentities()];
      const cnt = Object.keys(accessUsers).length;
      let detail, showIds=false;
      if (accessLoadState==='no-list')
        detail = `Die Liste <b>${esc(ACCESS_LIST_NAME)}</b> wurde nicht gefunden oder ist für dich nicht sichtbar. Sie muss in der Site existieren und für die App-Nutzer <b>lesbar</b> sein (siehe SETUP.md §2).`;
      else if (accessLoadState==='read-failed')
        detail = `Die Liste <b>${esc(ACCESS_LIST_NAME)}</b> ist für dich nicht lesbar (fehlende Leseberechtigung). Bitte für die App-Nutzer mindestens <b>Lesen</b> auf diese Liste vergeben.`;
      else if (accessLoadState==='parse-failed')
        detail = `Die Konfiguration konnte nicht gelesen werden (ungültiger Inhalt in <b>ConfigValue</b>). Der Administrator sollte den Eintrag unter ⚙️ neu speichern.`;
      else if (cnt===0)
        detail = `Es ist noch niemand freigeschaltet – oder das Speichern hat nicht funktioniert. Der Administrator muss dich unter ⚙️ → Zugriffsverwaltung eintragen (und die Erfolgsmeldung „gespeichert ✓" abwarten).`;
      else { detail = `Du bist unter keiner deiner Anmelde-Adressen eingetragen. Der Administrator sollte dich unter <b>einer dieser</b> Adressen freischalten:`; showIds=true; }
      bootErr('');
      const eb = $id('boot-err');
      eb.innerHTML =
        `Kein Zugriff freigeschaltet – angemeldet als <b>${esc(myUPN())}</b>.<br><br>${detail}` +
        (showIds ? `<br><span style="font-size:.85em">${ids.map(esc).join('<br>')}</span>` : '') +
        `<br><br><span style="font-size:.8em;color:#9ca3af">Diagnose: Konfig-Liste ${accessListId?'gefunden':'nicht gefunden'} · Status „${esc(accessLoadState)}" · ${cnt} Nutzer freigeschaltet.</span>`;
      $id('boot-btn').style.display='inline-block'; $id('boot-btn').textContent='Erneut versuchen'; $id('boot-btn').onclick=()=>location.reload();
      return;
    }

    // Header
    const name = account.name||myUPN();
    $id('hdr-name').textContent = name;
    $id('hdr-mail').textContent = myUPN();
    $id('hdr-av').textContent = (name[0]||'?').toUpperCase();
    applyNavVisibility();

    $id('boot').style.display='none';
    $id('app').style.display='';
    await loadItems();
    navigate(canSeeDashboard() ? 'dashboard' : 'new');
  }catch(e){
    console.error(e);
    $id('boot-spinner').style.display='none';
    const host = SP_SITE.split(':/')[0], path = SP_SITE.split(':/')[1];
    const siteUrl = `https://${host}/${path}`;
    const eb = $id('boot-err');
    if (e.message==='SITE_NOT_FOUND'){
      eb.innerHTML = `SharePoint-Site nicht gefunden:<br><b>${esc(siteUrl)}</b><br><br>`+
        `Bitte die Site anlegen (SETUP.md §2) – oder, falls sie unter einem anderen Namen existiert, `+
        `<code>SP_SITE</code> in <code>app.js</code> auf den echten Pfad anpassen.`;
    } else if (e.message==='LIST_NOT_FOUND'){
      eb.innerHTML = `Site gefunden, aber die Liste <b>„${esc(SP_LIST)}"</b> fehlt.<br>`+
        `Bitte die Liste in der Site anlegen (SETUP.md §2).`;
    } else {
      eb.textContent = 'Fehler beim Start: '+e.message;
    }
    const btn = $id('boot-btn');
    btn.style.display='inline-block'; btn.textContent='Erneut versuchen'; btn.onclick=()=>location.reload();
  }
}
boot();
