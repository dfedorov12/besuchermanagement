// ── Aufbewahrung / automatische Löschung ─────────────────────────────────────
// Löscht Besucher-Datensätze, die älter als RETENTION_DAYS (Standard 90) sind.
// Läuft app-only (Client-Credentials) via GitHub Actions – siehe retention.yml.
// Benötigt GitHub-Secrets: TENANT_ID, CLIENT_ID, CLIENT_SECRET
// sowie die Graph-*Application*-Berechtigung Sites.ReadWrite.All (oder Sites.Selected
// auf diese eine Site) mit erteilter Administratorzustimmung.

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SP_SITE = process.env.SP_SITE || 'dihag.sharepoint.com:/sites/IT';
const SP_LIST = process.env.SP_LIST || 'Besucheranmeldung';
const RETENTION_DAYS = parseInt(process.env.RETENTION_DAYS || '90', 10);
const API = 'https://graph.microsoft.com/v1.0';

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
  console.error('Fehlende Secrets: TENANT_ID / CLIENT_ID / CLIENT_SECRET');
  process.exit(1);
}

async function getToken() {
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });
  const r = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body
  });
  if (!r.ok) throw new Error('Token: ' + r.status + ' ' + await r.text());
  return (await r.json()).access_token;
}

async function gGet(token, url) {
  const r = await fetch(url.startsWith('http') ? url : API + url, { headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok) throw new Error(r.status + ' ' + await r.text());
  return r.json();
}
async function gDelete(token, url) {
  const r = await fetch(API + url, { method: 'DELETE', headers: { Authorization: 'Bearer ' + token } });
  if (!r.ok && r.status !== 404) throw new Error(r.status + ' ' + await r.text());
}

(async () => {
  const token = await getToken();
  const site = await gGet(token, `/sites/${SP_SITE}`);
  const list = await gGet(token, `/sites/${site.id}/lists/${encodeURIComponent(SP_LIST)}?$select=id`);
  const cutoff = Date.now() - RETENTION_DAYS * 86400000;

  let url = `/sites/${site.id}/lists/${list.id}/items?$select=id,createdDateTime&$top=200`;
  let scanned = 0, deleted = 0;
  while (url) {
    const page = await gGet(token, url);
    for (const it of page.value || []) {
      scanned++;
      if (new Date(it.createdDateTime).getTime() < cutoff) {
        await gDelete(token, `/sites/${site.id}/lists/${list.id}/items/${it.id}`);
        deleted++;
      }
    }
    url = page['@odata.nextLink'] || null;
  }
  console.log(`Aufbewahrung ${RETENTION_DAYS} Tage · geprüft: ${scanned} · gelöscht: ${deleted}`);
})().catch(e => { console.error(e); process.exit(1); });
