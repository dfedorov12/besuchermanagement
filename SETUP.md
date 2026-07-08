# DIHAG Besuchermanagement – Einrichtung

Statische SPA (GitHub Pages) · MSAL-Login · SharePoint-Liste als Backend · Zugriff nach **Werk + Rolle**.

> **Datenschutz/Legal zuerst:** Vor Produktivstart mit DSB, Rechtsabteilung und Betriebsrat (§ 87 BetrVG,
> da Gastgeber-Mitarbeiter + Zeiten erfasst werden) abstimmen. Datenschutzhinweis (in der App unter ⚙️ →
> „Datenschutzhinweis") am Empfang aushängen. Verarbeitung im VVT eintragen.

## 1. Azure AD – App „DIHAG Besuchermanagement"

- **Client-ID:** `674a4aed-2a41-4c31-9d3f-ded1a1377afa`
- **Tenant:** `fdb70646-023a-403b-a4b9-1f474a935123` (Single-Tenant)

**API-Berechtigungen → Microsoft Graph → Delegiert:** `User.Read`, `Sites.ReadWrite.All`
→ danach **Administratorzustimmung erteilen**.

**Authentifizierung → Plattform → Single-Page-Application (SPA):**
- `https://dfedorov12.github.io/besuchermanagement/`
- (optional Custom-Domain, falls genutzt)
- Kein Client-Secret für den Login (öffentliche SPA, PKCE).

> Hinweis Datensparsamkeit: Im delegierten Flow ist der Zugriff durch die **SharePoint-Berechtigungen des
> angemeldeten Nutzers** begrenzt. Die restriktive Berechtigung der Site/Liste ist daher die eigentliche Datengrenze.

## 2. SharePoint – Site + Listen

Aktuell konfiguriert: **`https://dihag.sharepoint.com/sites/IT`** (`SP_SITE = 'dihag.sharepoint.com:/sites/IT'`).
Falls der Name abweicht: `SP_SITE` in `app.js` und in den Workflow-Variablen anpassen.

> **Datenschutz-Hinweis:** Die Besucher-PII liegt damit in der IT-Site. Wer Zugriff auf diese Site hat, kann
> die Datensätze grundsätzlich sehen. Daher entweder den Site-Zugriff eng halten **oder** die Listen per
> „Listen-Berechtigungen" (Vererbung unterbrechen) auf den Besucher-Personenkreis beschränken
> (Empfang/Wachschutz, Sekretariat, verantwortliche Bereiche). Eine dedizierte Site bleibt die sauberere Variante.

**Versionsverlauf der Liste aktiviert lassen** (Prüfbarkeit).

### Liste `Besucheranmeldung`

| Spalte (Anzeige) | Typ | Hinweis |
|---|---|---|
| `Title` | Einzeltext | wird als **Besuchername** verwendet (Anzeige „Besucher") |
| `Werk` | Auswahl | DIHAG, DSO, EIS, EMH, EWA, Kein, Kernwerk, LEG, MEG, OZB, SCH, SHB, WGC, ZAI |
| `Bereich` | Einzeltext | |
| `AnsprechpartnerName` | Einzeltext | Gastgeber |
| `AnsprechpartnerTelefon` | Einzeltext | |
| `Besuchsdatum` | Datum/Uhrzeit | |
| `Ankunftszeit` | Datum/Uhrzeit | optional (geplant) |
| `Firma` | Einzeltext | |
| `Funktion` | Einzeltext | optional |
| `BesucherTelefon` | Einzeltext | optional (PII) |
| `BesucherEmail` | Einzeltext | optional (PII) |
| `Autokennzeichen` | Einzeltext | optional (PII) |
| `Besuchszweck` | Auswahl (Mehrfach) | Werksbesichtigung, Kundenbesuch, Audit, Lieferantenbesuch, Sonstiges, DIHAG |
| `PSA` | Auswahl (Mehrfach) | Schutzhelm, Schutzbrille, Eigene PSA, Warnweste, Gehörschutz |
| `SHBAkzeptiert` | Ja/Nein | |
| `Signatur` | Mehrere Textzeilen | Base64-PNG der Unterschrift |
| `Eingangszeit` | Datum/Uhrzeit | per Klick gesetzt |
| `Abgangszeit` | Datum/Uhrzeit | per Klick gesetzt (schließt Datensatz) |
| `Status` | Auswahl | Angemeldet, Eingecheckt, Geschlossen |
| `Bemerkungen` | Mehrere Textzeilen | |
| `GruppenId` | Einzeltext | verbindet Personen einer Anmeldung |

> Interne Spaltennamen sollten den Anzeigenamen entsprechen (keine Sonderzeichen/Leerzeichen). Die App
> matcht zusätzlich auf Anzeigenamen und meldet fehlende Spalten im gelben Banner oben.

### Liste `BESU_Konfiguration` (Zugriffssteuerung)

| Spalte | Typ |
|---|---|
| `Title` | Einzeltext (ein Element mit Wert `access`) |
| `ConfigValue` | Mehrere Textzeilen (hält JSON) |

Die App legt das `access`-Element selbst an; die **Liste** muss manuell existieren.
Admin (`administrator@dihag.com`) pflegt in der App unter ⚙️ pro Person **Rolle + freigegebene Werke**.

**Rollen:** `wachschutz` (nur Ein-/Abgang stempeln) · `verantwortlicher` (anlegen & bearbeiten) · `sekretariat` (voll).
Wer nicht eingetragen ist, hat **keinen** Zugriff (privacy-by-default).

## 3. GitHub Pages

Repo `dfedorov12/besuchermanagement`, Pages aus `main` / root. Dateien: `index.html`, `app.js`, `style.css`.

## 4. Aufbewahrung (Auto-Löschung 90 Tage)

`.github/workflows/retention.yml` + `scripts/retention.mjs` löschen Datensätze älter als 90 Tage.

Benötigt **app-only**-Zugang (getrennt vom SPA-Login):
1. In derselben App-Registrierung ein **Client-Secret** anlegen.
2. **Application**-Berechtigung `Sites.ReadWrite.All` (oder `Sites.Selected` nur auf diese Site) + Adminzustimmung.
3. GitHub **Secrets**: `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`.
4. GitHub **Variables**: `SP_SITE` (= `dihag.sharepoint.com:/sites/IT`), `SP_LIST` (= `Besucheranmeldung`).
5. Frist über `RETENTION_DAYS` im Workflow anpassbar.

Alternativ/ergänzend: SharePoint-**Aufbewahrungslabel** (Retention Policy) auf der Liste.

## 5. Datenschutz-Checkliste

- [ ] Rechtsgrundlage dokumentiert (Art. 6 (1) f/c DSGVO), VVT-Eintrag
- [ ] Datenschutzhinweis am Empfang ausgehängt (Art. 13)
- [ ] Betriebsrat / § 87 BetrVG geklärt
- [ ] Zugriff need-to-know (Werk + Rolle) + Site-Berechtigungen restriktiv
- [ ] Aufbewahrung 90 Tage aktiv (Job und/oder Retention-Label)
- [ ] Optional-Felder bleiben optional (Datenminimierung)
- [ ] Versionsverlauf aktiv (Prüfbarkeit)
