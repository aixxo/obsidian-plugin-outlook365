

# **Obsidian Plugin: Outlook 365 Kalender-Integration**
**Ziel:** Automatisierte Erstellung von Obsidian-Notizen aus Outlook-Terminen (mit Vorlagen) – **ohne Admin-Rechte im Mandanten**, falls möglich.

---

## **1. Projektübersicht**
### **Zweck**
- Nutzer können **Outlook 365-Termine** (Titel, Beschreibung, Uhrzeit, Teilnehmer, Ort) in **Obsidian-Notizen** umwandeln.
- **Vorlagenbasiert**: Nutzer wählt eine Notizvorlage (z. B. `Meeting-{{Datum}}.md` mit Platzhaltern).
- **Sicherheit**: Nur Zugriff auf **eigene Termine** des Nutzers (keine Admin-Rechte).

### **Zielgruppe**
- Obsidian-Nutzer mit **Outlook 365** (Microsoft 365/Azure AD).
- Keine IT-Administratoren erforderlich.

---

## **2. Technische Architektur**
### **Komponenten**
| Komponente          | Technologie/Tool                     | Verantwortung                          |
|---------------------|--------------------------------------|----------------------------------------|
| **Frontend**        | Obsidian Plugin (TypeScript)         | UI, Nutzerinteraktion                  |
| **Authentifizierung** | Microsoft Identity Platform (OAuth 2.0) | Sichere Anmeldung ohne Admin-Rechte |
| **API-Zugriff**     | Microsoft Graph API (`/me/calendar`) | Termine auslesen                      |
| **Notizenerstellung** | Obsidian API (`app.vault.create`)  | Notizen aus Vorlagen generieren       |
| **Datenfluss**      | `Outlook 365 → Graph API → Plugin → Obsidian` | |

---

## **3. Sicherheitskonzept**
### **Authentifizierung (OAuth 2.0)**
- **Flow:** **Authorization Code Flow mit PKCE** (für Single-Page-Apps/Plugins).
  - **Warum?** Keine Admin-Rechte nötig, nur Nutzer-Zustimmung für **eigene Daten**.
  - **Scopes (Minimalberechtigungen):**
    - `Calendars.Read` (nur Lesen der eigenen Termine)
    - `offline_access` (für Token-Refresh)
  - **Keine Admin-Rechte:** Nutzer autorisiert nur **seine eigenen Daten**.

- **Token-Management:**
  - Tokens **lokal im Plugin speichern** (verschlüsselt, z. B. mit `obsidian-plugin` Storage).
  - **Refresh-Token** nutzen, um Zugriff ohne erneute Authentifizierung zu verlängern.

- **Fallback für Mandanten mit strengen Richtlinien:**
  - Falls **delegierte Berechtigungen** (`/me/...`) blockiert sind:
    - **Application Permissions** (`Calendars.Read.All`) + **Admin-Zustimmung** (nur als letzte Option).
    - **Sicherer Absatz:** Nutzer muss **explizit zustimmen**, dass das Plugin **nur seine Termine** liest.

---

## **4. Implementierungsschritte**
### **Phase 1: Plugin-Grundgerüst (Obsidian)**
1. **Projekt initialisieren:**
   ```bash
   npm create obsidian-plugin@latest outlook-calendar-notes
   cd outlook-calendar-notes
   npm install @microsoft/microsoft-graph-client @azure/msal-browser
   ```
2. **Manifest (`manifest.json`):**
   ```json
   {
     "id": "outlook-calendar-notes",
     "name": "Outlook Calendar to Note",
     "version": "1.0.0",
     "minAppVersion": "1.0.0",
     "author": "Oliver Kuhl",
     "description": "Erstellt Obsidian-Notizen aus Outlook-Terminen.",
     "isDesktopOnly": false
   }
   ```

---

### **Phase 2: Microsoft Graph API-Integration**
1. **App-Registrierung in Azure AD:**
   - [Azure Portal → App Registrierungen](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/NewApplication)
   - **Redirect URI:** `obsidian://outlook-calendar-notes` (für Desktop-Apps).
   - **API-Berechtigungen:** `Calendars.Read` (delegiert).

2. **Authentifizierung im Plugin:**
   ```typescript
   import { PublicClientApplication } from "@azure/msal-browser";

   const msalConfig = {
     auth: {
       clientId: "DEINE_CLIENT_ID",
       authority: "https://login.microsoftonline.com/common",
       redirectUri: "obsidian://outlook-calendar-notes",
     },
   };
   const msalInstance = new PublicClientApplication(msalConfig);
   ```
   - **Login-Button in Obsidian:**
     ```typescript
     async function signIn() {
       const loginRequest = { scopes: ["Calendars.Read"] };
       const authResult = await msalInstance.loginPopup(loginRequest);
       return authResult.accessToken;
     }
     ```

3. **Termine abrufen (Graph API):**
   ```typescript
   import { Client } from "@microsoft/microsoft-graph-client";

   async function fetchEvents(accessToken: string) {
     const client = Client.init({ authProvider: (done) => done(null, accessToken) });
     const events = await client
       .api("/me/calendar/events")
       .select("subject,start,end,body,location,attendees")
       .orderby("start/dateTime desc")
       .get();
     return events.value;
   }
   ```

---

### **Phase 3: Notizenerstellung in Obsidian**
1. **Vorlagen-System:**
   - Nutzer kann **eigene Vorlagen** in einem Ordner (z. B. `Templates/Meeting.md`) ablegen.
   - Platzhalter: `{{Titel}}`, `{{Datum}}`, `{{Beschreibung}}`, `{{Teilnehmer}}`, etc.
   - **Beispiel-Vorlage:**
     ```markdown
     # {{Titel}}
     **Datum:** {{Datum}}
     **Uhrzeit:** {{Startzeit}} – {{Endzeit}}
     **Ort:** {{Ort}}
     **Teilnehmer:** {{Teilnehmer}}

     ---
     ## Notizen
     - [ ]
     ```

2. **Notiz generieren:**
   ```typescript
   async function createNoteFromEvent(event: any, template: string) {
     const noteContent = template
       .replace(/{{Titel}}/g, event.subject)
       .replace(/{{Datum}}/g, new Date(event.start.dateTime).toLocaleDateString("de-DE"))
       .replace(/{{Startzeit}}/g, new Date(event.start.dateTime).toLocaleTimeString("de-DE"))
       .replace(/{{Endzeit}}/g, new Date(event.end.dateTime).toLocaleTimeString("de-DE"))
       .replace(/{{Ort}}/g, event.location?.displayName || "–")
       .replace(/{{Beschreibung}}/g, event.body?.content || "–")
       .replace(/{{Teilnehmer}}/g, event.attendees?.map((a: any) => a.emailAddress.name).join(", ") || "–");

     const fileName = `Meetings/${event.subject.replace(/[^a-z0-9]/gi, "_")}_${new Date(event.start.dateTime).toISOString().split("T")[0]}.md`;
     await this.app.vault.create(fileName, noteContent);
   }
   ```

3. **UI-Integration:**
   - **Command Palette:** Befehl `"Outlook-Termin in Notiz umwandeln"`.
   - **Modal-Dialog:**
     - Liste der **letzten 50 Termine** (aus Graph API).
     - Nutzer wählt Termin + Vorlage aus.
     - Button: **"Notiz erstellen"**.

---

### **Phase 4: Fehlerbehandlung & Edge Cases**
| Szenario                          | Lösung                                                                 |
|-----------------------------------|------------------------------------------------------------------------|
| **Kein Internet**                 | Offline-Modus: Zuletzt heruntergeladene Termine anzeigen.            |
| **Token abgelaufen**              | Automatischer Refresh mit `msalInstance.acquireTokenSilent()`.      |
| **Admin-Zustimmung erforderlich** | Nutzer wird aufgefordert, Admin zu kontaktieren (Fallback-Option).  |
| **Keine Berechtigungen**          | Klare Fehlermeldung: "Berechtigung für Kalenderzugriff fehlt."      |
| **Doppelte Notizen vermeiden**     | Prüfen, ob Notiz bereits existiert (z. B. über Dateinamen).          |

---

## **5. Testing & Deployment**
### **Testplan**
1. **Lokale Tests:**
   - Mock-Daten für Graph API (z. B. mit [MSW](https://mswjs.io/)).
   - Testen mit **echtem Outlook-Kalender** (manuelle Authentifizierung).
2. **Sicherheitstests:**
   - Prüfen, ob **nur eigene Termine** gelesen werden.
   - Token-Speicherung verschlüsseln (`obsidian-plugin` Storage).
3. **Nutzerfeedback:**
   - Beta-Test mit Obsidian-Community (z. B. über GitHub Issues).

### **Deployment**
1. **Obsidian Community Plugins:**
   - Plugin auf [GitHub](https://github.com/) veröffentlichen.
   - Einreichung für den **Obsidian Plugin Store** (über Pull Request in [obsidian-releases](https://github.com/obsidianmd/obsidian-releases)).
2. **Dokumentation:**
   - `README.md` mit:
     - Installationsanleitung.
     - Screenshots der UI.
     - Beispiel-Vorlagen.
     - Troubleshooting (z. B. "Admin-Zustimmung erforderlich").

---

## **6. Beispiel-Code-Struktur**
```
outlook-calendar-notes/
├── src/
│   ├── main.ts                # Plugin-Einstiegspunkt
│   ├── auth.ts               # MSAL-Authentifizierung
│   ├── graphApi.ts           # Graph API-Aufrufe
│   ├── noteGenerator.ts     # Notizenerstellung
│   ├── ui/
│   │   ├── modal.ts          # Modal-Dialog für Terminauswahl
│   │   └── settings.ts       # Plugin-Einstellungen (z. B. Vorlagenordner)
│   └── utils/
│       └── templates.ts      # Vorlagen-Parser
├── templates/
│   └── default.md            # Standard-Vorlage
├── manifest.json
├── package.json
└── README.md
```

---
## **7. Offene Fragen & Risiken**
| Frage/Risiko                          | Lösungsidee                                                                 |
|---------------------------------------|-----------------------------------------------------------------------------|
| **Wird `/me/calendar` in allen Mandanten unterstützt?** | Falls nein: Nutzer muss Admin um `Calendars.Read.All` + Zustimmung bitten. |
| **Wie mit wiederkehrenden Terminen umgehen?** | Option: Nur **einzelne Instanzen** anzeigen oder alle expandieren.        |
| **Performance bei vielen Terminen?**   | Pagination in Graph API (`$top=50`, `$skip=...`).                        |
| **Datenschutz (DSGVO)?**              | Nutzer muss **explizit zustimmen** (Opt-in). Keine Daten an Dritte.      |

---
## **8. Nächste Schritte für GitHub Copilot**
1. **Code-Generierung:**
   - Copilot kann **Boilerplate-Code** für:
     - MSAL-Authentifizierung.
     - Graph API-Aufrufe.
     - Obsidian-Plugin-Struktur generieren.
   - **Prompt-Beispiel:**
     > *"Erstelle eine TypeScript-Klasse für die Authentifizierung mit MSAL in einem Obsidian-Plugin. Nutze den Authorization Code Flow mit PKCE und speichere Tokens im Obsidian-Storage. Scopes: Calendars.Read."*

2. **Dokumentation:**
   - Copilot kann **README.md** oder **API-Dokumentation** basierend auf diesem Plan erstellen.

3. **Testing:**
   - Copilot kann **Jest-Tests** für die Vorlagen-Parser oder Graph API-Aufrufe vorschlagen.

---
## **9. Ressourcen & Links**
- [Microsoft Graph API – Kalender](https://learn.microsoft.com/de-de/graph/api/resources/calendar?view=graph-rest-1.0)
- [MSAL.js für Browser](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-browser)
- [Obsidian Plugin Entwicklung](https://docs.obsidian.md/Plugins/Getting+started)
- [OAuth 2.0 für Single-Page-Apps](https://datatracker.ietf.org/doc/html/rfc8252)
