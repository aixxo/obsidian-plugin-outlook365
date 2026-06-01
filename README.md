# Outlook Calendar to Notes

Ein Obsidian-Plugin, das Termine aus deinem Outlook 365-Kalender abruft und daraus Notizen im Vault erstellt – wahlweise nach eigenen Vorlagen.

---

## Voraussetzungen

- Obsidian **Desktop** (Windows / macOS / Linux)
- Ein Microsoft 365-Konto mit Zugriff auf Outlook-Kalender
- Eine eigene **Azure AD App-Registrierung** (kostenlos, ohne Admin-Rechte möglich)

---

## Einrichtung

### 1. Azure AD App registrieren

1. Öffne [portal.azure.com](https://portal.azure.com) und gehe zu **Microsoft Entra ID → App-Registrierungen → Neue Registrierung**.
2. Name: beliebig, z. B. `Obsidian Outlook Plugin`.
3. Unterstützte Kontotypen: **Konten in einem beliebigen Organisationsverzeichnis und persönliche Microsoft-Konten** (oder nur deine Organisation).
4. Redirect-URI: Plattform **„Mobile- und Desktopanwendungen"** → URI: `obsidian://outlook-calendar-notes`
5. Registrieren.
6. Unter **API-Berechtigungen → Berechtigung hinzufügen → Microsoft Graph → Delegierte Berechtigungen**: `Calendars.Read` hinzufügen.
7. Die **Application (Client) ID** aus der Übersicht kopieren.

### 2. Plugin konfigurieren

**Einstellungen → Outlook Calendar to Notes:**

| Feld | Inhalt |
|---|---|
| Client-ID | Die Application (Client) ID aus Schritt 1.7 |
| Tenant-ID | Deine Tenant-ID oder `common` für persönliche Konten |
| Vorlagen-Ordner | Pfad im Vault, der Vorlagen-Dateien enthält (Standard: `Templates`) |
| Ausgabe-Ordner | Pfad, in dem Notizen gespeichert werden (Standard: `Meetings`) |
| Standard-Zeitraum | Vorausgewählter Zeitraum beim Öffnen des Termin-Dialogs |

### 3. Anmelden

Klicke in den Einstellungen auf **„Jetzt anmelden"**. Ein Browser-Fenster öffnet sich für den Microsoft-Login. Nach erfolgreichem Login kehrt Obsidian automatisch zurück.

---

## Nutzung

1. Klicke auf das Kalender-Symbol in der linken Symbolleiste **oder** öffne die Befehlspalette (`Ctrl/Cmd + P`) und suche nach **„Termin aus Outlook in Notiz umwandeln"**.
2. Im Dialog erscheinen deine Termine, nach Tagen gruppiert.
3. Wähle einen Zeitraum und eine Vorlage.
4. Wähle einen oder mehrere Termine per Checkbox aus.
5. Klicke **„Notizen erstellen"**.

Die erstellten Notizen landen im konfigurierten Ausgabe-Ordner. Bereits vorhandene Dateien werden übersprungen (kein Überschreiben).

---

## Vorlagen

### Vorlagen-Dateien erstellen

Lege `.md`-Dateien im Vorlagen-Ordner an (Standard: `Templates/`). Das Plugin listet automatisch alle Markdown-Dateien in diesem Ordner auf.

**Beispiel-Struktur:**

```
Templates/
  Meeting-Standard.md
  1-on-1.md
  Workshop.md
```

### Platzhalter

Platzhalter werden beim Erstellen der Notiz durch echte Termindaten ersetzt. Syntax: `{{PlatzhalterName}}`

| Platzhalter | Beschreibung | Beispiel |
|---|---|---|
| `{{Titel}}` | Betreff des Termins | `Wöchentliches Teammeeting` |
| `{{Datum}}` | Datum im Format TT.MM.JJJJ | `02.06.2026` |
| `{{Startzeit}}` | Startzeit HH:MM | `09:00` |
| `{{Endzeit}}` | Endzeit HH:MM | `09:30` |
| `{{Ort}}` | Ort/Raum (oder `–` wenn leer) | `Konferenzraum A` |
| `{{Beschreibung}}` | Kurzbeschreibung/Vorschau des Termintexts | `Agenda: Status-Update…` |
| `{{Teilnehmer}}` | Kommagetrennte Liste der Teilnehmer | `Max Mustermann, Anna Schmidt` |
| `{{Organizer}}` | Name des Organisators | `Oliver Kuhl` |
| `{{OnlineMeetingUrl}}` | Teams/Zoom-Link (oder `–` wenn kein Online-Meeting) | `https://teams.microsoft.com/…` |

> **Hinweis:** Unbekannte Platzhalter (z. B. Tippfehler) werden durch einen leeren String ersetzt.  
> Die eingebaute Standard-Vorlage wird verwendet, wenn im Dialog keine Vorlage ausgewählt ist.

### Beispiel-Vorlage

```markdown
# {{Titel}}

**Datum:** {{Datum}}
**Zeit:** {{Startzeit}} – {{Endzeit}}
**Ort:** {{Ort}}
**Organizer:** {{Organizer}}
**Teilnehmer:** {{Teilnehmer}}
**Online-Meeting:** {{OnlineMeetingUrl}}

---

## Agenda / Notizen

{{Beschreibung}}

---

## Aktionspunkte

- [ ] 

---

## Entscheidungen

- 
```

### Dateiname der erstellten Notiz

Der Dateiname wird automatisch generiert:

```
{Terminbetreff (bereinigt)}_{JJJJ-MM-TT}.md
```

Beispiel: `Wöchentliches Teammeeting_2026-06-02.md`

Sonderzeichen wie `/`, `\`, `:`, `*`, `?`, `"`, `<`, `>`, `|` werden aus dem Betreff entfernt.

---

## Befehle

| Befehl | Beschreibung |
|---|---|
| `Termin aus Outlook in Notiz umwandeln` | Öffnet den Termin-Dialog |
| `Outlook Calendar: Debug-Info in Console ausgeben` | Gibt Auth-Status und API-Antwort in die Entwickler-Konsole aus (Cmd+Opt+I / Ctrl+Shift+I) |

---

## Datenschutz

- Das Plugin speichert nur das OAuth-Token lokal in den Obsidian-Plugin-Daten (`data.json` im Plugin-Ordner).
- Es werden keine Termindaten an externe Server übertragen. Alle Anfragen gehen direkt an `graph.microsoft.com`.
- Es werden keine Metriken oder Telemetrie gesammelt.
