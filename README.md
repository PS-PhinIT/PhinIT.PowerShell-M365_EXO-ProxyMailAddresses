# easyEXCH ProxyAddress Manager

![Version](https://img.shields.io/badge/version-1.0.3-blue.svg)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)
![Language](https://img.shields.io/badge/language-DE%20%7C%20EN-orange.svg)

**Tool zur Verwaltung von ProxyMailAddresses für Microsoft 365 ADSync**

Entwickelt von **Andreas Hepp** | [www.phinit.de](https://www.phinit.de) | Version 1.0.3 (17.10.2026)

---

## 📋 Inhaltsverzeichnis

- [Übersicht](#-übersicht)
- [Features](#-features)
- [Projektstruktur](#-projektstruktur)
- [Installation](#-installation)
- [Verwendung](#-verwendung)
- [MS365 ADSync Validierung](#-ms365-adsync-validierung)
- [Logging](#-logging)
- [EXE-Konvertierung](#-exe-konvertierung)
- [Systemanforderungen](#-systemanforderungen)
- [Bekannte Einschränkungen](#-bekannte-einschränkungen)
- [Troubleshooting](#-troubleshooting)
- [Changelog](#-changelog)
- [Dokumentation](#-dokumentation)
- [Lizenz](#-lizenz)

---

## 🎯 Übersicht

**easyEXCH ProxyAddress Manager** ist ein PowerShell-basiertes Tool mit moderner WPF-GUI zur Verwaltung von ProxyMailAddresses in Active Directory. Es wurde speziell für Umgebungen entwickelt, in denen kein lokaler Exchange Server mehr vorhanden ist, aber MS365 ADSync verwendet wird.

### Hauptmerkmale

- ✅ **Moderne WPF-GUI** - Benutzerfreundliche Oberfläche im easyEXCH-Design
- ✅ **Mehrsprachig** - Deutsch und Englisch mit Live-Umschaltung (🌐 DE/EN)
- ✅ **Live-Suche** - Automatische Suche ab 3 Zeichen
- ✅ **MS365-Validierung** - Prüft ProxyAddresses auf ADSync-Kompatibilität
- ✅ **Vollständiges Logging** - Alle Änderungen werden protokolliert (Haupt-Log + Debug-Log)
- ✅ **Export & History** - CSV-Export und Änderungsverlauf pro Benutzer
- ✅ **Statistik-Dashboard** - Detaillierte Übersicht über alle Benutzer und ProxyAddresses
- ✅ **Filter-Funktion** - Benutzer nach Kompatibilität filtern
- ✅ **EXE-Konvertierung** - PS2EXE-kompatibel für einfache Verteilung
- ✅ **Legacy Exchange Support** - Berücksichtigt alte Exchange-Attribute

---

## 🚀 Features

### Benutzer-Verwaltung

- **Alle Benutzer laden** - Lädt alle aktiven AD-Benutzer
- **Intelligente Suche** - Sucht in Vorname, Nachname, DisplayName, E-Mail, SAM-Account und UPN
- **Live-Suche** - Automatische Suche ab 3 Zeichen mit 500ms Verzögerung
- **Tabellarische Darstellung** - Übersichtliche Anzeige mit Vorname, Nachname, DisplayName und primärer E-Mail
- **Filter-Funktion** - Benutzer nach Kompatibilität filtern (Alle, Kompatible, Warnungen, Fehler)
- **Export-Funktion** - Alle ProxyAddresses als CSV exportieren
- **History-Funktion** - Änderungsverlauf pro Benutzer anzeigen
- **Statistik-Dashboard** - Übersicht über alle geladenen Benutzer und ProxyAddresses

### ProxyAddress-Verwaltung

#### ➕ Hinzufügen
- Dialog zur Eingabe neuer ProxyAddresses
- Typ-Auswahl: SMTP: (primär), smtp: (sekundär), SIP:, X500:
- X500: Eingabe des vollständigen Exchange-DN (wird mit `legacyExchangeDN` vorbelegt)
- Format-Prüfung vor dem Schreiben (RFC-konforme E-Mail-Syntax, X500-DN, Steuerzeichen, max. 1123 Zeichen)
- Enter = Hinzufügen, Esc = Abbrechen
- Eindeutigkeitsprüfung im gesamten AD (Global Catalog) gegen `proxyAddresses`, `mail` und `userPrincipalName` – verhindert ADSync-Fehler `AttributeValueMustBeUnique`
- Automatische Konvertierung alter primärer Adressen zu sekundären, Mail-Attribut wird bei neuer primärer Adresse aktualisiert

#### ✏️ Bearbeiten
- Bestehende ProxyAddresses ändern
- Atomare Operation: Ersetzen in einem einzigen AD-Schreibvorgang (kein Teilzustand bei Fehlern)
- Änderung der Groß-/Kleinschreibung (z.B. `smtp:` → `SMTP:`) möglich

#### 🗑️ Löschen
- Löschen mit Sicherheitsabfrage
- Warnung bei primären SMTP-Adressen

#### ⭐ Als Primär setzen
- Sekundäre SMTP-Adressen zu primären machen
- Automatische Konvertierung der alten primären zu sekundär
- Aktualisierung des Mail-Attributs

#### 🔄 Aktualisieren
- Lädt aktuelle ProxyAddresses vom AD neu

### MS365 ADSync Validierung

#### Pro Benutzer
- Automatische Validierung bei Auswahl
- Anzeige von Fehlern und Warnungen
- Farbcodierte Status-Anzeige (✅ Grün = OK, ❌ Rot = Fehler)

#### Bulk-Prüfung
- **ADSync-Prüfung** für alle geladenen Benutzer
- Detailliertes Ergebnis-Fenster mit:
  - Zusammenfassung (Kompatibel/Fehler/Warnungen)
  - Tabelle mit allen Benutzern
  - Auflistung aller Fehler und Warnungen

### Validierungs-Kriterien

✅ **Geprüft wird**:
- Vorhandensein von ProxyAddresses
- Primäre SMTP-Adresse vorhanden (SMTP: Großbuchstaben)
- Keine mehrfachen primären SMTP-Adressen
- Keine doppelten ProxyAddresses (case-insensitive)
- Keine ungültigen SMTP-Adressen
- Mail-Attribut stimmt mit primärer SMTP überein
- E-Mail-Format (RFC 5322: Local-Part max. 64 Zeichen, keine doppelten/führenden Punkte, Domain-Labels max. 63 Zeichen, IDN-TLDs `xn--`)
- Länge (Adresse max. 254 Zeichen)
- Legacy Exchange Attribute (msExchRecipientTypeDetails, legacyExchangeDN)
- MailNickname vorhanden
- TargetAddress für Remote Mailboxes

⚠️ **Legacy Exchange Unterstützung**:
- Erkennt alte Exchange-Attribute
- Prüft X500-Adressen für Migration
- Warnt bei fehlenden Legacy-Konvertierungen

---

## 📁 Projektstruktur

```
.
├── src/
│   └── easyEXCH-PMA_V1.0.3.ps1   # Anwendung
├── docs/
│   ├── AUDIT.md                  # Audit-Bericht 1.0.2
│   ├── LOGGING.md                # Logging-Dokumentation
│   └── screenshots/              # Screenshots der aktuellen Version
├── CHANGELOG.md                  # Versionshistorie
├── CONTRIBUTING.md               # Mitwirken & Release-Checkliste
├── LICENSE                       # MIT-Lizenz
├── README.md
└── SECURITY.md                   # Sicherheitsrichtlinie
```

Ältere Versionen sind über die Git-Historie und den [Changelog](CHANGELOG.md) nachvollziehbar. Kompilierte EXE-Dateien werden nicht im Repository abgelegt (siehe [EXE-Konvertierung](#-exe-konvertierung)).

Zur Laufzeit legt das Script neben sich die Ordner `Logs\`, `Exports\` und `History\` an (siehe [Logging](#-logging)). Diese enthalten personenbezogene AD-Daten und sind per `.gitignore` ausgeschlossen.

---

## 💾 Installation

### Voraussetzungen

- Windows 10/11 oder Windows Server 2016+
- PowerShell 5.1
- Active Directory PowerShell Module
- Berechtigungen zum Ändern von AD-Benutzern

### Schnellstart

1. **Repository klonen**:
   ```powershell
   git clone https://github.com/PS-PhinIT/PhinIT.PowerShell-M365_EXO-ProxyMailAddresses.git
   cd PhinIT.PowerShell-M365_EXO-ProxyMailAddresses
   ```

2. **Script ausführen** (Windows PowerShell 5.1):
   ```powershell
   powershell.exe -STA -File .\src\easyEXCH-PMA_V1.0.3.ps1
   ```

   > Hinweis: Bei aus dem Internet geladenen Dateien ggf. `Unblock-File .\src\easyEXCH-PMA_V1.0.3.ps1` ausführen bzw. die Execution Policy beachten.

3. **Optional: Als EXE kompilieren** (siehe [EXE-Konvertierung](#-exe-konvertierung))

---

## 🎮 Verwendung

### Benutzer suchen

1. **Alle Benutzer laden**:
   - Klick auf "👥 Alle Benutzer laden"

2. **Gezielt suchen**:
   - Mindestens 3 Zeichen eingeben
   - Automatische Live-Suche nach 500ms
   - Sucht in: Vorname, Nachname, DisplayName, E-Mail, SAM-Account, UPN

3. **Benutzer auswählen**:
   - Klick auf Benutzer in der Tabelle
   - ProxyAddresses werden automatisch geladen
   - Validierung wird durchgeführt

### ProxyAddress hinzufügen

1. Benutzer auswählen
2. Klick auf **➕** Button
3. Typ auswählen (SMTP:, smtp:, SIP:, X500:)
4. Adresse eingeben (z.B. `max.mustermann@newdomain.com`)
5. Klick auf "Hinzufügen"

**Tipp**: 
- `SMTP:` (Großbuchstaben) = Primäre Adresse
- `smtp:` (Kleinbuchstaben) = Sekundäre Adresse

### ProxyAddress bearbeiten

1. Benutzer auswählen
2. ProxyAddress in der Liste auswählen
3. Klick auf **✏️** Button
4. Neue Adresse eingeben (mit Typ, z.B. `smtp:neue@domain.com`)
5. Klick auf "Speichern"

### ProxyAddress löschen

1. Benutzer auswählen
2. ProxyAddress in der Liste auswählen
3. Klick auf **🗑️** Button
4. Bestätigen

⚠️ **Warnung**: Beim Löschen der primären SMTP-Adresse erscheint eine zusätzliche Warnung!

### Als Primär setzen

1. Benutzer auswählen
2. **Sekundäre** SMTP-Adresse auswählen (smtp:)
3. Klick auf **⭐** Button
4. Bestätigen

**Automatisch**:
- Alte primäre wird zu sekundärer (SMTP: → smtp:)
- Ausgewählte wird zu primärer (smtp: → SMTP:)
- Mail-Attribut wird aktualisiert

### ADSync-Prüfung

1. Benutzer laden (alle oder suchen)
2. Klick auf **🔍 ADSync prüfen**
3. Ergebnis-Fenster wird angezeigt mit:
   - Zusammenfassung
   - Detaillierte Tabelle
   - Alle Fehler und Warnungen

### Filter verwenden

1. Benutzer laden
2. Filter-Dropdown auswählen:
   - 🔍 Alle anzeigen
   - ✅ Nur kompatible
   - ⚠️ Mit Warnungen
   - ❌ Mit Fehlern
3. Liste wird automatisch gefiltert

### ProxyAddresses exportieren

1. Benutzer laden
2. Klick auf **📤 Export CSV**
3. CSV-Datei wird im Exports-Ordner erstellt (alle geladenen Benutzer, unabhängig vom aktiven Filter)
4. Enthält: DisplayName, SamAccountName, Mail, ProxyAddress, Typ, IsPrimary

### History anzeigen

1. Benutzer auswählen
2. Klick auf **📜 History**
3. Zeigt alle Änderungen für diesen Benutzer:
   - Zeitpunkt
   - Aktion (Hinzugefügt, Gelöscht, Bearbeitet)
   - ProxyAddress
   - Ausgeführt von

### Statistik anzeigen

1. Benutzer laden
2. Klick auf **📊 Statistik**
3. Dashboard zeigt:
   - Benutzer-Statistiken (Gesamt, mit/ohne ProxyAddresses)
   - MS365-Kompatibilität (Kompatibel, Warnungen, Fehler)
   - ProxyAddress-Statistiken (Gesamt, Durchschnitt, Typen)
   - Prozentuale Verteilung

---

## 🔍 MS365 ADSync Validierung

### Validierungs-Status

| Status | Bedeutung | Aktion |
|--------|-----------|--------|
| ✅ OK | ProxyAddress ist MS365-kompatibel | Keine Aktion nötig |
| ⚠️ Legacy | X500-Adresse von altem Exchange | Normal für Migration |
| ❌ Ungültig | Format-Fehler (z.B. ungültige E-Mail-Syntax) | Korrektur erforderlich |

### Häufige Fehler

#### ❌ Keine primäre SMTP-Adresse
**Problem**: Benutzer hat keine `SMTP:` Adresse (Großbuchstaben)

**Lösung**: 
1. Sekundäre smtp: Adresse auswählen
2. Klick auf **⭐ Als Primär**

#### ❌ Mehrere primäre SMTP-Adressen
**Problem**: Mehrere `SMTP:` Adressen vorhanden

**Lösung**:
1. Falsche primäre Adresse auswählen
2. Klick auf **✏️ Bearbeiten**
3. Zu `smtp:` ändern (Kleinbuchstaben)

#### ⚠️ Mail-Attribut stimmt nicht überein
**Problem**: Mail-Attribut ≠ primäre SMTP-Adresse

**Lösung**:
1. Korrekte Adresse als primär setzen
2. Mail-Attribut wird automatisch aktualisiert

#### ⚠️ Legacy Exchange DN fehlt als X500
**Problem**: `legacyExchangeDN` nicht als X500 ProxyAddress vorhanden

**Lösung**:
1. Klick auf **➕ Hinzufügen**
2. Typ: `X500:`
3. Adresse: wird automatisch mit dem Wert aus legacyExchangeDN vorbelegt

---

## 📝 Logging

### Ordner-Struktur

```
Script-Verzeichnis\
├── Logs\
│   └── USERNAME\
│       ├── USERNAME_2025-12-08.log          # Haupt-Log
│       └── DEBUG_USERNAME_2025-12-08.log    # Debug-Log
├── Exports\
│   └── ProxyAddresses_Export_2025-12-08_14-30-00.csv
└── History\
    ├── user1_2025-12-08_14-25-30.csv
    └── user2_2025-12-08_15-10-45.csv
```

### Log-Dateien

#### Haupt-Log (`USERNAME_DATUM.log`)
- Session-Start und -Ende
- Alle Benutzer-Änderungen
- AD-Abfragen
- Fehler und Warnungen
- Erfolgreiche Operationen

#### Debug-Log (`DEBUG_USERNAME_DATUM.log`)
- Alle Meldungen inklusive DEBUG-Level
- Detaillierte Kontext-Informationen
- Thread-IDs
- Performance-Daten

### Log-Levels

| Level | Beschreibung | Haupt-Log | Debug-Log |
|-------|--------------|-----------|-----------|
| INFO | Allgemeine Informationen | ✅ | ✅ |
| SUCCESS | Erfolgreiche Operationen | ✅ | ✅ |
| WARNING | Warnungen | ✅ | ✅ |
| ERROR | Fehler | ✅ | ✅ |
| ACTION | Benutzer-Änderungen | ✅ | ✅ |
| DEBUG | Debug-Informationen | ❌ | ✅ |

### Beispiel: Benutzer-Änderung

```
╔═══════════════════════════════════════════════════════════════════════════╗
║                        BENUTZER-ÄNDERUNG PROTOKOLLIERT                     ║
╚═══════════════════════════════════════════════════════════════════════════╝
Zeitpunkt: 2025-12-08 20:15:32
Ausgeführt von: administrator@DOMAIN
Computer: DC01

AKTION: ProxyAddress HINZUGEFÜGT
ZIEL-BENUTZER: Max Mustermann (mmustermann)
E-Mail: max.mustermann@domain.com
UPN: mmustermann@domain.com
DN: CN=Max Mustermann,OU=Users,DC=domain,DC=com
ProxyAddress: smtp:max.mustermann@newdomain.com
Details: ProxyAddress wurde erfolgreich hinzugefügt
═══════════════════════════════════════════════════════════════════════════
```

### Konfiguration

```powershell
# Im Abschnitt "Logging-Konfiguration" des Scripts:
$script:MaxLogAgeDays = 90          # Logs älter als 90 Tage werden gelöscht
$script:EnableDebugLogging = $true  # Debug-Log aktivieren/deaktivieren
```

### Speicherorte

**Logs**:
- Standard: `<Script-Verzeichnis>\Logs\<USERNAME>\`
- Fallback (Verzeichnis nicht beschreibbar): `%TEMP%\easyEXCH_PMA_Logs\<USERNAME>\`

**Exports**:
- Standard: `<Script-Verzeichnis>\Exports\`
- Fallback: `%LOCALAPPDATA%\easyEXCH_PMA\Exports\`

**History**:
- Standard: `<Script-Verzeichnis>\History\`
- Fallback: `%LOCALAPPDATA%\easyEXCH_PMA\History\`
- Automatischer Backup vor jeder Änderung (schlägt er fehl, wird vor der Änderung nachgefragt)
- Format: `SAMACCOUNTNAME_TIMESTAMP.csv`

Zeilenumbrüche in Log-Meldungen werden neutralisiert (Schutz vor Log-Injection). Details: [docs/LOGGING.md](docs/LOGGING.md).

---

## 🔧 EXE-Konvertierung

Das Script ist vollständig **PS2EXE-kompatibel** und kann als eigenständige EXE-Datei kompiliert werden.

### Voraussetzungen

```powershell
# PS2EXE Modul installieren
Install-Module PS2EXE -Scope CurrentUser
```

### Empfohlene Kompilierung

```powershell
# -iconFile ist optional (Icon-Datei ist nicht im Repository enthalten).
# Hinweis: Nach einem Backtick (`) darf kein Kommentar stehen, sonst bricht die Zeilenfortsetzung.
ps2exe .\src\easyEXCH-PMA_V1.0.3.ps1 .\easyEXCH-PMA_V1.0.3.exe `
    -noConsole `
    -STA `
    -x64 `
    -iconFile .\icon.ico `
    -title "easyEXCH ProxyAddress Manager" `
    -description "ProxyAddress Manager für MS365 ADSync" `
    -company "PhinIT.DE" `
    -product "easyEXCH ProxyAddress Manager" `
    -copyright "© 2025 Andreas Hepp - www.phinit.de" `
    -version "1.0.3" `
    -DPIAware `
    -winFormsDPIAware
```

### Parameter-Erklärung

| Parameter | Beschreibung |
|-----------|--------------|
| `-noConsole` | Keine Console-Fenster (GUI-Modus) |
| `-STA` | Single Thread Apartment (erforderlich für GUI-Dialoge) |
| `-x64` | 64-bit Executable |
| `-iconFile` | Icon für die EXE-Datei |
| `-title` | Titel in Eigenschaften |
| `-version` | Versions-Nummer |

### PS2EXE-Kompatibilität

✅ **Implementiert**:
- Keine Console-Ausgaben
- Progress Bars deaktiviert (`$ProgressPreference = 'SilentlyContinue'`)
- Visual Styles aktiviert
- Pfad-Erkennung für Script & EXE
- Logging funktioniert in beiden Modi

Die erstellte EXE wird nicht eingecheckt, sondern bei Bedarf als Release-Asset veröffentlicht.

---

## 💻 Systemanforderungen

### Minimum

- **OS**: Windows 10 / Windows Server 2016
- **PowerShell**: 5.1
- **RAM**: 512 MB
- **Disk**: 50 MB

### Empfohlen

- **OS**: Windows 11 / Windows Server 2022
- **PowerShell**: 5.1
- **RAM**: 2 GB
- **Disk**: 100 MB (inkl. Logs)

### Erforderliche Module

- **ActiveDirectory** (automatisch auf Domain Controllern)

Installation auf Clients:
```powershell
# RSAT Tools installieren
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

### Berechtigungen

- **Lesen**: Alle Benutzer-Attribute
- **Schreiben**: ProxyAddresses, Mail

---

## ⚠️ Bekannte Einschränkungen

### PowerShell Core
- ❌ **Nicht unterstützt**: PowerShell 6.x / 7.x
- ✅ **Nur**: PowerShell 5.1

### Active Directory
- Funktioniert nur mit lokalem Active Directory
- Keine direkte Microsoft Entra ID Unterstützung

### Exchange Online
- Keine direkte Exchange Online Verbindung
- Nur AD-Verwaltung (für ADSync)

### Bedienung
- AD-Abfragen laufen im GUI-Thread – bei sehr großen Verzeichnissen reagiert das Fenster während „Alle Benutzer laden“ verzögert (Wartecursor)
- Beim Löschen der primären SMTP-Adresse bleibt das `mail`-Attribut unverändert (es erscheint eine Warnung)

### Bulk-Operationen
- Keine Massen-Änderungen an ProxyAddresses
- Benutzer müssen einzeln bearbeitet werden
- Bulk-Export und Bulk-Validierung sind verfügbar

---

## 🔧 Troubleshooting

### Problem: "Get-ADUser: Command not found"

**Lösung**: Active Directory Modul installieren
```powershell
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

### Problem: "Fehler beim Laden der GUI"

**Lösung**: 
1. PowerShell 5.1 verwenden (nicht Core)
2. .NET Framework 4.7+ installieren
3. Script als Administrator ausführen

### Problem: Logs werden nicht erstellt

**Lösung**:
1. Schreibrechte im Script-Verzeichnis prüfen
2. Logs werden automatisch in `%TEMP%\easyEXCH_PMA_Logs\` erstellt (Fallback, auch wenn der Ordner existiert, aber nicht beschreibbar ist)
3. Prüfen: `$script:LogFile` Variable im Debug-Modus

### Problem: Export/History-Ordner fehlen

**Lösung**:
1. Ordner werden beim Start automatisch neben dem Script erstellt
2. Ist das Script-Verzeichnis nicht beschreibbar (z.B. `C:\Program Files`), werden `%LOCALAPPDATA%\easyEXCH_PMA\Exports\` und `...\History\` verwendet (Hinweis im Log)
3. Schreibrechte im Script-Verzeichnis prüfen

### Problem: Filter zeigt keine Benutzer

**Lösung**:
1. Zuerst Benutzer laden ("Alle Benutzer laden" oder Suche)
2. Filter funktioniert nur mit geladenen Benutzern
3. Filter zurücksetzen auf "Alle anzeigen"

### Problem: Statistik zeigt "Keine Daten"

**Lösung**:
1. Zuerst Benutzer laden
2. Mindestens 1 Benutzer muss geladen sein
3. Statistik funktioniert nur mit geladenen Benutzern

### Problem: Suche findet keine Benutzer

**Lösung**:
1. Mindestens 3 Zeichen eingeben
2. Domain Controller Erreichbarkeit prüfen
3. Berechtigungen prüfen
4. Debug-Log prüfen

### Problem: ProxyAddress kann nicht geändert werden

**Lösung**:
1. Berechtigungen prüfen (Schreibrechte auf ProxyAddresses)
2. Benutzer ist nicht gesperrt
3. AD-Replikation abwarten

---

## 📋 Changelog

Alle Änderungen je Version (1.0.3, 1.0.2, 1.0.1, 1.0.0, 0.0.3, 0.0.2, 0.0.1) stehen im **[CHANGELOG.md](CHANGELOG.md)**.

**Neu in 1.0.3** (Bugfix):
- 🐛 Laden/Suchen von Benutzern funktioniert auch in AD-Umgebungen ohne Exchange-Schemaerweiterung (Fehler „One or more properties are invalid. (Parameter 'msExchRecipientTypeDetails')“)

**Neu in 1.0.2** (Audit-Release 2, Details in [docs/AUDIT.md](docs/AUDIT.md)):
- 🔒 Schutz vor Log-Injection, strengere Eingabeprüfung, Eindeutigkeit auch gegen `mail`/`userPrincipalName`, abgesichertes History-Backup
- 🐛 RFC-konforme E-Mail-Prüfung, keine doppelten primären Adressen mehr, korrekte Fallbacks für Logs/Exports/History
- ⚡ Schnellere Validierung, Statistik und Benutzeraktualisierung
- 🌐 Vollständige DE/EN-Lokalisierung aller Dialoge und Meldungen
- 📁 Neue Projektstruktur (`src/`, `docs/`), alte Versionen, Screenshots und EXE entfernt

---

## 📚 Dokumentation

| Datei | Inhalt |
|-------|--------|
| [CHANGELOG.md](CHANGELOG.md) | Versionshistorie |
| [docs/AUDIT.md](docs/AUDIT.md) | Audit-Bericht (Befunde und Status) |
| [docs/LOGGING.md](docs/LOGGING.md) | Logging-System im Detail |
| [SECURITY.md](SECURITY.md) | Sicherheitsrichtlinie und Meldeweg |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Mitwirken und Release-Checkliste |
| [LICENSE](LICENSE) | MIT-Lizenz |

---

## 📄 Lizenz

Dieses Projekt ist unter der [MIT-Lizenz](LICENSE) lizenziert – © 2025 Andreas Hepp - www.phinit.de

---

## 🙏 Danksagungen

- **PS2EXE** - Markus Scholtes für das großartige PowerShell zu EXE Konvertierungs-Tool
- **Microsoft** - Für PowerShell, WPF und Active Directory Module
- **Community** - Für Feedback und Testing

---

## 📊 Statistiken

- **Zeilen Code**: ~3300 Zeilen PowerShell
- **Funktionen**: 39 Funktionen
- **Sprachen**: 2 (Deutsch, Englisch)
- **Übersetzungen**: 162 UI-Strings pro Sprache
- **Features**: 15+ Hauptfunktionen

---

## 📚 Weitere Ressourcen

- [PS2EXE GitHub](https://github.com/MScholtes/PS2EXE)
- [Microsoft Learn - ActiveDirectory-Modul](https://learn.microsoft.com/powershell/module/activedirectory/)
- [Microsoft Learn - Microsoft Entra Connect](https://learn.microsoft.com/entra/identity/hybrid/connect/whatis-azure-ad-connect)
- [Microsoft Learn - Doppelte Attribute bei der Synchronisierung](https://learn.microsoft.com/entra/identity/hybrid/connect/how-to-connect-syncservice-duplicate-attribute-resiliency)

---


<div align="center">

**Entwickelt mit ❤️ von Andreas Hepp**

[www.phinit.de](https://www.phinit.de)

</div>
