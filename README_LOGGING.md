# easyEXCH ProxyAddress Manager - Logging System

## 📋 Übersicht

Das Tool verfügt über ein vollständiges Logging-System, das speziell für die Konvertierung zu einer EXE-Datei optimiert ist.

## 📁 Log-Struktur

```
Script-Verzeichnis\
└── Logs\
    ├── USERNAME1\
    │   ├── USERNAME1_2025-12-08.log          # Haupt-Log (alle Aktionen)
    │   ├── DEBUG_USERNAME1_2025-12-08.log    # Debug-Log (detailliert)
    │   ├── USERNAME1_2025-12-09.log
    │   └── ...
    ├── USERNAME2\
    │   ├── USERNAME2_2025-12-08.log
    │   └── ...
    └── ...
```

## 📝 Log-Dateien

### Haupt-Log (`USERNAME_DATUM.log`)
- **Zweck**: Protokollierung aller wichtigen Aktionen
- **Inhalt**:
  - Session-Start und -Ende
  - Benutzer-Änderungen (Hinzufügen/Ändern/Löschen von ProxyAddresses)
  - AD-Abfragen
  - Fehler und Warnungen
  - Erfolgreiche Operationen

### Debug-Log (`DEBUG_USERNAME_DATUM.log`)
- **Zweck**: Detaillierte Debugging-Informationen
- **Inhalt**: Alle Meldungen inklusive DEBUG-Level
- **Verwendung**: Fehlersuche und Analyse

## 🔍 Log-Levels

| Level | Beschreibung | Haupt-Log | Debug-Log |
|-------|--------------|-----------|-----------|
| **INFO** | Allgemeine Informationen | ✅ | ✅ |
| **SUCCESS** | Erfolgreiche Operationen | ✅ | ✅ |
| **WARNING** | Warnungen | ✅ | ✅ |
| **ERROR** | Fehler | ✅ | ✅ |
| **ACTION** | Benutzer-Änderungen | ✅ | ✅ |
| **DEBUG** | Debug-Informationen | ❌ | ✅ |

## 📊 Beispiel: Benutzer-Änderung

```
╔═══════════════════════════════════════════════════════════════════════════╗
║                        BENUTZER-ÄNDERUNG PROTOKOLLIERT                     ║
╚═══════════════════════════════════════════════════════════════════════════╝
Zeitpunkt: 2025-12-08 19:54:32
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

## 🔧 Konfiguration

### Log-Aufbewahrung
```powershell
$script:MaxLogAgeDays = 90  # Logs älter als 90 Tage werden automatisch gelöscht
```

### Debug-Logging aktivieren/deaktivieren
```powershell
$script:EnableDebugLogging = $true  # true = Debug-Log aktiv, false = nur Haupt-Log
```

## 🚀 EXE-Konvertierung

### PS2EXE-Kompatibilität
Das Script ist vollständig PS2EXE-kompatibel:

```powershell
# Empfohlene Kompilierung:
ps2exe .\easyEXCH-PMA_V0.0.1.ps1 .\easyEXCH-PMA.exe `
    -noConsole `
    -STA `
    -x64 `
    -iconFile .\icon.ico `
    -title "easyEXCH ProxyAddress Manager" `
    -version "0.0.1" `
    -company "PhinIT.DE" `
    -copyright "© 2025 Andreas Hepp"
```

### Wichtige Hinweise
- ✅ **Keine Console-Ausgaben** - Alle Meldungen nur in Logs
- ✅ **Progress Bars deaktiviert** - Verhindert Popup-Fenster
- ✅ **Visual Styles aktiviert** - Moderne GUI
- ✅ **Pfad-Erkennung** - Funktioniert als Script UND EXE

## 📍 Log-Speicherort

### Standard
```
<Script-Verzeichnis>\Logs\<USERNAME>\
```

### Fallback (bei Schreibfehler)
```
%TEMP%\easyEXCH_PMA_Logs\<USERNAME>\
```

## 🔒 Sicherheit

### Protokollierte Informationen
- ✅ Zeitstempel (Millisekunden-Genauigkeit)
- ✅ Ausführender Benutzer (Windows-Anmeldename)
- ✅ Computer-Name
- ✅ Domain
- ✅ Ziel-Benutzer (bei Änderungen)
- ✅ Durchgeführte Aktion
- ✅ Alte und neue Werte
- ✅ Thread-ID (für Debugging)

### NICHT protokolliert
- ❌ Passwörter
- ❌ Sensible Daten

## 📖 Log-Format

### Standard-Eintrag
```
[2025-12-08 19:54:32.123] [TID:1] [INFO] [GUI] Bereit
```

### Mit Kontext
```
[2025-12-08 19:54:32.456] [TID:1] [ACTION] [UserChange] ProxyAddress HINZUGEFÜGT für Benutzer mmustermann | Context: {TargetUser=mmustermann, ProxyAddress=smtp:max@domain.com, Action=ProxyAddress HINZUGEFÜGT}
```

## 🛠️ Troubleshooting

### Problem: Logs werden nicht erstellt
**Lösung**: 
1. Prüfen Sie Schreibrechte im Script-Verzeichnis
2. Logs werden automatisch in `%TEMP%\easyEXCH_PMA_Logs\` erstellt

### Problem: Zu viele alte Logs
**Lösung**: 
- Automatische Bereinigung nach 90 Tagen
- Manuelle Bereinigung: Logs-Ordner löschen

### Problem: Debug-Log zu groß
**Lösung**: 
```powershell
# Debug-Logging deaktivieren (Zeile 261 im Script)
$script:EnableDebugLogging = $false
```

## 📞 Support

Bei Fragen oder Problemen:
- **Autor**: Andreas Hepp
- **Website**: www.PhinIT.DE
- **Version**: 0.0.1
- **Datum**: 2025-12-08
