# Quick Start Guide - easyEXCH ProxyAddress Manager

## 🚀 Schnellstart in 3 Schritten

### Schritt 1: Voraussetzungen prüfen ✓

```powershell
# PowerShell als Administrator öffnen

# 1. PowerShell Version prüfen (mindestens 5.1)
$PSVersionTable.PSVersion

# 2. Active Directory Modul prüfen
Get-Module -ListAvailable ActiveDirectory

# Falls nicht vorhanden, installieren:
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

### Schritt 2: Anwendung starten 🖥️

```powershell
# 1. In das Verzeichnis wechseln
cd C:\Path\To\easyEXCH-ProxyMailAddresses

# 2. Falls erforderlich: Ausführungsrichtlinie anpassen
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# 3. Anwendung starten
.\easyEXCH-ProxyMailAddresses.ps1
```

### Schritt 3: Benutzer verwalten 👤

**3.1 Benutzer suchen**
```
1. Suchfeld: Namen eingeben (z.B. "Müller")
2. Enter drücken oder "Search" klicken
3. Benutzer aus Liste auswählen
```

**3.2 ProxyAddress hinzufügen**
```
1. "Add Address" klicken
2. Typ auswählen (z.B. "smtp (Secondary)")
3. Adresse eingeben (z.B. "alias@domain.com")
4. "Save" klicken
```

**3.3 Primäre Adresse ändern**
```
1. Gewünschte Adresse in Liste auswählen
2. "Set Primary" klicken
3. Mit "Yes" bestätigen
```

**3.4 ProxyAddress entfernen**
```
1. Zu entfernende Adresse auswählen
2. "Remove" klicken
3. Mit "Yes" bestätigen
```

## 📝 Erste Schritte - Praxisbeispiel

### Beispiel: Alias für Benutzer hinzufügen

**Ausgangssituation:**
- Benutzer: Max Müller
- Aktuelle E-Mail: max.mueller@contoso.com
- Benötigt: Zusätzlichen Alias m.mueller@contoso.com

**Vorgehensweise:**

1. **Anwendung starten**
   ```powershell
   .\easyEXCH-ProxyMailAddresses.ps1
   ```

2. **Benutzer suchen**
   - Suchfeld: "Müller" eingeben
   - Enter drücken
   - "Max Müller" aus Liste wählen

3. **Aktuellen Status prüfen**
   - ProxyAddresses-Liste zeigt:
     - `SMTP:max.mueller@contoso.com` (Primary)

4. **Neuen Alias hinzufügen**
   - "Add Address" klicken
   - Typ: "smtp (Secondary)" auswählen
   - Adresse: "m.mueller@contoso.com" eingeben
   - "Save" klicken

5. **Ergebnis prüfen**
   - ProxyAddresses-Liste zeigt nun:
     - `SMTP:max.mueller@contoso.com` (Primary)
     - `smtp:m.mueller@contoso.com` (Secondary) ✓

✅ **Fertig!** Max Müller erhält jetzt E-Mails an beide Adressen.

## 🎯 Häufige Aufgaben

### Aufgabe 1: MS365 Migration vorbereiten

```
Ziel: Lokale Adresse durch MS365-Adresse ersetzen

Schritte:
1. MS365-Adresse als sekundär hinzufügen
2. Einige Tage testen
3. MS365-Adresse als primär setzen
4. Alte lokale Adresse als Alias behalten (optional)

Details siehe: EXAMPLES.md - Szenario 1
```

### Aufgabe 2: Service-Account mit mehreren Kontaktpunkten

```
Ziel: Ein Postfach für mehrere E-Mail-Adressen

Schritte:
1. Primäre Adresse festlegen (z.B. SMTP:support@domain.com)
2. Weitere Aliase hinzufügen:
   - smtp:help@domain.com
   - smtp:service@domain.com
   - smtp:helpdesk@domain.com

Ergebnis: Alle E-Mails landen im selben Postfach
```

### Aufgabe 3: Nach Abteilungswechsel E-Mail ändern

```
Ziel: Von sales@domain.com zu marketing@domain.com wechseln

Schritte:
1. Neue Adresse als sekundär hinzufügen
2. Übergangsphase einplanen (z.B. 2 Wochen)
3. Neue Adresse als primär setzen
4. Alte Adresse später entfernen (optional)

Details siehe: EXAMPLES.md - Szenario 2
```

## ⚠️ Wichtige Hinweise

### ✓ Do's
- Änderungen in Testumgebung zuerst testen
- Screenshots von wichtigen Konfigurationen machen
- Bei primärer SMTP-Änderung: Übergangsphase einplanen
- Nach Änderungen: ADSync-Status prüfen
- Refresh-Button verwenden für aktuelle Daten

### ✗ Don'ts
- Primäre SMTP-Adresse nicht ohne Plan entfernen
- Keine ungültigen E-Mail-Formate verwenden
- Nicht direkt in Produktion ohne Test arbeiten
- ADSync nicht ignorieren (kann bis zu 30 Min. dauern)

## 🆘 Häufige Probleme

### Problem: "Active Directory module is not installed"
```powershell
# Lösung: RSAT installieren
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0

# PowerShell neu starten
```

### Problem: "Execution Policy" Fehler
```powershell
# Lösung: Richtlinie anpassen
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Oder einmalig:
PowerShell.exe -ExecutionPolicy Bypass -File ".\easyEXCH-ProxyMailAddresses.ps1"
```

### Problem: "Access Denied"
```
Lösung: 
- Prüfen Sie Ihre AD-Berechtigungen
- Kontaktieren Sie Ihren Domain-Administrator
- Benötigt: Lese- und Schreibrechte auf ProxyAddresses-Attribut
```

### Problem: GUI wird nicht angezeigt
```
Lösung:
- .NET Framework 4.5+ installiert?
- Als anderer Benutzer anmelden
- Windows Updates installieren
```

## 📚 Weitere Ressourcen

- **Detaillierte Installation:** [INSTALLATION.md](INSTALLATION.md)
- **Umfangreiche Beispiele:** [EXAMPLES.md](EXAMPLES.md)
- **Vollständige Dokumentation:** [README.md](README.md)
- **Probleme melden:** [GitHub Issues](https://github.com/PS-easyIT/easyEXCH-ProxyMailAddresses/issues)

## 🎓 Video-Tutorials (geplant)

- [ ] Grundlagen: Erste Schritte
- [ ] MS365 Migration durchführen
- [ ] Mehrere Aliase verwalten
- [ ] Troubleshooting häufiger Probleme

## 💡 Tipps für Fortgeschrittene

### Tipp 1: PowerShell-Befehle für Batch-Operationen
```powershell
# Für viele Benutzer: Siehe EXAMPLES.md
# Für Automatisierung: PowerShell-Scripts verwenden
```

### Tipp 2: Desktop-Verknüpfung erstellen
```
Rechtsklick Desktop → Neu → Verknüpfung
Ziel: powershell.exe -ExecutionPolicy Bypass -File "C:\Pfad\zur\Datei.ps1"
Name: easyEXCH ProxyAddress Manager
```

### Tipp 3: Regelmäßige Wartung
```
- Alle 6 Monate: Ungenutzte Aliase prüfen
- Vor großen Änderungen: Backup erstellen
- Nach Updates: Funktionalität testen
```

## 📞 Support

Bei Fragen:
1. README.md und Dokumentation lesen
2. EXAMPLES.md für Ihr Szenario prüfen
3. GitHub Issues durchsuchen
4. Neues Issue erstellen mit:
   - Beschreibung des Problems
   - Fehlermeldung (falls vorhanden)
   - Schritte zur Reproduktion
   - Umgebungsdetails (Windows, PowerShell Version)

---

**Viel Erfolg mit easyEXCH ProxyAddress Manager! 🚀**
