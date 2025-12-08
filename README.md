# easyEXCH ProxyAddress Manager

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)
![Version](https://img.shields.io/badge/version-1.0.0-orange.svg)

**easyEXCH ProxyAddress Manager** ist ein PowerShell-basiertes Tool mit moderner WPF-GUI zur Verwaltung von ProxyMailAddresses in Active Directory. Es wurde speziell für Umgebungen entwickelt, in denen kein lokaler Exchange Server mehr vorhanden ist, aber MS365 ADSync verwendet wird.

## 🌟 Features

- **Moderne WPF-GUI**: Benutzerfreundliche grafische Oberfläche mit modernem Design
- **Benutzersuche**: Schnelle Suche nach AD-Benutzern über DisplayName, SamAccountName oder E-Mail
- **ProxyAddress-Verwaltung**:
  - Anzeigen aller ProxyAddresses eines Benutzers
  - Hinzufügen neuer ProxyAddresses (SMTP, smtp, SIP, X400, X500)
  - Entfernen bestehender ProxyAddresses
  - Festlegen der primären SMTP-Adresse
- **Validierung**: Automatische E-Mail-Format-Validierung
- **Sicherheit**: Warnmeldungen bei kritischen Operationen
- **Status-Updates**: Echtzeit-Statusmeldungen für alle Operationen

## 📋 Voraussetzungen

### Software-Anforderungen

- **Windows PowerShell 5.1** oder höher
- **Active Directory PowerShell Module** (RSAT-Tools)
- **.NET Framework 4.5** oder höher (für WPF)

### Installation RSAT-Tools

**Windows 10/11:**
```powershell
# Als Administrator ausführen
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

**Windows Server:**
```powershell
# Als Administrator ausführen
Install-WindowsFeature -Name RSAT-AD-PowerShell
```

### Berechtigungen

Der ausführende Benutzer benötigt:
- Leseberechtigung für Active Directory-Benutzer
- Schreibberechtigung für ProxyAddresses-Attribut der zu verwaltenden Benutzer

## 🚀 Installation

1. **Repository klonen oder herunterladen:**
   ```powershell
   git clone https://github.com/PS-easyIT/easyEXCH-ProxyMailAddresses.git
   cd easyEXCH-ProxyMailAddresses
   ```

2. **Ausführungsrichtlinie prüfen (falls erforderlich):**
   ```powershell
   # Aktuelle Richtlinie anzeigen
   Get-ExecutionPolicy
   
   # Falls erforderlich, Richtlinie für aktuellen Benutzer ändern
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. **Skript ausführen:**
   ```powershell
   .\easyEXCH-ProxyMailAddresses.ps1
   ```

## 📖 Verwendung

### Benutzeroberfläche

Die Anwendung besteht aus mehreren Hauptbereichen:

#### 1. User Search (Benutzersuche)
- Geben Sie einen Suchbegriff ein (Name, SamAccountName oder E-Mail)
- Klicken Sie auf "Search" oder drücken Sie Enter
- Die Ergebnisse werden in der Benutzerliste angezeigt (max. 50 Ergebnisse)

#### 2. Users (Benutzerliste)
- Zeigt alle gefundenen Benutzer an
- Wählen Sie einen Benutzer aus, um dessen ProxyAddresses anzuzeigen

#### 3. ProxyAddresses (ProxyAddresses-Verwaltung)
- **Liste**: Zeigt alle ProxyAddresses des ausgewählten Benutzers
  - Type: Art der Adresse (SMTP, smtp, SIP, etc.)
  - Address: Die eigentliche E-Mail-Adresse oder URI
  - Primary: Kennzeichnet die primäre Adresse (nur eine pro Typ)

- **Aktionen**:
  - **Add Address**: Neue ProxyAddress hinzufügen
  - **Set Primary**: Ausgewählte Adresse als primär festlegen
  - **Remove**: Ausgewählte Adresse entfernen
  - **Refresh**: ProxyAddresses aktualisieren

### Schritt-für-Schritt-Anleitung

#### ProxyAddress hinzufügen

1. Benutzer suchen und auswählen
2. Klick auf "Add Address"
3. Adresstyp auswählen:
   - **SMTP (Primary)**: Primäre SMTP-Adresse (nur eine pro Benutzer)
   - **smtp (Secondary)**: Sekundäre SMTP-Adresse (mehrere möglich)
   - **SIP**: Session Initiation Protocol Adresse (z.B. für Skype/Teams)
   - **X400**: X.400-Adresse (Legacy)
   - **X500**: X.500-Distinguished Name (Legacy)
4. E-Mail-Adresse oder URI eingeben
5. Klick auf "Save"

#### Primäre SMTP-Adresse ändern

1. Benutzer auswählen
2. Gewünschte sekundäre SMTP-Adresse in der Liste markieren
3. Klick auf "Set Primary"
4. Bestätigung mit "Yes"
5. Die bisherige primäre Adresse wird automatisch zur sekundären Adresse

#### ProxyAddress entfernen

1. Benutzer auswählen
2. Zu entfernende Adresse in der Liste markieren
3. Klick auf "Remove"
4. Bestätigung mit "Yes"
5. **Hinweis**: Beim Entfernen der primären SMTP-Adresse erscheint eine zusätzliche Warnung

## 🔧 Technische Details

### ProxyAddress-Format

ProxyAddresses in Active Directory folgen diesem Format: `TYPE:address`

- **SMTP** (Großbuchstaben): Primäre SMTP-Adresse (Reply-To)
- **smtp** (Kleinbuchstaben): Sekundäre SMTP-Adressen (Alias)
- **sip**: SIP-Adresse für Unified Communications
- **X400**: X.400-Adresse für Legacy-Systeme
- **X500**: X.500-Distinguished Name

### Beispiele

```
SMTP:max.mustermann@contoso.com    # Primäre E-Mail-Adresse
smtp:max@contoso.com                # Alias
smtp:m.mustermann@contoso.com       # Weiterer Alias
sip:max.mustermann@contoso.com      # SIP-Adresse
```

### Active Directory-Integration

Das Tool verwendet folgende PowerShell-Cmdlets:

- `Get-ADUser`: Benutzer suchen und laden
- `Set-ADUser`: ProxyAddresses modifizieren
  - `-Add`: ProxyAddress hinzufügen
  - `-Remove`: ProxyAddress entfernen
  - `-Replace`: ProxyAddresses ersetzen (für primäre Adresse)

## 🛡️ Sicherheit

### Warnsysteme

- **Primäre SMTP-Adresse**: Warnung beim Entfernen oder Ändern
- **Duplikate**: Prüfung auf bereits vorhandene Adressen
- **Format-Validierung**: E-Mail-Adressen werden validiert

### Best Practices

1. **Sicherung**: Erstellen Sie ein Backup vor größeren Änderungen
2. **Testen**: Testen Sie Änderungen zunächst mit Testbenutzern
3. **Dokumentation**: Dokumentieren Sie wichtige Änderungen
4. **Berechtigungen**: Vergeben Sie nur notwendige Berechtigungen

## 🔍 Fehlerbehebung

### "Active Directory module is not installed"

**Lösung**: RSAT-Tools installieren (siehe Voraussetzungen)

### "Access Denied" / Zugriff verweigert

**Lösung**: 
- Prüfen Sie Ihre AD-Berechtigungen
- Führen Sie das Skript als Benutzer mit entsprechenden Rechten aus

### Skript startet nicht

**Lösung**:
```powershell
# Ausführungsrichtlinie prüfen
Get-ExecutionPolicy

# Falls zu restriktiv
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### GUI wird nicht angezeigt

**Lösung**:
- Prüfen Sie, ob .NET Framework 4.5+ installiert ist
- Prüfen Sie, ob Windows WPF-Komponenten verfügbar sind

## 📝 Verwendungsszenarien

### Szenario 1: Migration zu MS365

Bei der Migration von lokalem Exchange zu MS365:
1. Benutzer hat noch alte lokale SMTP-Adresse als primär
2. Fügen Sie neue MS365-Adresse als sekundär hinzu
3. Setzen Sie MS365-Adresse als primär
4. ADSync synchronisiert Änderungen zu MS365

### Szenario 2: Alias-Verwaltung

Benutzer benötigt zusätzliche E-Mail-Aliase:
1. Suchen Sie den Benutzer
2. Fügen Sie weitere smtp:-Adressen hinzu
3. MS365 akzeptiert E-Mails an alle Aliase

### Szenario 3: Abteilungswechsel

Benutzer wechselt Abteilung und benötigt neue E-Mail:
1. Fügen Sie neue SMTP-Adresse als sekundär hinzu
2. Testen Sie die neue Adresse
3. Setzen Sie neue Adresse als primär
4. Alte Adresse bleibt als Alias erhalten (optional)

## 🤝 Beitragen

Beiträge sind willkommen! Bitte beachten Sie:

1. Forken Sie das Repository
2. Erstellen Sie einen Feature-Branch (`git checkout -b feature/AmazingFeature`)
3. Committen Sie Ihre Änderungen (`git commit -m 'Add some AmazingFeature'`)
4. Pushen Sie den Branch (`git push origin feature/AmazingFeature`)
5. Öffnen Sie einen Pull Request

## 📄 Lizenz

Dieses Projekt steht unter der MIT-Lizenz - siehe LICENSE-Datei für Details.

## 👥 Autoren

- **PS-easyIT** - *Initial work*

## 🙏 Danksagungen

- Microsoft Active Directory Team für das PowerShell-Modul
- WPF-Community für UI-Best-Practices
- Alle Tester und Beitragenden

## 📞 Support

Bei Fragen oder Problemen:
- Öffnen Sie ein Issue im GitHub-Repository
- Konsultieren Sie die Microsoft-Dokumentation zu ProxyAddresses

## 🔄 Changelog

### Version 1.0.0 (2025-12-08)
- Initial Release
- WPF-GUI Implementation
- Basis-Funktionalität für ProxyAddress-Verwaltung
- Benutzersuche und -auswahl
- Add/Remove/Set Primary Funktionen
- Format-Validierung
- Status- und Fehlermeldungen

## 🗺️ Roadmap

Geplante Features für zukünftige Versionen:

- [ ] Export/Import von ProxyAddresses (CSV)
- [ ] Batch-Operationen für mehrere Benutzer
- [ ] Logging-Funktionalität
- [ ] Konfigurationsdatei für Standardwerte
- [ ] Domain-Filter
- [ ] Erweiterte Suchoptionen
- [ ] ProxyAddress-Vorlagen
- [ ] Undo/Redo-Funktionalität
- [ ] Dark Mode
- [ ] Mehrsprachigkeit (EN/DE)