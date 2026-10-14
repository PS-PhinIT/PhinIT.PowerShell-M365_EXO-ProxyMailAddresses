# Mitwirken am easyEXCH ProxyAddress Manager

Danke für dein Interesse! Fehlermeldungen, Verbesserungsvorschläge und Pull Requests sind willkommen.

## Repository-Struktur

```
.
├── src/
│   └── easyEXCH-PMA_V<Version>.ps1   # Anwendung (einzige Script-Datei)
├── docs/
│   ├── AUDIT.md                      # Audit-Bericht (Befunde und Status)
│   ├── LOGGING.md                    # Logging-Dokumentation
│   └── screenshots/                  # Screenshots der aktuellen Version
├── CHANGELOG.md
├── CONTRIBUTING.md
├── LICENSE
├── README.md
└── SECURITY.md
```

- Es gibt **genau eine** Script-Datei in `src/`. Ältere Versionen werden nicht im Repository abgelegt, sondern sind über die Git-Historie bzw. GitHub Releases erreichbar.
- Kompilierte EXE-Dateien werden **nicht** eingecheckt, sondern bei Bedarf per PS2EXE erstellt und als Release-Asset veröffentlicht.
- Screenshots gehören nach `docs/screenshots/` und müssen die aktuelle Version zeigen; veraltete Screenshots werden entfernt. Keine echten Benutzerdaten abbilden.
- Laufzeitordner (`Logs/`, `Exports/`, `History/`) enthalten personenbezogene Daten und dürfen nie eingecheckt werden (siehe `.gitignore`).

## Fehler melden / Vorschläge

Bitte ein GitHub-Issue mit folgenden Angaben anlegen:
- Version (Footer der GUI bzw. `.VERSION` im Script)
- Windows- und PowerShell-Version (`$PSVersionTable`)
- Schritte zur Reproduktion und erwartetes Verhalten
- relevante Auszüge aus dem Debug-Log – **anonymisiert**

Sicherheitslücken bitte gemäß [SECURITY.md](SECURITY.md) melden.

## Entwicklungsrichtlinien

- Zielplattform: **Windows PowerShell 5.1** (WPF, ActiveDirectory-Modul). PowerShell 7 wird nicht unterstützt.
- Script-Datei als **UTF-8 mit BOM** speichern (Umlaute/Emojis unter PowerShell 5.1).
- Alle GUI-Texte über `Get-LocalizedString` – neue Schlüssel immer in **beiden** Sprachen (`DE` und `EN`) mit identischen Platzhaltern (`{0}`, `{1}`, …) anlegen.
- AD-Änderungen atomar über `Set-ADUser -Replace` / `-Clear` und immer mit vorherigem History-Backup (`Invoke-PreChangeBackup`).
- Benutzereingaben in LDAP-Filtern mit `ConvertTo-LdapFilterValue`, Export-Werte mit `ConvertTo-SafeCsvValue` behandeln.
- Jede Änderung an AD-Objekten mit `Write-UserChangeLog` protokollieren.
- PS2EXE-Kompatibilität erhalten: keine Konsolenausgaben (`Write-Host`), keine Progress-Bars.

## Prüfung vor dem Pull Request

```powershell
# Syntax prüfen
$errors = $null
[void][System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path .\src\easyEXCH-PMA_V*.ps1), [ref]$null, [ref]$errors)
$errors

# Statische Analyse
Install-Module PSScriptAnalyzer -Scope CurrentUser
Invoke-ScriptAnalyzer -Path .\src\ -Recurse -Severity Error, Warning
```

Bekannte, bewusst akzeptierte Analyzer-Hinweise sind in [docs/AUDIT.md](docs/AUDIT.md) dokumentiert. Zusätzlich die betroffenen Funktionen manuell in einer Testumgebung (Test-OU) prüfen.

## Release-Checkliste

1. Version in `.VERSION` (Script-Header) und `$script:AppVersion` erhöhen – Footer und Fenstertitel übernehmen den Wert automatisch.
2. Script-Datei umbenennen: `git mv src/easyEXCH-PMA_V<alt>.ps1 src/easyEXCH-PMA_V<neu>.ps1`.
3. `CHANGELOG.md`: neuen Abschnitt `## [<Version>] – <JJJJ-MM-TT>` anlegen (es gibt keinen „Unreleased“-Abschnitt) und Vergleichslink ergänzen.
4. `README.md`: Versions-Badge, Pfade und PS2EXE-Beispiel (`-version`) anpassen.
5. Screenshots in `docs/screenshots/` aktualisieren oder entfernen.
6. Git-Tag `v<Version>` setzen und optional die EXE als Release-Asset hochladen.
