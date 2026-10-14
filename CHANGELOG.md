# Changelog

Alle relevanten Änderungen am **easyEXCH ProxyAddress Manager** werden in dieser Datei dokumentiert.

Das Format orientiert sich an [Keep a Changelog](https://keepachangelog.com/de/1.1.0/), die Versionierung an [Semantic Versioning](https://semver.org/lang/de/).
Jeder Eintrag gehört zu einer konkreten Version mit Datum.

> **Quellen der Historie:** Die Versionen 0.0.1 bis 1.0.0 wurden am 08.12.2025 gemeinsam hochgeladen (Commit `00be449`) und aus den damals im Ordner `Versions/` abgelegten Script-Ständen rekonstruiert. Ab Version 1.0.1 basiert die Historie auf den Pull Requests des Repositorys.

---

## [1.0.2] – 2026-10-14 – Audit-Release 2

Vollständiges Audit (Sicherheit, Performance, Funktionen, Verweise, GUI, Befehle) inkl. Bereinigung der Repository-Struktur. Details: [docs/AUDIT.md](docs/AUDIT.md).

### Sicherheit
- Log-Injection verhindert: Zeilenumbrüche (CR/LF) in Log-Meldungen und Kontextwerten werden neutralisiert
- Steuerzeichen in neuen ProxyAddresses werden abgelehnt
- Maximale Länge einer ProxyAddress (1123 Zeichen, AD-Schema) wird geprüft
- CSV-Injection-Schutz zusätzlich für die History-Spalten `User` und `ExecutedBy`
- Eindeutigkeitsprüfung neuer SMTP-Adressen prüft jetzt auch die Attribute `mail` und `userPrincipalName` anderer Objekte (vermeidet ADSync-Konflikte)
- Schlägt das History-Backup vor einer Änderung fehl, wird die Änderung erst nach expliziter Bestätigung durchgeführt (Protokollierung als WARNING)
- Fallback der Global-Catalog-Abfrage wird als WARNING protokolliert

### Behoben
- E-Mail-Validierung: RFC-5322-konforme Prüfung (u.a. Apostroph `o'brien@…` gültig, `a..b@…` ungültig, Local-Part ≤ 64 Zeichen, Domain-Labels ≤ 63 Zeichen, IDN-TLDs `xn--…`)
- „Als Primär setzen“ entfernt case-insensitive Duplikate der Zieladresse – es bleibt genau eine `SMTP:`-Adresse
- `Update-CurrentUser`: quadratische Laufzeit (`@($script:AllUsers).Count` im Schleifenkopf) behoben
- Validierungs-Cache wird beim Sprachwechsel geleert (Meldungen waren sonst in der alten Sprache)
- Domänen-Cache wird nach Laden/Ändern von Benutzern zurückgesetzt
- Logging-Initialisierung: nicht beschreibbarer, aber vorhandener Log-Ordner führt jetzt korrekt zum Fallback `%TEMP%\easyEXCH_PMA_Logs`
- Export/History: Fallback auf `%LOCALAPPDATA%\easyEXCH_PMA\` wenn das Script-Verzeichnis nicht beschreibbar ist (z.B. `C:\Program Files`)
- „History“-Button bezieht sich auf den im linken Bereich angezeigten Benutzer und wurde bisher beim Aufheben der Auswahl bzw. nach einer neuen Suche deaktiviert, obwohl dieser Benutzer weiterhin angezeigt wurde – er bleibt jetzt aktiv, solange ein Benutzer angezeigt wird
- Doppelte `FindName`-Aufrufe in der GUI-Initialisierung entfernt
- Footer zeigte eine fest codierte Versionsnummer – Version wird jetzt aus `$script:AppVersion` erzeugt (Footer und Fenstertitel)

### Geändert
- Vollständige Lokalisierung (DE/EN) aller Dialoge, Meldungen, Bestätigungen, Spaltenüberschriften, Statistik-, History-, Export- und ADSync-Fenster (86 neue Sprach-Schlüssel, insgesamt 162 je Sprache)
- Start-Fehlermeldungen (fehlendes AD-Modul, GUI-Ladefehler, Logging-Warnung) zweisprachig
- Hinzufügen-/Bearbeiten-Dialoge: Enter = Bestätigen, Esc = Abbrechen, Fokus und Textauswahl beim Öffnen, feste Größe
- Ergebnisfenster (ADSync, History) mit Esc schließbar
- Validierung pro Benutzer in einem Durchlauf mit `List`/`HashSet` (Performance bei vielen Adressen)
- Statistik in einem einzigen Durchlauf über alle Benutzer
- Ermittelte Proxy-Domänen werden pro Datenstand zwischengespeichert
- Mindestgröße des Hauptfensters (1100 × 650)
- PSScriptAnalyzer: Trailing Whitespace, leere `catch`-Blöcke und ungenutzte Parameter bereinigt

### Repository
- Neue Ordnerstruktur: `src/` (Script), `docs/` (Dokumentation), `docs/screenshots/` (Platzhalter für neue Screenshots)
- Script umbenannt/verschoben: `easyEXCH-PMA_V1.0.1.ps1` → `src/easyEXCH-PMA_V1.0.2.ps1`
- `README_LOGGING.md` → `docs/LOGGING.md`
- Entfernt: alte Script-Versionen (`Versions/`, V0.0.1 – V1.0.1), alle Screenshots (`Screenshots/`) und die veraltete `easyEXCH-PMA_V1.0.0.exe` (EXE wird bei Bedarf per PS2EXE erstellt)
- Neu: `CHANGELOG.md`, `LICENSE`, `SECURITY.md`, `CONTRIBUTING.md`, `docs/AUDIT.md`
- PS2EXE-Beispiel in der Dokumentation korrigiert (Kommentar nach Backtick brach die Zeilenfortsetzung)
- Veraltete `docs.microsoft.com`-Links auf `learn.microsoft.com` aktualisiert
- `.gitignore` deckt Laufzeitordner (`Logs/`, `Exports/`, `History/`) in allen Unterordnern ab

---

## [1.0.1] – 2026-10-14 – Audit-Release

Pull Request [#2](https://github.com/PS-PhinIT/PhinIT.PowerShell-M365_EXO-ProxyMailAddresses/pull/2) („Audit fixes: security, atomic AD writes, validation bugfix, performance, GUI, docs“).

### Sicherheit
- `.gitignore` schließt die Laufzeit-Ordner `Logs/`, `Exports/` und `History/` aus (personenbezogene AD-Daten)
- CSV-/Formel-Injection in Exporten verhindert (Werte mit `=`, `+`, `-`, `@` werden entschärft)
- LDAP-Filter-Escaping bei der Eindeutigkeitsprüfung
- Prüfung beim Start, ob das ActiveDirectory-Modul verfügbar ist

### Behoben
- Falsche Warnung „Mail-Attribut stimmt nicht mit primärer SMTP-Adresse überein“ bei jedem Benutzer mit genau einer primären Adresse (String-Indexierung lieferte nur das erste Zeichen)
- Änderungen (Hinzufügen, Bearbeiten, Löschen, Als Primär) erfolgen in einem atomaren AD-Schreibvorgang
- History-Backup wird vor der Änderung mit dem aktuellen AD-Stand erstellt
- Bearbeiten ohne Lösch-/Neuanlage-Kette und ohne doppelte Bestätigungsdialoge
- X500-Adressen wurden fälschlich als `X500:name@domain` erzeugt
- History zeigte Einträge anderer Benutzer mit gleichem Namenspräfix (z.B. `max` / `max_muster`)
- Filter mit 0 Treffern führte zu einem Fehler
- Domänen-Ermittlung: UPN-Suffixe berücksichtigt, echte öffentliche TLDs (z.B. `.dev`) nicht mehr gefiltert, nur SMTP-Adressen ausgewertet
- Aktualisierte Benutzer werden in Liste, Filter, Export und Statistik übernommen
- Filter bleibt nach neuem Laden/Suchen aktiv
- Zeilenumbruch im Hinweistext „Kein Benutzer ausgewählt“

### Performance
- Filter auf aktive Benutzer serverseitig statt clientseitig
- Validierungsergebnisse werden zwischengespeichert (Filter, Statistik, ADSync-Prüfung)
- Listen statt Array-`+=` bei großen Datenmengen
- Forest-Domänen werden einmal pro Sitzung ermittelt
- Ein wiederverwendeter Such-Timer statt eines neuen Timers pro Tastendruck

### Geändert
- Neue ProxyAddresses werden auf Format und Eindeutigkeit im gesamten AD (Global Catalog) geprüft
- Validierung erkennt doppelte und ungültige SMTP-Adressen
- Warnungen werden auch bei kompatiblen Benutzern angezeigt
- Validierungsstatus, Ja/Nein-Spalte und Filteranzeige lokalisiert (DE/EN)
- Wartecursor bei AD-Abfragen
- Befehle: `Refresh-CurrentUser` → `Update-CurrentUser` (genehmigtes Verb), keine Überschreibung der Automatik-Variablen `$Args`/`$sender`

---

## [1.0.0] – 2025-12-08

### Hinzugefügt
- Mehrsprachigkeit Deutsch/Englisch mit Live-Umschaltung (🌐 DE/EN im Header): GUI-Elemente, Buttons, Labels, Tooltips, Status-Meldungen und DataGrid-Spalten
- Breitere Aktions-Buttons mit Emoji + Text im Grid-Layout

### Behoben
- Array-Konvertierung und `Count` bei einzelnen Suchergebnissen
- Filter-Übergabe an `Get-ADUser`
- Parameter-Handling und Platzhalter-Ersetzung in `Get-LocalizedString`
- Footer-Anzeige bei Sprachwechsel

### Geändert
- Vollständige PS2EXE-Kompatibilität
- Erste Veröffentlichung als kompilierte EXE (`easyEXCH-PMA_V1.0.0.exe`, in 1.0.2 entfernt)

---

## [0.0.3] – 2025-12-08

### Hinzugefügt
- Intelligente Domänen-Filterung im Hinzufügen-Dialog: lokale/interne TLDs (`local`, `internal`, `corp`, `test`, …) werden ausgeblendet

---

## [0.0.2] – 2025-12-08

### Hinzugefügt
- Export aller ProxyAddresses als CSV (`Exports\`)
- History: automatisches Backup vor jeder Änderung und Anzeige des Änderungsverlaufs pro Benutzer (`History\`)
- Statistik-Dashboard (Benutzer, MS365-Kompatibilität, Adresstypen, Verteilung)
- Filter nach Kompatibilität (Alle, Kompatibel, Warnungen, Fehler)
- Domänen-Auswahl im Hinzufügen-Dialog (ermittelte Accepted Domains)
- Anzeigename des Benutzers im Titel des Hinzufügen-Dialogs

### Entfernt
- OU-Filter-Eingabefeld

---

## [0.0.1] – 2025-12-08

### Hinzugefügt
- Erste Version mit WPF-GUI im easyEXCH-Design
- Benutzer laden und Live-Suche (ab 3 Zeichen, 500 ms Verzögerung) in Vorname, Nachname, DisplayName, E-Mail, SAM-Account und UPN
- ProxyAddresses hinzufügen, bearbeiten, löschen und als primär setzen (inkl. Aktualisierung des `mail`-Attributs)
- MS365-ADSync-Validierung pro Benutzer und als Bulk-Prüfung
- Logging mit Haupt- und Debug-Log in benutzerspezifischen Ordnern, Log-Bereinigung und TEMP-Fallback

---

## Hinweise zu nicht übernommenen Änderungen

- Pull Request [#1](https://github.com/PS-PhinIT/PhinIT.PowerShell-M365_EXO-ProxyMailAddresses/pull/1) (08.12.2025, alternative Implementierung `easyEXCH-ProxyMailAddresses.ps1`) wurde ohne Merge geschlossen und ist daher in keiner Version enthalten.

[1.0.2]: https://github.com/PS-PhinIT/PhinIT.PowerShell-M365_EXO-ProxyMailAddresses/compare/0ffc9df...main
[1.0.1]: https://github.com/PS-PhinIT/PhinIT.PowerShell-M365_EXO-ProxyMailAddresses/pull/2
[1.0.0]: https://github.com/PS-PhinIT/PhinIT.PowerShell-M365_EXO-ProxyMailAddresses/commit/00be449
[0.0.3]: https://github.com/PS-PhinIT/PhinIT.PowerShell-M365_EXO-ProxyMailAddresses/commit/00be449
[0.0.2]: https://github.com/PS-PhinIT/PhinIT.PowerShell-M365_EXO-ProxyMailAddresses/commit/00be449
[0.0.1]: https://github.com/PS-PhinIT/PhinIT.PowerShell-M365_EXO-ProxyMailAddresses/commit/00be449
