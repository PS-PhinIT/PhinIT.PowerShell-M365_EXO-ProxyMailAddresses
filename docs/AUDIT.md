# Audit-Bericht – easyEXCH ProxyAddress Manager 1.0.2

| | |
|---|---|
| **Datum** | 2026-10-14 |
| **Geprüfter Stand** | `easyEXCH-PMA_V1.0.1.ps1` (nach PR #2) inkl. Repository und Dokumentation |
| **Ergebnis-Version** | `src/easyEXCH-PMA_V1.0.2.ps1` |
| **Umfang** | Sicherheit, Performance, Funktionen, Verweise, GUI, Befehle, Repository/Dokumentation |

Legende: ✅ behoben · ⚠️ bewusst akzeptiert (Begründung angegeben)

---

## 1. Sicherheit

| # | Befund | Schwere | Status |
|---|--------|---------|--------|
| S1 | Log-Injection: Benutzereingaben (Suchbegriffe, Adressen) mit CR/LF konnten gefälschte Logzeilen erzeugen | Mittel | ✅ `Write-Log` ersetzt Zeilenumbrüche in Meldung und Kontextwerten |
| S2 | Neue ProxyAddresses mit Steuerzeichen wurden akzeptiert | Mittel | ✅ Ablehnung von `\p{Cc}` |
| S3 | Keine Längenprüfung gegen das AD-Schema (`proxyAddresses` max. 1123 Zeichen) | Niedrig | ✅ Prüfung in `Test-NewProxyAddressFormat` |
| S4 | Eindeutigkeitsprüfung nur gegen `proxyAddresses` – Kollisionen mit `mail`/`userPrincipalName` anderer Objekte führen in Entra Connect zu Sync-Fehlern | Mittel | ✅ LDAP-Filter `(|(proxyAddresses=…)(mail=…)(userPrincipalName=…))` |
| S5 | History-CSV: Spalten `User`/`ExecutedBy` nicht gegen CSV-Injection geschützt | Niedrig | ✅ `ConvertTo-SafeCsvValue` |
| S6 | Fehlgeschlagenes History-Backup wurde ignoriert – Änderung ohne Rücksicherung | Mittel | ✅ `Invoke-PreChangeBackup` fragt nach, Abbruch möglich, WARNING im Log |
| S7 | Fallback von Global Catalog auf Domänen-Suche nur als DEBUG protokolliert (Prüfung ist dann nicht forestweit) | Niedrig | ✅ WARNING |
| S8 | Benutzersuche mit `Get-ADUser -Filter { … -like $wildcard }` | – | ✅ geprüft: Variablenbindung, keine Filter-Injection möglich |
| S9 | Repository enthielt eine nicht nachvollziehbar gebaute Binärdatei (`easyEXCH-PMA_V1.0.0.exe`, veraltet) | Mittel | ✅ entfernt; EXE wird aus dem Quellcode per PS2EXE erstellt |
| S10 | Keine Sicherheitsrichtlinie / kein Meldeweg | Niedrig | ✅ `SECURITY.md` |
| S11 | Exports/History werden nicht automatisch bereinigt (personenbezogene Daten) | Niedrig | ⚠️ dokumentiert in `SECURITY.md`; Löschfristen sind betreiberspezifisch |

## 2. Performance

| # | Befund | Status |
|---|--------|--------|
| P1 | `Update-CurrentUser`: `@($script:AllUsers).Count` im Schleifenkopf erzeugte pro Iteration eine Array-Kopie (O(n²)), keine Abbruchbedingung | ✅ einmalige Zählung, `break` nach Treffer |
| P2 | `Test-UserMS365Compatibility`: mehrere Pipelines/`Where-Object` pro Benutzer | ✅ ein Durchlauf mit `List`/`HashSet` |
| P3 | `Show-Statistics`: mehrfache Durchläufe über alle Benutzer | ✅ ein Durchlauf |
| P4 | `Get-ADAcceptedDomains`: Auswertung aller ProxyAddresses bei jedem Öffnen des Hinzufügen-Dialogs | ✅ Cache, invalidiert bei Laden/Ändern |
| P5 | `Get-UserPrimarySmtp` per Pipeline | ✅ `foreach` |
| P6 | AD-Abfragen laufen auf dem UI-Thread (GUI blockiert während „Alle laden“) | ⚠️ Wartecursor vorhanden; Umbau auf Runspaces wäre ein größerer Eingriff mit PS2EXE-Risiko |

## 3. Funktionen

| # | Befund | Status |
|---|--------|--------|
| F1 | E-Mail-Regex lehnte gültige Adressen ab (z.B. `o'brien@…`) und akzeptierte ungültige (`a..b@…`, führender Punkt) | ✅ `Test-SmtpAddressSyntax` (RFC 5322 dot-atom, Längen, IDN-TLD `xn--`) |
| F2 | „Als Primär setzen“ konnte bei case-insensitiven Duplikaten zwei `SMTP:`-Einträge erzeugen | ✅ Deduplizierung |
| F3 | Validierungs-Cache überlebte Sprachwechsel → Meldungen in falscher Sprache | ✅ Cache wird geleert |
| F4 | Vorhandener, aber nicht beschreibbarer Log-Ordner verhinderte den TEMP-Fallback | ✅ Schreibprobe, zweistufiger Fallback |
| F5 | Exports/History scheiterten bei schreibgeschütztem Script-Verzeichnis (z.B. `C:\Program Files`) | ✅ Fallback `%LOCALAPPDATA%\easyEXCH_PMA\` |
| F6 | Tote Prüfungen („Unerlaubte Zeichen“, „Zu lang“) in `Test-ProxyAddressMS365Compatibility` | ✅ entfernt (durch Syntaxprüfung abgedeckt) |
| F7 | Löschen der primären SMTP-Adresse lässt das `mail`-Attribut unverändert | ⚠️ Benutzer wird vorab gewarnt; automatisches Ändern von `mail` wäre eine nicht gewünschte Seiteneffekt-Änderung |
| F8 | Export umfasst alle geladenen Benutzer, nicht nur die gefilterte Ansicht | ⚠️ dokumentiertes Verhalten |
| F9 | History aus dem Fallback-Ordner wird nicht mit dem Standard-Ordner zusammengeführt | ⚠️ Fallback ist ein Ausnahmefall |

## 4. GUI

| # | Befund | Status |
|---|--------|--------|
| G1 | Zahlreiche Dialoge, Bestätigungen, Fehlermeldungen, Spaltenköpfe und Fenster (Statistik, History, ADSync, Export) nur auf Deutsch | ✅ 86 neue Sprach-Schlüssel (162 je Sprache), Schlüssel- und Platzhalter-Parität geprüft |
| G2 | Start-Meldungen (AD-Modul fehlt, GUI-Ladefehler, Logging-Warnung) nur Deutsch | ✅ zweisprachig |
| G3 | Footer mit fest codierter Version | ✅ aus `$script:AppVersion`, ebenso Fenstertitel |
| G4 | Hinzufügen/Bearbeiten: kein Enter/Esc, kein Fokus, frei skalierbar | ✅ `IsDefault`/`IsCancel`, Fokus + Textauswahl, `NoResize` |
| G5 | Typ-Präfix im Hinzufügen-Dialog per String-Vergleich des angezeigten Textes ermittelt (bricht bei Lokalisierung) | ✅ Ermittlung über `SelectedIndex` |
| G6 | „History“-Button wurde beim Aufheben der Tabellenauswahl / nach neuer Suche deaktiviert, obwohl der Benutzer weiterhin angezeigt wurde (inkonsistent zu den übrigen Aktionen) | ✅ aktiv, solange ein Benutzer angezeigt wird (`$script:CurrentUser`) |
| G7 | Fenster ohne Mindestgröße – Layout brach bei kleinen Fenstern | ✅ `MinWidth=1100`, `MinHeight=650` |
| G8 | Doppelte `FindName`-Aufrufe | ✅ entfernt |
| G9 | Rückgabewerte von `Children.Add`/`MessageBox::Show` gelangten in die Pipeline | ✅ `[void]` / `Out-Null` |

## 5. Befehle / Codequalität (PSScriptAnalyzer)

| Regel | Vorher | Nachher | Status |
|-------|-------:|--------:|--------|
| PSAvoidTrailingWhitespace | 356 | 0 | ✅ |
| PSAvoidUsingEmptyCatchBlock | 3 | 0 | ✅ |
| PSReviewUnusedParameter | 1 | 0 | ✅ |
| PSUseShouldProcessForStateChangingFunctions | 10 | 10 | ⚠️ GUI-Anwendung mit eigenen Bestätigungsdialogen; `-WhatIf`/`-Confirm` werden nicht genutzt |
| PSUseSingularNouns | 8 | 8 | ⚠️ interne Funktionsnamen (z.B. `Get-ADAcceptedDomains`), keine öffentliche API |
| PSAvoidOverwritingBuiltInCmdlets (`Write-Log`) | 1 | 1 | ⚠️ `Write-Log` existiert nur in PowerShell 7 (Windows); das Script ist auf 5.1 beschränkt (`#Requires -Version 5.1`, WPF) |

Parser: 0 Fehler. Lokalisierung: identische Schlüsselmengen DE/EN, alle verwendeten Schlüssel definiert, identische Platzhalter.

## 6. Verweise / Dokumentation / Repository

| # | Befund | Status |
|---|--------|--------|
| V1 | README verwies auf Screenshot `Screenshots/GUIScreenshot_V1.0.0.jpg` (veraltet) | ✅ Screenshots entfernt, Ordner `docs/screenshots/` angelegt |
| V2 | Alte Script-Versionen (V0.0.1–V1.0.1) als Kopien in `Versions/` | ✅ entfernt (Historie in Git und `CHANGELOG.md`) |
| V3 | Kein `CHANGELOG.md` (Changelog nur im README), keine `LICENSE`-Datei, kein `SECURITY.md`/`CONTRIBUTING.md` | ✅ angelegt |
| V4 | PS2EXE-Beispiel fehlerhaft: Kommentar nach Backtick (`-iconFile .\icon.ico `   # …`) bricht die Zeilenfortsetzung | ✅ korrigiert (README, LOGGING) |
| V5 | Veraltete Links `docs.microsoft.com` / Azure AD | ✅ `learn.microsoft.com`, Microsoft Entra Connect |
| V6 | ASCII-Screenshot mit defektem Zeichen und Version 1.0.0 | ✅ entfernt |
| V7 | `.gitignore` deckte Laufzeitordner nur im Root bzw. eine Ebene tiefer ab | ✅ Muster für alle Ebenen |
| V8 | `README_LOGGING.md` im Root | ✅ nach `docs/LOGGING.md` verschoben und aktualisiert |

---

## Prüfmethodik

- Statische Analyse: PowerShell-Parser und PSScriptAnalyzer (Severity Error/Warning/Information).
- Funktionstests: Funktionsdefinitionen per AST geladen und mit gemockten AD-Cmdlets (`Get-ADUser`, `Set-ADUser`, `Get-ADObject`, `Get-ADForest`) ausgeführt – u.a. Adress-Syntax, Formatprüfung, Benutzervalidierung (DE/EN), Hinzufügen/Bearbeiten/Löschen/Als Primär setzen, Domänen-Ermittlung, CSV- und Log-Schutz.
- Manuelle Code-Durchsicht aller Event-Handler und Dialoge.
- Nicht automatisiert prüfbar: WPF-Darstellung und Verhalten gegen ein echtes Active Directory – vor dem produktiven Einsatz in einer Test-OU prüfen.
