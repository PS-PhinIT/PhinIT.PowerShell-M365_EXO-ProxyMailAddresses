# Sicherheitsrichtlinie

## Unterstützte Versionen

Sicherheitskorrekturen werden ausschließlich für die jeweils aktuelle Version bereitgestellt.

| Version | Unterstützt |
|---------|-------------|
| 1.0.3   | ✅ |
| < 1.0.3 | ❌ |

## Sicherheitslücke melden

Bitte melde Sicherheitslücken **nicht** über öffentliche GitHub-Issues.

- Nutze die private Meldefunktion von GitHub: **Security → Report a vulnerability** im Repository, oder
- kontaktiere den Autor über [www.phinit.de](https://www.phinit.de).

Hilfreiche Angaben:
- betroffene Version (`.VERSION` im Script-Header bzw. Footer der GUI)
- Beschreibung und Auswirkung
- Schritte zur Reproduktion (ohne echte personenbezogene AD-Daten)

Du erhältst eine Rückmeldung, sobald die Meldung geprüft wurde. Bitte veröffentliche Details erst nach Bereitstellung einer Korrektur.

## Sicherheitsmodell des Tools

Das Tool läuft **im Kontext des angemeldeten Benutzers** und verwendet dessen Active-Directory-Berechtigungen. Es speichert keine Anmeldedaten und baut keine Verbindung zu Exchange Online oder Microsoft Entra ID auf.

### Umgesetzte Schutzmaßnahmen

| Bereich | Maßnahme |
|---------|----------|
| LDAP | Escaping aller Benutzereingaben in LDAP-Filtern (RFC 4515) |
| AD-Schreibzugriffe | Atomare Änderungen (`Set-ADUser -Replace`) mit vorherigem History-Backup |
| Eingabevalidierung | Typ-Präfix, RFC-5322-Adresssyntax, Steuerzeichen, max. 1123 Zeichen, X500-DN-Format |
| Eindeutigkeit | Prüfung von `proxyAddresses`, `mail` und `userPrincipalName` im Global Catalog |
| CSV-Export/History | Schutz vor CSV-/Formel-Injection (`=`, `+`, `-`, `@`, Tab, CR) |
| Logging | Neutralisierung von Zeilenumbrüchen (Log-Injection) |
| Daten im Repository | `.gitignore` schließt `Logs/`, `Exports/` und `History/` aus |

### Hinweise für Betreiber

- `Logs\`, `Exports\` und `History\` enthalten **personenbezogene Daten** (Namen, E-Mail-Adressen, DNs). Zugriff auf diese Ordner beschränken und Aufbewahrungsfristen (DSGVO) beachten. Logs werden nach `$script:MaxLogAgeDays` (Standard 90 Tage) gelöscht; Exports und History werden **nicht** automatisch bereinigt.
- Ist das Script-Verzeichnis nicht beschreibbar, werden Daten unter `%TEMP%\easyEXCH_PMA_Logs\` bzw. `%LOCALAPPDATA%\easyEXCH_PMA\` abgelegt.
- Das Tool nach dem Prinzip der minimalen Rechte betreiben: Schreibrechte nur auf `proxyAddresses` und `mail` der zu verwaltenden OUs delegieren.
- Scripts und EXE-Dateien nur aus vertrauenswürdiger Quelle verwenden; für den produktiven Einsatz wird eine Code-Signatur empfohlen (`Set-AuthenticodeSignature`).
