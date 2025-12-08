# Beispiele / Examples

## Grundlegende Verwendung / Basic Usage

### 1. Benutzer suchen / Search for User

```powershell
# Starten Sie die Anwendung / Start the application
.\easyEXCH-ProxyMailAddresses.ps1

# In der GUI / In the GUI:
# 1. Geben Sie einen Suchbegriff ein (z.B. "Mueller" oder "max.mueller@contoso.com")
# 2. Klicken Sie auf "Search" oder drücken Sie Enter
# 3. Wählen Sie den gewünschten Benutzer aus der Liste
```

### 2. ProxyAddress hinzufügen / Add ProxyAddress

**Beispiel: Sekundäre E-Mail-Adresse hinzufügen**

```
1. Benutzer auswählen: Max Müller
2. Klick auf "Add Address"
3. Typ auswählen: "smtp (Secondary)"
4. Adresse eingeben: max.mueller@contoso.de
5. Klick auf "Save"

Ergebnis:
- smtp:max.mueller@contoso.de wird hinzugefügt
- Benutzer kann jetzt E-Mails an beide Adressen empfangen
```

**Beispiel: SIP-Adresse für Teams hinzufügen**

```
1. Benutzer auswählen: Max Müller
2. Klick auf "Add Address"
3. Typ auswählen: "SIP"
4. Adresse eingeben: max.mueller@contoso.com
5. Klick auf "Save"

Ergebnis:
- sip:max.mueller@contoso.com wird hinzugefügt
- Wird für Microsoft Teams/Skype verwendet
```

### 3. Primäre E-Mail-Adresse ändern / Change Primary Email

**Szenario: Benutzer wechselt von alter zu neuer Domain**

```
Ausgangssituation:
- SMTP:max.mueller@oldcompany.com (primär)
- smtp:m.mueller@oldcompany.com (alias)

Schritte:
1. Neue Adresse hinzufügen:
   - Typ: "smtp (Secondary)"
   - Adresse: max.mueller@newcompany.com
   
2. Neue Adresse als primär setzen:
   - max.mueller@newcompany.com in Liste auswählen
   - Klick auf "Set Primary"
   - Bestätigen mit "Yes"

Ergebnis:
- SMTP:max.mueller@newcompany.com (primär) ✓
- smtp:max.mueller@oldcompany.com (alias)
- smtp:m.mueller@oldcompany.com (alias)
```

## Erweiterte Szenarien / Advanced Scenarios

### Szenario 1: Migration zu MS365

**Situation**: Lokaler Exchange Server wird abgeschaltet, Migration zu MS365

```
Vor der Migration:
- SMTP:max.mueller@contoso.local (primär, on-premises)
- smtp:m.mueller@contoso.local (alias)

Migrations-Schritte:

1. MS365-Adresse als sekundär hinzufügen:
   Typ: smtp (Secondary)
   Adresse: max.mueller@contoso.com
   
2. Testen der neuen Adresse:
   - E-Mail an max.mueller@contoso.com senden
   - Empfang bestätigen
   
3. MS365-Adresse als primär setzen:
   - max.mueller@contoso.com auswählen
   - "Set Primary" klicken
   
4. ADSync laufen lassen:
   - Warten auf Synchronisation
   - In MS365 prüfen
   
Nach der Migration:
- SMTP:max.mueller@contoso.com (primär, MS365) ✓
- smtp:max.mueller@contoso.local (alias, für Übergangszeit)
- smtp:m.mueller@contoso.local (alias)
```

### Szenario 2: Abteilungswechsel mit E-Mail-Änderung

**Situation**: Mitarbeiter wechselt von Sales zu Marketing

```
Vor dem Wechsel:
- SMTP:max.mueller@sales.contoso.com (primär)
- smtp:m.mueller@sales.contoso.com (alias)

Schritte:

1. Neue Marketing-Adresse hinzufügen:
   Typ: smtp (Secondary)
   Adresse: max.mueller@marketing.contoso.com
   
2. Zusätzlichen Alias hinzufügen:
   Typ: smtp (Secondary)
   Adresse: m.mueller@marketing.contoso.com
   
3. Übergangsphase (2 Wochen):
   - Beide Adressen funktionieren
   - E-Mail-Signatur aktualisieren
   - Kollegen informieren
   
4. Marketing-Adresse als primär setzen:
   - max.mueller@marketing.contoso.com auswählen
   - "Set Primary" klicken
   
Nach dem Wechsel:
- SMTP:max.mueller@marketing.contoso.com (primär) ✓
- smtp:m.mueller@marketing.contoso.com (alias)
- smtp:max.mueller@sales.contoso.com (alias, temporär)
- smtp:m.mueller@sales.contoso.com (alias, temporär)

Nach Übergangsphase:
- Alte Sales-Adressen entfernen (optional)
```

### Szenario 3: Mehrere E-Mail-Aliase für Service-Account

**Situation**: Service-Account für verschiedene Kontaktpunkte

```
Benötigte Aliase für support@contoso.com:

1. Hauptadresse hinzufügen/setzen:
   SMTP:support@contoso.com (primär)

2. Mehrere Aliase hinzufügen:
   smtp:help@contoso.com
   smtp:service@contoso.com
   smtp:helpdesk@contoso.com
   smtp:kundenservice@contoso.com
   smtp:support@contoso.de

Ergebnis:
- Alle E-Mails an diese Adressen landen im selben Postfach
- Flexibilität bei Marketing-Materialien
- Verschiedene Sprach-Domains abgedeckt
```

### Szenario 4: Externe E-Mail-Domain hinzufügen

**Situation**: Firma übernimmt andere Firma, deren E-Mail-Domain behalten

```
Hauptfirma: contoso.com
Übernommene Firma: fabrikam.com

Schritte für betroffene Mitarbeiter:

1. Aktuelle contoso.com Adresse bleibt primär:
   SMTP:john.doe@contoso.com

2. Fabrikam-Alias hinzufügen:
   Typ: smtp (Secondary)
   Adresse: john.doe@fabrikam.com

3. Optional weitere Fabrikam-Aliase:
   smtp:j.doe@fabrikam.com

Ergebnis:
- Mitarbeiter erreichbar unter beiden Domains
- Fabrikam-Kunden können weiter alte Adressen nutzen
- Schrittweise Migration möglich
```

## PowerShell-Äquivalente / PowerShell Equivalents

Für Automatisierung oder Batch-Operationen:

### ProxyAddress hinzufügen

```powershell
# Sekundäre SMTP-Adresse hinzufügen
$user = Get-ADUser -Identity "mmüller" -Properties ProxyAddresses
Set-ADUser -Identity $user -Add @{ProxyAddresses="smtp:max.mueller@contoso.de"}

# SIP-Adresse hinzufügen
Set-ADUser -Identity "mmüller" -Add @{ProxyAddresses="sip:max.mueller@contoso.com"}
```

### Primäre SMTP-Adresse ändern

```powershell
$user = Get-ADUser -Identity "mmüller" -Properties ProxyAddresses
$newAddresses = @()

# Alle existierenden Adressen durchgehen
foreach ($addr in $user.ProxyAddresses) {
    if ($addr -like "SMTP:*") {
        # Alte primäre zu sekundärer machen
        $newAddresses += $addr -replace "^SMTP:", "smtp:"
    } else {
        $newAddresses += $addr
    }
}

# Neue primäre Adresse hinzufügen
$newAddresses += "SMTP:max.mueller@newcompany.com"

# Alle Adressen setzen
Set-ADUser -Identity $user -Replace @{ProxyAddresses=$newAddresses}
```

### ProxyAddress entfernen

```powershell
# Einzelne Adresse entfernen
Set-ADUser -Identity "mmüller" -Remove @{ProxyAddresses="smtp:old.address@contoso.com"}
```

### Alle ProxyAddresses anzeigen

```powershell
# Alle ProxyAddresses eines Benutzers
$user = Get-ADUser -Identity "mmüller" -Properties ProxyAddresses
$user.ProxyAddresses | ForEach-Object {
    if ($_ -match "^([^:]+):(.+)$") {
        [PSCustomObject]@{
            Type = $matches[1]
            Address = $matches[2]
            IsPrimary = ($matches[1] -ceq $matches[1].ToUpper())
        }
    }
} | Format-Table -AutoSize
```

## Häufige Fehler und Lösungen / Common Errors and Solutions

### Fehler: "This proxy address already exists"

```
Ursache: Sie versuchen eine bereits vorhandene Adresse hinzuzufügen
Lösung: Prüfen Sie die Liste der vorhandenen Adressen
```

### Fehler: "Invalid email format"

```
Ursache: E-Mail-Adresse entspricht nicht dem Standard-Format
Lösung: Prüfen Sie das Format (user@domain.com)
Beispiele gültig:
  ✓ max.mueller@contoso.com
  ✓ m.mueller@sub.contoso.com
  ✓ mueller_max@contoso.de
  
Beispiele ungültig:
  ✗ max.mueller (fehlt @domain)
  ✗ @contoso.com (fehlt Benutzername)
  ✗ max..mueller@contoso.com (doppelter Punkt)
```

### Warnung: "You are about to remove the primary SMTP address"

```
Ursache: Sie versuchen die primäre SMTP-Adresse zu entfernen
Auswirkung: E-Mail-Empfang könnte beeinträchtigt werden
Empfehlung: 
  1. Zuerst andere Adresse als primär setzen
  2. Dann alte primäre Adresse entfernen
```

## Best Practices

1. **Vor größeren Änderungen**:
   - Screenshot der aktuellen ProxyAddresses machen
   - Änderungen dokumentieren
   - Bei kritischen Accounts: Mit Test-User erst testen

2. **Bei Domain-Wechsel**:
   - Neue Adresse erst als sekundär hinzufügen
   - Einige Tage parallel laufen lassen
   - Dann als primär setzen
   - Alte Adresse noch einige Wochen als Alias behalten

3. **Bei Problemen**:
   - Refresh-Button verwenden für aktuelle Daten
   - ADSync-Status prüfen
   - In MS365 Admin Center gegenchecken

4. **Regelmäßig**:
   - Ungenutzte Aliase entfernen
   - Konsistenz der Adressformate prüfen
   - Dokumentation aktuell halten
