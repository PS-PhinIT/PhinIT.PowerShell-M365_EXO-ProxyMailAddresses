<#
.SYNOPSIS
    Easy Exchange ProxyAddress Manager - Verwaltung von ProxyMailAddresses für MS365 ADSync
.DESCRIPTION
    PowerShell Script mit WPF GUI zur Verwaltung von ProxyMailAddresses für Active Directory Benutzer.
    Lädt Benutzer aus lokalem AD, zeigt aktuelle ProxyAddresses und ermöglicht Hinzufügen/Ändern/Löschen.
    Prüft automatisch MS365 ADSync-Kompatibilität (berücksichtigt Legacy Exchange Attribute).
.VERSION
    0.0.1
.AUTHOR
    Andreas Hepp  |  PhinIT.DE
.DATE
    2025-12-08
#>

#Requires -Version 5.1

# Progress Bars deaktivieren (wichtig für -noConsole Kompilierung)
$ProgressPreference = 'SilentlyContinue'

# Assemblies laden
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# Visual Styles aktivieren (MUSS vor GUI-Erstellung erfolgen)
[System.Windows.Forms.Application]::EnableVisualStyles()

# XAML Definition
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="easyEXCH ProxyAddress Manager" Height="900" Width="1400"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#0078D4" Padding="10" CornerRadius="5" Margin="0,0,0,10">
            <StackPanel>
                <TextBlock Text="easyEXCH ProxyAddress Manager" FontSize="20" FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="Verwaltung von ProxyMailAddresses für MS365 ADSync" FontSize="12" Foreground="WhiteSmoke" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>

        <!-- Main Content -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2.5*"/>
                <ColumnDefinition Width="10"/>
                <ColumnDefinition Width="3.5*"/>
            </Grid.ColumnDefinitions>

            <!-- Left Panel: ProxyAddress Management -->
            <Border Grid.Column="0" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="5" Padding="10" Background="White">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <!-- User Info Header -->
                    <Border Grid.Row="0" Background="#F0F0F0" Padding="10" CornerRadius="5" Margin="0,0,0,10" Name="pnlUserInfo" Visibility="Collapsed">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0">
                                <TextBlock Name="txtSelectedUserName" Text="" FontSize="16" FontWeight="Bold"/>
                                <TextBlock Name="txtSelectedUserMail" Text="" FontSize="12" Foreground="Gray" Margin="0,2,0,0"/>
                                <TextBlock Name="txtSelectedUserSAM" Text="" FontSize="10" Foreground="DarkGray" Margin="0,2,0,0"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                <TextBlock Name="txtValidationStatus" Text="✅ MS365 Kompatibel" FontSize="12" FontWeight="Bold" Foreground="Green" HorizontalAlignment="Right"/>
                                <TextBlock Name="txtValidationDetails" Text="" FontSize="10" Foreground="Gray" Margin="0,2,0,0" TextWrapping="Wrap" MaxWidth="200" HorizontalAlignment="Right"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <!-- Validation Info -->
                    <Border Grid.Row="1" Background="#FFF3CD" BorderBrush="#FFC107" BorderThickness="1" Padding="8" CornerRadius="5" Margin="0,0,0,10" Name="pnlValidationInfo" Visibility="Collapsed">
                        <StackPanel>
                            <TextBlock Text="⚠️ MS365 ADSync Validierung" FontWeight="Bold" FontSize="11" Foreground="#856404"/>
                            <TextBlock Name="txtValidationMessages" Text="" FontSize="10" Foreground="#856404" Margin="0,5,0,0" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>

                    <!-- ProxyAddresses List -->
                    <GroupBox Header="Aktuelle ProxyAddresses" Grid.Row="2" Margin="0,0,0,10">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <!-- DataGrid -->
                            <DataGrid Name="dgProxyAddresses" Grid.Row="0"
                                      AutoGenerateColumns="False" 
                                      IsReadOnly="True" 
                                      AlternatingRowBackground="LightGray"
                                      GridLinesVisibility="All"
                                      CanUserSortColumns="True"
                                      CanUserResizeColumns="True"
                                      SelectionMode="Single">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Typ" Binding="{Binding Type}" Width="60"/>
                                    <DataGridTextColumn Header="Adresse" Binding="{Binding Address}" Width="*"/>
                                    <DataGridTextColumn Header="Primär" Binding="{Binding IsPrimary}" Width="50"/>
                                    <DataGridTextColumn Header="Status" Binding="{Binding ValidationStatus}" Width="80"/>
                                </DataGrid.Columns>
                            </DataGrid>

                            <!-- Action Buttons -->
                            <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,10,0,0" HorizontalAlignment="Left">
                                <Button Name="btnAddProxy" Content="➕" Height="30" Width="40" Background="#28A745" Foreground="White" FontWeight="Bold" Margin="0,0,5,0" IsEnabled="False" ToolTip="Hinzufügen"/>
                                <Button Name="btnEditProxy" Content="✏️" Height="30" Width="40" Background="#FFC107" Foreground="Black" FontWeight="Bold" Margin="0,0,5,0" IsEnabled="False" ToolTip="Bearbeiten"/>
                                <Button Name="btnDeleteProxy" Content="🗑️" Height="30" Width="40" Background="#DC3545" Foreground="White" FontWeight="Bold" Margin="0,0,5,0" IsEnabled="False" ToolTip="Löschen"/>
                                <Button Name="btnSetPrimary" Content="⭐" Height="30" Width="40" Background="#0078D4" Foreground="White" FontWeight="Bold" Margin="0,0,5,0" IsEnabled="False" ToolTip="Als Primär"/>
                                <Button Name="btnRefresh" Content="🔄" Height="30" Width="40" Background="#6C757D" Foreground="White" FontWeight="Bold" IsEnabled="False" ToolTip="Aktualisieren"/>
                            </StackPanel>

                            <!-- No Data Message -->
                            <TextBlock Name="txtNoProxyData" Grid.Row="0" Text="Kein Benutzer ausgewählt.&#x0a;Bitte wählen Sie einen Benutzer aus der Liste." 
                                       HorizontalAlignment="Center" VerticalAlignment="Center" 
                                       FontSize="14" Foreground="Gray" Visibility="Visible" TextAlignment="Center"/>
                        </Grid>
                    </GroupBox>

                </Grid>
            </Border>

            <!-- Right Panel: User Search & List -->
            <Border Grid.Column="2" BorderBrush="#E0E0E0" BorderThickness="1" CornerRadius="5" Padding="10" Background="White">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <!-- Search Options -->
                    <GroupBox Header="Benutzer suchen" Grid.Row="0" Margin="0,0,0,10">
                        <StackPanel Margin="5">
                            <TextBlock Text="Suchbegriff (min. 3 Zeichen - Vorname, Nachname, Email, SAM):" Margin="0,0,0,5" FontSize="11"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="5"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox Name="txtSearch" Grid.Column="0" Height="25" ToolTip="Automatische Suche ab 3 Zeichen"/>
                                <Button Name="btnSearch" Grid.Column="2" Content="🔍 Suchen" Height="25" Width="100" Background="#0078D4" Foreground="White" FontWeight="Bold"/>
                            </Grid>
                            <TextBlock Text="OU Filter (Optional):" Margin="0,10,0,5" FontSize="11"/>
                            <TextBox Name="txtOUFilter" Height="25" ToolTip="z.B. OU=Users,DC=domain,DC=com"/>
                        </StackPanel>
                    </GroupBox>

                    <!-- Load All Button -->
                    <GroupBox Header="Alle Benutzer" Grid.Row="1" Margin="0,0,0,10">
                        <StackPanel Margin="5">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="5"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <Button Name="btnLoadAll" Grid.Column="0" Content="👥 Alle Benutzer laden" Height="35" Background="#28A745" Foreground="White" FontWeight="Bold"/>
                                <Button Name="btnCheckADSync" Grid.Column="2" Content="🔍 ADSync prüfen" Height="35" Width="130" Background="#17A2B8" Foreground="White" FontWeight="Bold" IsEnabled="False" ToolTip="Prüft ob alle geladenen Benutzer ADSync-kompatibel sind"/>
                            </Grid>
                            <TextBlock Name="txtUserCount" Text="0 Benutzer geladen" Margin="0,5,0,0" FontSize="10" Foreground="Gray" HorizontalAlignment="Center"/>
                        </StackPanel>
                    </GroupBox>

                    <!-- User List -->
                    <GroupBox Header="Gefundene Benutzer" Grid.Row="2">
                        <Grid>
                            <DataGrid Name="dgUsers" 
                                      AutoGenerateColumns="False"
                                      IsReadOnly="True"
                                      SelectionMode="Single"
                                      AlternatingRowBackground="LightGray"
                                      GridLinesVisibility="All"
                                      CanUserSortColumns="True"
                                      CanUserResizeColumns="True"
                                      HeadersVisibility="Column">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Vorname" Binding="{Binding GivenName}" Width="100"/>
                                    <DataGridTextColumn Header="Nachname" Binding="{Binding Surname}" Width="120"/>
                                    <DataGridTextColumn Header="DisplayName" Binding="{Binding DisplayName}" Width="150"/>
                                    <DataGridTextColumn Header="Primäre E-Mail" Binding="{Binding PrimarySMTP}" Width="*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                            <TextBlock Name="txtNoUsers" Text="Keine Benutzer geladen.&#x0a;Bitte suchen oder alle laden." 
                                       HorizontalAlignment="Center" VerticalAlignment="Center" 
                                       FontSize="12" Foreground="Gray" Visibility="Visible" TextAlignment="Center"/>
                        </Grid>
                    </GroupBox>
                </Grid>
            </Border>
        </Grid>

        <!-- Footer -->
        <Border Grid.Row="2" Background="#F0F0F0" Padding="10" CornerRadius="5" Margin="0,10,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Name="txtStatus" Text="Bereit" VerticalAlignment="Center" Grid.Column="0"/>
                <TextBlock Text="Version 0.0.1 | Andreas Hepp  -  www.PhinIT.DE  2025" VerticalAlignment="Center" FontSize="10" Foreground="Gray" Grid.Column="1"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

#region Functions

# PS2EXE-kompatible Pfad-Erkennung
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
    $script:ScriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
} else {
    $script:ScriptPath = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
    if (!$script:ScriptPath) { $script:ScriptPath = "." }
}

# Globale Variablen
$script:CurrentUser = $null
$script:AllUsers = @()
$script:MS365Domains = @()

# Logging-Konfiguration
$script:CurrentExecutingUser = $env:USERNAME
$script:CurrentDate = Get-Date -Format "yyyy-MM-dd"
$script:LogRootFolder = Join-Path $script:ScriptPath "Logs"
$script:UserLogFolder = Join-Path $script:LogRootFolder $script:CurrentExecutingUser
$script:LogFile = Join-Path $script:UserLogFolder "$($script:CurrentExecutingUser)_$($script:CurrentDate).log"
$script:DebugLogFile = Join-Path $script:UserLogFolder "DEBUG_$($script:CurrentExecutingUser)_$($script:CurrentDate).log"
$script:MaxLogAgeDays = 90
$script:EnableDebugLogging = $true

# Log-Ordner initialisieren
function Initialize-LogFolder {
    try {
        # Haupt-Log-Ordner erstellen
        if (-not (Test-Path $script:LogRootFolder)) {
            New-Item -Path $script:LogRootFolder -ItemType Directory -Force | Out-Null
        }
        
        # Benutzer-spezifischen Log-Ordner erstellen
        if (-not (Test-Path $script:UserLogFolder)) {
            New-Item -Path $script:UserLogFolder -ItemType Directory -Force | Out-Null
        }
        
        # Alte Logs bereinigen (älter als MaxLogAgeDays)
        try {
            $cutoffDate = (Get-Date).AddDays(-$script:MaxLogAgeDays)
            Get-ChildItem -Path $script:UserLogFolder -Filter "*.log" -File | 
                Where-Object { $_.LastWriteTime -lt $cutoffDate } | 
                Remove-Item -Force -ErrorAction SilentlyContinue
        } catch {
            # Stille Fehlerbehandlung für Log-Cleanup
        }
        
        # Session-Header schreiben
        $psVersion = $PSVersionTable.PSVersion.ToString()
        $osVersion = [System.Environment]::OSVersion.VersionString
        $is64Bit = [Environment]::Is64BitProcess
        $domain = $env:USERDOMAIN
        $computer = $env:COMPUTERNAME
        
        $header = @"
================================================================================
easyEXCH ProxyAddress Manager - Logging Session
================================================================================
Session Start: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Script Version: 0.0.1
Ausführender Benutzer: $script:CurrentExecutingUser@$domain
Computer: $computer
PowerShell Version: $psVersion
OS Version: $osVersion
Architektur: $(if($is64Bit){'64-bit'}else{'32-bit'})
Script Path: $script:ScriptPath
Log Path: $script:UserLogFolder
================================================================================

"@
        Add-Content -Path $script:LogFile -Value $header -Encoding UTF8 -ErrorAction SilentlyContinue
        
        if ($script:EnableDebugLogging) {
            Add-Content -Path $script:DebugLogFile -Value $header -Encoding UTF8 -ErrorAction SilentlyContinue
        }
        
        return $true
    } catch {
        # Fallback: Versuche in TEMP zu loggen
        try {
            $script:LogRootFolder = Join-Path $env:TEMP "easyEXCH_PMA_Logs"
            $script:UserLogFolder = Join-Path $script:LogRootFolder $script:CurrentExecutingUser
            if (-not (Test-Path $script:UserLogFolder)) {
                New-Item -Path $script:UserLogFolder -ItemType Directory -Force | Out-Null
            }
            $script:LogFile = Join-Path $script:UserLogFolder "$($script:CurrentExecutingUser)_$($script:CurrentDate).log"
            $script:DebugLogFile = Join-Path $script:UserLogFolder "DEBUG_$($script:CurrentExecutingUser)_$($script:CurrentDate).log"
            return $true
        } catch {
            return $false
        }
    }
}

# Erweiterte Logging-Funktion
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR','ACTION','DEBUG')]
        [string]$Level = 'INFO',
        [string]$Source = '',
        [hashtable]$Context = @{}
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $threadId = [Threading.Thread]::CurrentThread.ManagedThreadId
    
    # Erweiterte Log-Message mit Kontext
    $logMessage = "[$timestamp] [TID:$threadId] [$Level]"
    if ($Source) {
        $logMessage += " [$Source]"
    }
    $logMessage += " $Message"
    
    # Kontext-Informationen hinzufügen falls vorhanden
    if ($Context.Count -gt 0) {
        $contextStr = ($Context.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ', '
        $logMessage += " | Context: {$contextStr}"
    }
    
    # Datei-Logging
    try {
        # Haupt-Log (alle Meldungen außer DEBUG)
        if ($Level -ne 'DEBUG') {
            Add-Content -Path $script:LogFile -Value $logMessage -Encoding UTF8 -ErrorAction SilentlyContinue
        }
        
        # Debug-Log (alle Meldungen)
        if ($script:EnableDebugLogging) {
            Add-Content -Path $script:DebugLogFile -Value $logMessage -Encoding UTF8 -ErrorAction SilentlyContinue
        }
    } catch {
        # Stille Fehlerbehandlung - Logging darf Hauptprozess nicht blockieren
    }
    
    # KEINE Console-Ausgaben für EXE-Konvertierung (PS2EXE -noConsole Kompatibilität)
}

# Spezielle Logging-Funktion für Änderungen an Benutzern
function Write-UserChangeLog {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Action,
        [Parameter(Mandatory=$true)]
        $TargetUser,
        [string]$ProxyAddress = '',
        [string]$OldValue = '',
        [string]$NewValue = '',
        [string]$Details = ''
    )
    
    $changeMessage = @"

╔═══════════════════════════════════════════════════════════════════════════╗
║                        BENUTZER-ÄNDERUNG PROTOKOLLIERT                     ║
╚═══════════════════════════════════════════════════════════════════════════╝
Zeitpunkt: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Ausgeführt von: $script:CurrentExecutingUser@$env:USERDOMAIN
Computer: $env:COMPUTERNAME

AKTION: $Action
ZIEL-BENUTZER: $($TargetUser.DisplayName) ($($TargetUser.SamAccountName))
E-Mail: $($TargetUser.Mail)
UPN: $($TargetUser.UserPrincipalName)
DN: $($TargetUser.DistinguishedName)
"@

    if ($ProxyAddress) {
        $changeMessage += "`nProxyAddress: $ProxyAddress"
    }
    
    if ($OldValue) {
        $changeMessage += "`nAlter Wert: $OldValue"
    }
    
    if ($NewValue) {
        $changeMessage += "`nNeuer Wert: $NewValue"
    }
    
    if ($Details) {
        $changeMessage += "`nDetails: $Details"
    }
    
    $changeMessage += "`n" + ("═" * 79) + "`n"
    
    # In Haupt-Log schreiben
    Add-Content -Path $script:LogFile -Value $changeMessage -Encoding UTF8 -ErrorAction SilentlyContinue
    
    # Auch als ACTION-Level loggen
    Write-Log -Message "$Action für Benutzer $($TargetUser.SamAccountName)" -Level ACTION -Source "UserChange" -Context @{
        TargetUser = $TargetUser.SamAccountName
        ProxyAddress = $ProxyAddress
        Action = $Action
    }
}

# Status-Funktion (GUI + Logging)
function Write-Status {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')]
        [string]$Level = 'INFO'
    )
    
    # In Log-Datei schreiben
    Write-Log -Message $Message -Level $Level -Source "GUI"
    
    # GUI Update
    if ($window -and $txtStatus) {
        try {
            $window.Dispatcher.Invoke([action]{
                $txtStatus.Text = $Message
                
                # Farbe je nach Level
                switch ($Level) {
                    'SUCCESS' { $txtStatus.Foreground = 'Green' }
                    'WARNING' { $txtStatus.Foreground = 'Orange' }
                    'ERROR' { $txtStatus.Foreground = 'Red' }
                    default { $txtStatus.Foreground = 'Black' }
                }
            }, [System.Windows.Threading.DispatcherPriority]::Send)
        } catch {
            # GUI nicht verfügbar
        }
    }
}

# AD Benutzer laden
function Get-ADUsersWithProxy {
    param(
        [string]$DomainController,
        [string]$SearchBase,
        [string]$Filter
    )
    
    try {
        Write-Log -Message "Starte AD-Benutzer-Abfrage" -Level DEBUG -Source "AD-Query" -Context @{
            DomainController = $DomainController
            SearchBase = $SearchBase
            Filter = $Filter
        }
        
        Write-Status -Message "Lade Benutzer aus Active Directory..." -Level INFO
        
        # Filter setzen - wenn keiner übergeben wurde, alle nehmen
        $adFilter = if ($Filter) { $Filter } else { '*' }
        
        $params = @{
            Filter = $adFilter
            Properties = @(
                'DisplayName', 'GivenName', 'Surname', 'Mail', 'SamAccountName', 'UserPrincipalName',
                'ProxyAddresses', 'msExchRecipientTypeDetails', 'msExchRecipientDisplayType',
                'mailNickname', 'targetAddress', 'legacyExchangeDN'
            )
        }
        
        if ($DomainController) {
            $params['Server'] = $DomainController
        }
        
        if ($SearchBase) {
            $params['SearchBase'] = $SearchBase
        }
        
        $users = @(Get-ADUser @params | Where-Object {
            $_.Enabled -eq $true
        } | Select-Object DisplayName, GivenName, Surname, Mail, SamAccountName, UserPrincipalName, 
                          ProxyAddresses, msExchRecipientTypeDetails, msExchRecipientDisplayType,
                          mailNickname, targetAddress, legacyExchangeDN, DistinguishedName)
        
        $userCount = if ($users) { $users.Count } else { 0 }
        Write-Log -Message "AD-Benutzer-Abfrage abgeschlossen: $userCount Benutzer gefunden" -Level INFO -Source "AD-Query"
        Write-Status -Message "$userCount Benutzer geladen" -Level SUCCESS
        return $users
        
    } catch {
        Write-Status -Message "Fehler beim Laden der Benutzer: $($_.Exception.Message)" -Level ERROR
        [System.Windows.MessageBox]::Show("Fehler beim Laden der Benutzer:`n`n$($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return @()
    }
}

# Benutzer suchen
function Search-ADUsers {
    param(
        [string]$SearchTerm,
        [string]$DomainController,
        [string]$SearchBase
    )
    
    try {
        Write-Log -Message "Suche nach '$SearchTerm'" -Level DEBUG -Source "Search" -Context @{
            SearchTerm = $SearchTerm
            SearchBase = $SearchBase
        }
        
        Write-Status -Message "Suche nach '$SearchTerm'..." -Level INFO
        
        # Erweiterte Suche: DisplayName, Vorname (GivenName), Nachname (Surname), Mail, SAM, UPN
        $filter = "(DisplayName -like '*$SearchTerm*') -or (GivenName -like '*$SearchTerm*') -or (Surname -like '*$SearchTerm*') -or (Mail -like '*$SearchTerm*') -or (SamAccountName -like '*$SearchTerm*') -or (UserPrincipalName -like '*$SearchTerm*')"
        
        $users = @(Get-ADUsersWithProxy -DomainController $DomainController -SearchBase $SearchBase -Filter $filter)
        
        $userCount = if ($users) { $users.Count } else { 0 }
        Write-Log -Message "Suche abgeschlossen: $userCount Benutzer gefunden" -Level INFO -Source "Search"
        Write-Status -Message "$userCount Benutzer gefunden" -Level SUCCESS
        return $users
        
    } catch {
        Write-Log -Message "Fehler bei der Suche: $($_.Exception.Message)" -Level ERROR -Source "Search"
        Write-Status -Message "Fehler bei der Suche: $($_.Exception.Message)" -Level ERROR
        return @()
    }
}

# ProxyAddresses parsen
function Get-ParsedProxyAddresses {
    param(
        [Parameter(Mandatory=$true)]
        $User
    )
    
    $proxyList = @()
    
    if ($User.ProxyAddresses) {
        foreach ($proxy in $User.ProxyAddresses) {
            $isPrimary = $false
            $type = ""
            $address = ""
            
            # Typ und Adresse trennen
            if ($proxy -match '^([A-Za-z0-9]+):(.+)$') {
                $type = $matches[1]
                $address = $matches[2]
                
                # Primäre Adresse erkennen (Großbuchstaben)
                if ($type -ceq $type.ToUpper()) {
                    $isPrimary = $true
                }
            }
            
            # Validierung
            $validationStatus = Test-ProxyAddressMS365Compatibility -ProxyAddress $proxy -User $User
            
            $proxyList += [PSCustomObject]@{
                FullAddress = $proxy
                Type = $type
                Address = $address
                IsPrimary = if ($isPrimary) { "Ja" } else { "Nein" }
                ValidationStatus = $validationStatus
            }
        }
    }
    
    return $proxyList
}

# MS365 Kompatibilität prüfen
function Test-ProxyAddressMS365Compatibility {
    param(
        [string]$ProxyAddress,
        $User
    )
    
    $issues = @()
    
    # SMTP-Adressen prüfen
    if ($ProxyAddress -match '^smtp:(.+)$' -or $ProxyAddress -match '^SMTP:(.+)$') {
        $emailAddress = $matches[1]
        
        # E-Mail Format prüfen
        if ($emailAddress -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
            return "❌ Ungültig"
        }
        
        # Sonderzeichen prüfen (MS365 erlaubt nicht alle)
        if ($emailAddress -match '[<>()[\]\\,;:\s"]') {
            $issues += "Unerlaubte Zeichen"
        }
        
        # Länge prüfen (max 254 Zeichen)
        if ($emailAddress.Length -gt 254) {
            $issues += "Zu lang"
        }
    }
    
    # X500 Legacy-Adressen (von altem Exchange)
    if ($ProxyAddress -match '^X500:') {
        # X500 ist OK für Migration, aber warnen
        return "⚠️ Legacy"
    }
    
    # SIP-Adressen
    if ($ProxyAddress -match '^SIP:' -or $ProxyAddress -match '^sip:') {
        # SIP ist OK für Skype/Teams
        return "✅ OK"
    }
    
    if ($issues.Count -gt 0) {
        return "⚠️ " + ($issues -join ", ")
    }
    
    return "✅ OK"
}

# Benutzer-Validierung für MS365 ADSync
function Test-UserMS365Compatibility {
    param(
        [Parameter(Mandatory=$true)]
        $User
    )
    
    $issues = @()
    $warnings = @()
    
    # 1. ProxyAddresses vorhanden?
    if (-not $User.ProxyAddresses -or $User.ProxyAddresses.Count -eq 0) {
        $issues += "Keine ProxyAddresses definiert"
    } else {
        # 2. Primäre SMTP-Adresse vorhanden?
        $primarySMTP = $User.ProxyAddresses | Where-Object { $_ -cmatch '^SMTP:' }
        if (-not $primarySMTP) {
            $issues += "Keine primäre SMTP-Adresse (SMTP: Großbuchstaben)"
        }
        
        # 3. Mehrere primäre SMTP-Adressen?
        if ($primarySMTP.Count -gt 1) {
            $issues += "Mehrere primäre SMTP-Adressen gefunden"
        }
        
        # 4. Mail-Attribut stimmt mit primärer SMTP überein?
        if ($primarySMTP -and $User.Mail) {
            $primaryAddress = ($primarySMTP[0] -replace '^SMTP:', '')
            if ($primaryAddress -ne $User.Mail) {
                $warnings += "Mail-Attribut ($($User.Mail)) stimmt nicht mit primärer SMTP-Adresse ($primaryAddress) überein"
            }
        }
    }
    
    # 5. Legacy Exchange Attribute prüfen
    if ($User.msExchRecipientTypeDetails) {
        $warnings += "Legacy Exchange Attribute gefunden (msExchRecipientTypeDetails)"
    }
    
    if ($User.legacyExchangeDN) {
        # Legacy DN ist OK, sollte als X500 ProxyAddress vorhanden sein
        $x500Address = "X500:" + $User.legacyExchangeDN
        if ($User.ProxyAddresses -notcontains $x500Address) {
            $warnings += "Legacy Exchange DN nicht als X500 ProxyAddress vorhanden"
        }
    }
    
    # 6. MailNickname vorhanden? (für Exchange Online wichtig)
    if (-not $User.mailNickname) {
        $warnings += "Kein mailNickname definiert (empfohlen für Exchange Online)"
    }
    
    # 7. TargetAddress für Remote Mailbox?
    if ($User.targetAddress) {
        $warnings += "TargetAddress vorhanden (Remote Mailbox Konfiguration)"
    }
    
    return @{
        IsValid = ($issues.Count -eq 0)
        Issues = $issues
        Warnings = $warnings
    }
}

# ProxyAddress hinzufügen
function Add-ProxyAddressToUser {
    param(
        [Parameter(Mandatory=$true)]
        $User,
        [Parameter(Mandatory=$true)]
        [string]$ProxyAddress
    )
    
    try {
        # Validierung
        if ([string]::IsNullOrWhiteSpace($ProxyAddress)) {
            throw "ProxyAddress darf nicht leer sein"
        }
        
        # Format prüfen
        if ($ProxyAddress -notmatch '^[A-Za-z0-9]+:.+$') {
            throw "Ungültiges Format. Erwartet: TYP:adresse (z.B. smtp:user@domain.com)"
        }
        
        # Prüfen ob bereits vorhanden
        $existingProxy = Get-ADUser -Identity $User.SamAccountName -Properties ProxyAddresses | Select-Object -ExpandProperty ProxyAddresses
        if ($existingProxy -contains $ProxyAddress) {
            throw "Diese ProxyAddress ist bereits vorhanden"
        }
        
        # Wenn primäre SMTP-Adresse (SMTP:), alte primäre zu sekundärer machen
        if ($ProxyAddress -cmatch '^SMTP:') {
            $oldPrimary = $existingProxy | Where-Object { $_ -cmatch '^SMTP:' }
            if ($oldPrimary) {
                # Alte primäre zu sekundärer konvertieren
                $newSecondary = $oldPrimary -replace '^SMTP:', 'smtp:'
                Set-ADUser -Identity $User.SamAccountName -Remove @{ProxyAddresses=$oldPrimary}
                Set-ADUser -Identity $User.SamAccountName -Add @{ProxyAddresses=$newSecondary}
            }
        }
        
        # Neue ProxyAddress hinzufügen
        Set-ADUser -Identity $User.SamAccountName -Add @{ProxyAddresses=$ProxyAddress}
        
        # Änderung protokollieren
        Write-UserChangeLog -Action "ProxyAddress HINZUGEFÜGT" -TargetUser $User -ProxyAddress $ProxyAddress -Details "ProxyAddress wurde erfolgreich hinzugefügt"
        
        Write-Status -Message "ProxyAddress erfolgreich hinzugefügt" -Level SUCCESS
        return $true
        
    } catch {
        Write-Status -Message "Fehler beim Hinzufügen: $($_.Exception.Message)" -Level ERROR
        [System.Windows.MessageBox]::Show("Fehler beim Hinzufügen der ProxyAddress:`n`n$($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

# ProxyAddress löschen
function Remove-ProxyAddressFromUser {
    param(
        [Parameter(Mandatory=$true)]
        $User,
        [Parameter(Mandatory=$true)]
        [string]$ProxyAddress
    )
    
    try {
        # Primäre SMTP-Adresse darf nicht gelöscht werden
        if ($ProxyAddress -cmatch '^SMTP:') {
            $result = [System.Windows.MessageBox]::Show(
                "Dies ist die primäre SMTP-Adresse. Wirklich löschen?`n`nWARNUNG: Der Benutzer hat dann keine primäre E-Mail-Adresse mehr!",
                "Primäre Adresse löschen?",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )
            
            if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
                return $false
            }
        }
        
        Set-ADUser -Identity $User.SamAccountName -Remove @{ProxyAddresses=$ProxyAddress}
        
        # Änderung protokollieren
        Write-UserChangeLog -Action "ProxyAddress GELÖSCHT" -TargetUser $User -ProxyAddress $ProxyAddress -Details "ProxyAddress wurde gelöscht"
        
        Write-Status -Message "ProxyAddress erfolgreich gelöscht" -Level SUCCESS
        return $true
        
    } catch {
        Write-Status -Message "Fehler beim Löschen: $($_.Exception.Message)" -Level ERROR
        [System.Windows.MessageBox]::Show("Fehler beim Löschen der ProxyAddress:`n`n$($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

# ProxyAddress als primär setzen
function Set-PrimaryProxyAddress {
    param(
        [Parameter(Mandatory=$true)]
        $User,
        [Parameter(Mandatory=$true)]
        [string]$ProxyAddress
    )
    
    try {
        # Nur SMTP-Adressen können primär sein
        if ($ProxyAddress -notmatch '^smtp:(.+)$') {
            throw "Nur SMTP-Adressen können als primär gesetzt werden"
        }
        
        $emailAddress = $matches[1]
        
        # Aktuelle ProxyAddresses laden
        $currentProxies = Get-ADUser -Identity $User.SamAccountName -Properties ProxyAddresses | Select-Object -ExpandProperty ProxyAddresses
        
        # Alte primäre SMTP zu sekundärer machen
        $oldPrimary = $currentProxies | Where-Object { $_ -cmatch '^SMTP:' }
        if ($oldPrimary) {
            $newSecondary = $oldPrimary -replace '^SMTP:', 'smtp:'
            Set-ADUser -Identity $User.SamAccountName -Remove @{ProxyAddresses=$oldPrimary}
            Set-ADUser -Identity $User.SamAccountName -Add @{ProxyAddresses=$newSecondary}
        }
        
        # Ausgewählte Adresse entfernen
        Set-ADUser -Identity $User.SamAccountName -Remove @{ProxyAddresses=$ProxyAddress}
        
        # Als primäre hinzufügen
        $newPrimary = "SMTP:$emailAddress"
        Set-ADUser -Identity $User.SamAccountName -Add @{ProxyAddresses=$newPrimary}
        
        # Mail-Attribut aktualisieren
        Set-ADUser -Identity $User.SamAccountName -EmailAddress $emailAddress
        
        # Änderung protokollieren
        Write-UserChangeLog -Action "PRIMÄRE SMTP-ADRESSE GEÄNDERT" -TargetUser $User -ProxyAddress $newPrimary -OldValue $oldPrimary -NewValue $newPrimary -Details "Primäre SMTP-Adresse wurde geändert. Alte primäre Adresse wurde zu sekundärer Adresse konvertiert."
        
        Write-Status -Message "Primäre SMTP-Adresse erfolgreich gesetzt" -Level SUCCESS
        return $true
        
    } catch {
        Write-Status -Message "Fehler beim Setzen der primären Adresse: $($_.Exception.Message)" -Level ERROR
        [System.Windows.MessageBox]::Show("Fehler beim Setzen der primären Adresse:`n`n$($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return $false
    }
}

# Benutzer-Liste aktualisieren
function Update-UserList {
    param(
        [Parameter(Mandatory=$true)]
        $Users
    )
    
    # Benutzer für DataGrid vorbereiten
    $userList = @()
    
    foreach ($user in $Users) {
        # Primäre SMTP-Adresse extrahieren
        $primarySMTP = ""
        if ($user.ProxyAddresses) {
            $primaryProxy = $user.ProxyAddresses | Where-Object { $_ -cmatch '^SMTP:' } | Select-Object -First 1
            if ($primaryProxy) {
                $primarySMTP = $primaryProxy -replace '^SMTP:', ''
            }
        }
        
        # Fallback auf Mail-Attribut wenn keine primäre SMTP
        if ([string]::IsNullOrWhiteSpace($primarySMTP) -and $user.Mail) {
            $primarySMTP = $user.Mail
        }
        
        # Benutzer-Objekt für DataGrid erstellen
        $userList += [PSCustomObject]@{
            GivenName = $user.GivenName
            Surname = $user.Surname
            DisplayName = $user.DisplayName
            PrimarySMTP = $primarySMTP
            # Original-Objekt für spätere Verwendung
            OriginalUser = $user
        }
    }
    
    $dgUsers.ItemsSource = $userList
}

# Benutzer aktualisieren
function Refresh-CurrentUser {
    if ($script:CurrentUser) {
        try {
            $refreshedUser = Get-ADUser -Identity $script:CurrentUser.SamAccountName -Properties @(
                'DisplayName', 'Mail', 'SamAccountName', 'UserPrincipalName',
                'ProxyAddresses', 'msExchRecipientTypeDetails', 'msExchRecipientDisplayType',
                'mailNickname', 'targetAddress', 'legacyExchangeDN'
            )
            
            $script:CurrentUser = $refreshedUser
            Update-ProxyAddressDisplay
            
        } catch {
            Write-Status -Message "Fehler beim Aktualisieren: $($_.Exception.Message)" -Level ERROR
        }
    }
}

# ProxyAddress-Anzeige aktualisieren
function Update-ProxyAddressDisplay {
    if (-not $script:CurrentUser) {
        return
    }
    
    # User Info anzeigen
    $pnlUserInfo.Visibility = 'Visible'
    $txtSelectedUserName.Text = $script:CurrentUser.DisplayName
    $txtSelectedUserMail.Text = "E-Mail: $($script:CurrentUser.Mail)"
    $txtSelectedUserSAM.Text = "SAM: $($script:CurrentUser.SamAccountName) | UPN: $($script:CurrentUser.UserPrincipalName)"
    
    # Validierung durchführen
    $validation = Test-UserMS365Compatibility -User $script:CurrentUser
    
    if ($validation.IsValid) {
        $txtValidationStatus.Text = "✅ MS365 Kompatibel"
        $txtValidationStatus.Foreground = "Green"
        $pnlValidationInfo.Visibility = 'Collapsed'
    } else {
        $txtValidationStatus.Text = "❌ Probleme gefunden"
        $txtValidationStatus.Foreground = "Red"
        $pnlValidationInfo.Visibility = 'Visible'
        
        $messages = @()
        if ($validation.Issues.Count -gt 0) {
            $messages += "FEHLER:"
            $messages += $validation.Issues | ForEach-Object { "  • $_" }
        }
        if ($validation.Warnings.Count -gt 0) {
            $messages += "WARNUNGEN:"
            $messages += $validation.Warnings | ForEach-Object { "  • $_" }
        }
        
        $txtValidationMessages.Text = $messages -join "`n"
    }
    
    # ProxyAddresses laden
    $proxyList = Get-ParsedProxyAddresses -User $script:CurrentUser
    
    if ($proxyList.Count -gt 0) {
        $dgProxyAddresses.ItemsSource = $proxyList
        $txtNoProxyData.Visibility = 'Collapsed'
    } else {
        $dgProxyAddresses.ItemsSource = $null
        $txtNoProxyData.Visibility = 'Visible'
        $txtNoProxyData.Text = "Keine ProxyAddresses vorhanden.`nFügen Sie eine hinzu."
    }
    
    # Buttons aktivieren
    $btnAddProxy.IsEnabled = $true
    $btnRefresh.IsEnabled = $true
}

#endregion Functions

#region GUI Initialization

# XAML parsen
try {
    [xml]$xamlContent = $xaml
    $reader = New-Object System.Xml.XmlNodeReader $xamlContent
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    [System.Windows.MessageBox]::Show("Fehler beim Laden der GUI: $($_.Exception.Message)", "Kritischer Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit 1
}

# GUI Elemente referenzieren
$txtSearch = $window.FindName("txtSearch")
$txtOUFilter = $window.FindName("txtOUFilter")
$btnSearch = $window.FindName("btnSearch")
$btnLoadAll = $window.FindName("btnLoadAll")
$btnCheckADSync = $window.FindName("btnCheckADSync")
$dgUsers = $window.FindName("dgUsers")
$txtUserCount = $window.FindName("txtUserCount")
$txtNoUsers = $window.FindName("txtNoUsers")

$pnlUserInfo = $window.FindName("pnlUserInfo")
$txtSelectedUserName = $window.FindName("txtSelectedUserName")
$txtSelectedUserMail = $window.FindName("txtSelectedUserMail")
$txtSelectedUserSAM = $window.FindName("txtSelectedUserSAM")
$txtValidationStatus = $window.FindName("txtValidationStatus")
$txtValidationDetails = $window.FindName("txtValidationDetails")

$pnlValidationInfo = $window.FindName("pnlValidationInfo")
$txtValidationMessages = $window.FindName("txtValidationMessages")

$dgProxyAddresses = $window.FindName("dgProxyAddresses")
$txtNoProxyData = $window.FindName("txtNoProxyData")

$btnAddProxy = $window.FindName("btnAddProxy")
$btnEditProxy = $window.FindName("btnEditProxy")
$btnDeleteProxy = $window.FindName("btnDeleteProxy")
$btnSetPrimary = $window.FindName("btnSetPrimary")
$btnRefresh = $window.FindName("btnRefresh")

$txtStatus = $window.FindName("txtStatus")

#endregion GUI Initialization

#region Logging Initialization

# Logging-System initialisieren
$logInitialized = Initialize-LogFolder

if ($logInitialized) {
    Write-Log -Message "easyEXCH ProxyAddress Manager gestartet" -Level INFO -Source "Startup"
    Write-Log -Message "GUI erfolgreich geladen" -Level INFO -Source "Startup"
} else {
    # Fallback-Warnung wenn Logging fehlschlägt
    [System.Windows.MessageBox]::Show(
        "Warnung: Logging-System konnte nicht vollständig initialisiert werden.`nDas Programm funktioniert weiterhin, aber Änderungen werden möglicherweise nicht protokolliert.",
        "Logging-Warnung",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Warning
    )
}

#endregion Logging Initialization

#region Event Handlers

# ADSync-Prüfung für alle geladenen Benutzer
$btnCheckADSync.Add_Click({
    if (-not $script:AllUsers -or $script:AllUsers.Count -eq 0) {
        return
    }
    
    try {
        Write-Status -Message "Prüfe $($script:AllUsers.Count) Benutzer auf ADSync-Kompatibilität..." -Level INFO
        Write-Log -Message "Starte ADSync-Prüfung für $($script:AllUsers.Count) Benutzer" -Level INFO -Source "ADSync-Check"
        
        $totalUsers = $script:AllUsers.Count
        $validUsers = 0
        $usersWithIssues = 0
        $usersWithWarnings = 0
        $detailedResults = @()
        
        foreach ($user in $script:AllUsers) {
            $validation = Test-UserMS365Compatibility -User $user
            
            if ($validation.IsValid) {
                $validUsers++
                $status = "✅ OK"
            } else {
                $usersWithIssues++
                $status = "❌ Fehler"
            }
            
            if ($validation.Warnings.Count -gt 0) {
                $usersWithWarnings++
            }
            
            $detailedResults += [PSCustomObject]@{
                DisplayName = $user.DisplayName
                SamAccountName = $user.SamAccountName
                Status = $status
                Fehler = if ($validation.Issues.Count -gt 0) { $validation.Issues -join "; " } else { "-" }
                Warnungen = if ($validation.Warnings.Count -gt 0) { $validation.Warnings -join "; " } else { "-" }
            }
        }
        
        # Ergebnis-Dialog
        $resultWindow = New-Object System.Windows.Window
        $resultWindow.Title = "ADSync-Prüfung Ergebnis"
        $resultWindow.Width = 900
        $resultWindow.Height = 600
        $resultWindow.WindowStartupLocation = 'CenterOwner'
        $resultWindow.Owner = $window
        
        $grid = New-Object System.Windows.Controls.Grid
        $grid.Margin = "10"
        
        $row1 = New-Object System.Windows.Controls.RowDefinition
        $row1.Height = "Auto"
        $row2 = New-Object System.Windows.Controls.RowDefinition
        $row2.Height = "*"
        $row3 = New-Object System.Windows.Controls.RowDefinition
        $row3.Height = "Auto"
        
        $grid.RowDefinitions.Add($row1)
        $grid.RowDefinitions.Add($row2)
        $grid.RowDefinitions.Add($row3)
        
        # Zusammenfassung
        $summary = New-Object System.Windows.Controls.TextBlock
        $summary.Text = "Geprüfte Benutzer: $totalUsers`n✅ Kompatibel: $validUsers | ❌ Fehler: $usersWithIssues | ⚠️ Warnungen: $usersWithWarnings"
        $summary.FontSize = 14
        $summary.FontWeight = "Bold"
        $summary.Margin = "0,0,0,10"
        $summary.Padding = "10"
        $summary.Background = "#F0F0F0"
        [System.Windows.Controls.Grid]::SetRow($summary, 0)
        $grid.Children.Add($summary)
        
        # DataGrid mit Ergebnissen
        $dgResults = New-Object System.Windows.Controls.DataGrid
        $dgResults.AutoGenerateColumns = $false
        $dgResults.IsReadOnly = $true
        $dgResults.AlternatingRowBackground = "LightGray"
        $dgResults.GridLinesVisibility = "All"
        $dgResults.CanUserSortColumns = $true
        $dgResults.CanUserResizeColumns = $true
        
        $col1 = New-Object System.Windows.Controls.DataGridTextColumn
        $col1.Header = "DisplayName"
        $col1.Binding = New-Object System.Windows.Data.Binding("DisplayName")
        $col1.Width = 150
        $dgResults.Columns.Add($col1)
        
        $col2 = New-Object System.Windows.Controls.DataGridTextColumn
        $col2.Header = "SAM"
        $col2.Binding = New-Object System.Windows.Data.Binding("SamAccountName")
        $col2.Width = 100
        $dgResults.Columns.Add($col2)
        
        $col3 = New-Object System.Windows.Controls.DataGridTextColumn
        $col3.Header = "Status"
        $col3.Binding = New-Object System.Windows.Data.Binding("Status")
        $col3.Width = 80
        $dgResults.Columns.Add($col3)
        
        $col4 = New-Object System.Windows.Controls.DataGridTextColumn
        $col4.Header = "Fehler"
        $col4.Binding = New-Object System.Windows.Data.Binding("Fehler")
        $col4.Width = "*"
        $dgResults.Columns.Add($col4)
        
        $col5 = New-Object System.Windows.Controls.DataGridTextColumn
        $col5.Header = "Warnungen"
        $col5.Binding = New-Object System.Windows.Data.Binding("Warnungen")
        $col5.Width = "*"
        $dgResults.Columns.Add($col5)
        
        $dgResults.ItemsSource = $detailedResults
        [System.Windows.Controls.Grid]::SetRow($dgResults, 1)
        $grid.Children.Add($dgResults)
        
        # Schließen-Button
        $btnClose = New-Object System.Windows.Controls.Button
        $btnClose.Content = "Schließen"
        $btnClose.Width = 100
        $btnClose.Height = 30
        $btnClose.Margin = "0,10,0,0"
        $btnClose.HorizontalAlignment = "Right"
        $btnClose.Add_Click({ $resultWindow.Close() })
        [System.Windows.Controls.Grid]::SetRow($btnClose, 2)
        $grid.Children.Add($btnClose)
        
        $resultWindow.Content = $grid
        
        Write-Log -Message "ADSync-Prüfung abgeschlossen: $validUsers/$totalUsers kompatibel, $usersWithIssues mit Fehlern, $usersWithWarnings mit Warnungen" -Level INFO -Source "ADSync-Check"
        Write-Status -Message "ADSync-Prüfung abgeschlossen: $validUsers/$totalUsers kompatibel" -Level SUCCESS
        
        [void]$resultWindow.ShowDialog()
        
    } catch {
        Write-Status -Message "Fehler bei ADSync-Prüfung: $($_.Exception.Message)" -Level ERROR
        Write-Log -Message "Fehler bei ADSync-Prüfung: $($_.Exception.Message)" -Level ERROR -Source "ADSync-Check"
    }
})

# Alle Benutzer laden
$btnLoadAll.Add_Click({
    try {
        $ou = if ($txtOUFilter.Text) { $txtOUFilter.Text } else { $null }
        
        $script:AllUsers = Get-ADUsersWithProxy -DomainController $null -SearchBase $ou
        
        if ($script:AllUsers.Count -gt 0) {
            Update-UserList -Users $script:AllUsers
            $txtNoUsers.Visibility = 'Collapsed'
            $txtUserCount.Text = "$($script:AllUsers.Count) Benutzer geladen"
            $btnCheckADSync.IsEnabled = $true
        } else {
            $dgUsers.ItemsSource = $null
            $txtNoUsers.Visibility = 'Visible'
            $txtUserCount.Text = "0 Benutzer geladen"
            $btnCheckADSync.IsEnabled = $false
        }
        
    } catch {
        Write-Status -Message "Fehler: $($_.Exception.Message)" -Level ERROR
    }
})

# Benutzer suchen (Funktion für Button und TextChanged)
$performSearch = {
    try {
        $searchTerm = $txtSearch.Text.Trim()
        
        # Mindestens 3 Zeichen erforderlich
        if ($searchTerm.Length -lt 3) {
            if ($searchTerm.Length -eq 0) {
                # Leeres Suchfeld - Liste leeren
                $dgUsers.ItemsSource = $null
                $txtNoUsers.Visibility = 'Visible'
                $txtNoUsers.Text = "Keine Benutzer geladen.`nBitte suchen oder alle laden."
                $txtUserCount.Text = "0 Benutzer geladen"
            } else {
                # Weniger als 3 Zeichen
                $txtUserCount.Text = "Mindestens 3 Zeichen eingeben..."
            }
            return
        }
        
        $ou = if ($txtOUFilter.Text) { $txtOUFilter.Text } else { $null }
        
        $script:AllUsers = Search-ADUsers -SearchTerm $searchTerm -DomainController $null -SearchBase $ou
        
        if ($script:AllUsers.Count -gt 0) {
            Update-UserList -Users $script:AllUsers
            $txtNoUsers.Visibility = 'Collapsed'
            $txtUserCount.Text = "$($script:AllUsers.Count) Benutzer gefunden"
            $btnCheckADSync.IsEnabled = $true
        } else {
            $dgUsers.ItemsSource = $null
            $txtNoUsers.Visibility = 'Visible'
            $txtNoUsers.Text = "Keine Benutzer gefunden für '$searchTerm'.`nBitte ändern Sie den Suchbegriff."
            $txtUserCount.Text = "0 Benutzer gefunden"
            $btnCheckADSync.IsEnabled = $false
        }
        
    } catch {
        Write-Status -Message "Fehler: $($_.Exception.Message)" -Level ERROR
    }
}

# Benutzer suchen - Button Click
$btnSearch.Add_Click($performSearch)

# Automatische Suche beim Tippen (ab 3 Zeichen)
$txtSearch.Add_TextChanged({
    $searchTerm = $txtSearch.Text.Trim()
    
    # Nur suchen wenn mindestens 3 Zeichen
    if ($searchTerm.Length -ge 3) {
        # Verzögerung für bessere Performance (verhindert zu viele Anfragen)
        if ($script:SearchTimer) {
            $script:SearchTimer.Stop()
        }
        
        $script:SearchTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:SearchTimer.Interval = [TimeSpan]::FromMilliseconds(500)
        $script:SearchTimer.Add_Tick({
            $script:SearchTimer.Stop()
            & $performSearch
        })
        $script:SearchTimer.Start()
    } elseif ($searchTerm.Length -eq 0) {
        # Suchfeld geleert - Liste leeren
        $dgUsers.ItemsSource = $null
        $txtNoUsers.Visibility = 'Visible'
        $txtNoUsers.Text = "Keine Benutzer geladen.`nBitte suchen oder alle laden."
        $txtUserCount.Text = "0 Benutzer geladen"
    } else {
        # Weniger als 3 Zeichen
        $txtUserCount.Text = "Mindestens 3 Zeichen eingeben..."
    }
})

# Enter-Taste in Suchfeld
$txtSearch.Add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq 'Return') {
        $btnSearch.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
    }
})

# Benutzer ausgewählt
$dgUsers.Add_SelectionChanged({
    if ($dgUsers.SelectedItem) {
        $selectedItem = $dgUsers.SelectedItem
        # OriginalUser-Objekt aus dem DataGrid-Item extrahieren
        if ($selectedItem.OriginalUser) {
            $script:CurrentUser = $selectedItem.OriginalUser
            Update-ProxyAddressDisplay
        }
    }
})

# ProxyAddress ausgewählt
$dgProxyAddresses.Add_SelectionChanged({
    if ($dgProxyAddresses.SelectedItem) {
        $btnEditProxy.IsEnabled = $true
        $btnDeleteProxy.IsEnabled = $true
        
        # Nur SMTP-Adressen können primär gesetzt werden
        $selected = $dgProxyAddresses.SelectedItem
        if ($selected.Type -match '^smtp$' -and $selected.IsPrimary -eq "Nein") {
            $btnSetPrimary.IsEnabled = $true
        } else {
            $btnSetPrimary.IsEnabled = $false
        }
    } else {
        $btnEditProxy.IsEnabled = $false
        $btnDeleteProxy.IsEnabled = $false
        $btnSetPrimary.IsEnabled = $false
    }
})

# ProxyAddress hinzufügen (Dialog)
$btnAddProxy.Add_Click({
    if (-not $script:CurrentUser) {
        return
    }
    
    # Input Dialog
    $inputDialog = New-Object System.Windows.Window
    $inputDialog.Title = "ProxyAddress hinzufügen"
    $inputDialog.Width = 500
    $inputDialog.Height = 250
    $inputDialog.WindowStartupLocation = 'CenterOwner'
    $inputDialog.Owner = $window
    
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = "10"
    
    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = "Auto"
    $row2 = New-Object System.Windows.Controls.RowDefinition
    $row2.Height = "Auto"
    $row3 = New-Object System.Windows.Controls.RowDefinition
    $row3.Height = "Auto"
    $row4 = New-Object System.Windows.Controls.RowDefinition
    $row4.Height = "*"
    $row5 = New-Object System.Windows.Controls.RowDefinition
    $row5.Height = "Auto"
    
    $grid.RowDefinitions.Add($row1)
    $grid.RowDefinitions.Add($row2)
    $grid.RowDefinitions.Add($row3)
    $grid.RowDefinitions.Add($row4)
    $grid.RowDefinitions.Add($row5)
    
    # Typ
    $lblType = New-Object System.Windows.Controls.TextBlock
    $lblType.Text = "Typ:"
    $lblType.Margin = "0,0,0,5"
    [System.Windows.Controls.Grid]::SetRow($lblType, 0)
    $grid.Children.Add($lblType)
    
    $cmbType = New-Object System.Windows.Controls.ComboBox
    $cmbType.Height = 25
    $cmbType.Margin = "0,0,0,10"
    $cmbType.Items.Add("SMTP: (Primär)")
    $cmbType.Items.Add("smtp: (Sekundär)")
    $cmbType.Items.Add("SIP:")
    $cmbType.Items.Add("X500:")
    $cmbType.SelectedIndex = 1
    [System.Windows.Controls.Grid]::SetRow($cmbType, 1)
    $grid.Children.Add($cmbType)
    
    # Adresse
    $lblAddress = New-Object System.Windows.Controls.TextBlock
    $lblAddress.Text = "Adresse:"
    $lblAddress.Margin = "0,0,0,5"
    [System.Windows.Controls.Grid]::SetRow($lblAddress, 2)
    $grid.Children.Add($lblAddress)
    
    $txtAddress = New-Object System.Windows.Controls.TextBox
    $txtAddress.Height = 25
    $txtAddress.Margin = "0,0,0,10"
    [System.Windows.Controls.Grid]::SetRow($txtAddress, 3)
    $grid.Children.Add($txtAddress)
    
    # Buttons
    $btnPanel = New-Object System.Windows.Controls.StackPanel
    $btnPanel.Orientation = 'Horizontal'
    $btnPanel.HorizontalAlignment = 'Right'
    [System.Windows.Controls.Grid]::SetRow($btnPanel, 4)
    
    $btnOK = New-Object System.Windows.Controls.Button
    $btnOK.Content = "Hinzufügen"
    $btnOK.Width = 100
    $btnOK.Height = 30
    $btnOK.Margin = "0,0,10,0"
    $btnOK.Background = "#28A745"
    $btnOK.Foreground = "White"
    
    $btnCancel = New-Object System.Windows.Controls.Button
    $btnCancel.Content = "Abbrechen"
    $btnCancel.Width = 100
    $btnCancel.Height = 30
    
    $btnOK.Add_Click({
        $type = ($cmbType.SelectedItem.ToString() -split ' ')[0]
        $address = $txtAddress.Text.Trim()
        
        if ([string]::IsNullOrWhiteSpace($address)) {
            [System.Windows.MessageBox]::Show("Bitte geben Sie eine Adresse ein.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $fullProxy = "$type$address"
        
        if (Add-ProxyAddressToUser -User $script:CurrentUser -ProxyAddress $fullProxy) {
            $inputDialog.DialogResult = $true
            $inputDialog.Close()
            Refresh-CurrentUser
        }
    })
    
    $btnCancel.Add_Click({
        $inputDialog.Close()
    })
    
    $btnPanel.Children.Add($btnOK)
    $btnPanel.Children.Add($btnCancel)
    $grid.Children.Add($btnPanel)
    
    $inputDialog.Content = $grid
    [void]$inputDialog.ShowDialog()
})

# ProxyAddress bearbeiten
$btnEditProxy.Add_Click({
    if (-not $script:CurrentUser -or -not $dgProxyAddresses.SelectedItem) {
        return
    }
    
    $selected = $dgProxyAddresses.SelectedItem
    
    # Input Dialog
    $inputDialog = New-Object System.Windows.Window
    $inputDialog.Title = "ProxyAddress bearbeiten"
    $inputDialog.Width = 500
    $inputDialog.Height = 220
    $inputDialog.WindowStartupLocation = 'CenterOwner'
    $inputDialog.Owner = $window
    
    $grid = New-Object System.Windows.Controls.Grid
    $grid.Margin = "10"
    
    $row1 = New-Object System.Windows.Controls.RowDefinition
    $row1.Height = "Auto"
    $row2 = New-Object System.Windows.Controls.RowDefinition
    $row2.Height = "Auto"
    $row3 = New-Object System.Windows.Controls.RowDefinition
    $row3.Height = "*"
    $row4 = New-Object System.Windows.Controls.RowDefinition
    $row4.Height = "Auto"
    
    $grid.RowDefinitions.Add($row1)
    $grid.RowDefinitions.Add($row2)
    $grid.RowDefinitions.Add($row3)
    $grid.RowDefinitions.Add($row4)
    
    # Info
    $lblInfo = New-Object System.Windows.Controls.TextBlock
    $lblInfo.Text = "Aktuelle ProxyAddress: $($selected.FullAddress)"
    $lblInfo.Margin = "0,0,0,10"
    $lblInfo.FontWeight = "Bold"
    [System.Windows.Controls.Grid]::SetRow($lblInfo, 0)
    $grid.Children.Add($lblInfo)
    
    # Neue Adresse
    $lblAddress = New-Object System.Windows.Controls.TextBlock
    $lblAddress.Text = "Neue Adresse (mit Typ, z.B. smtp:neue@domain.com):"
    $lblAddress.Margin = "0,0,0,5"
    [System.Windows.Controls.Grid]::SetRow($lblAddress, 1)
    $grid.Children.Add($lblAddress)
    
    $txtAddress = New-Object System.Windows.Controls.TextBox
    $txtAddress.Height = 25
    $txtAddress.Margin = "0,0,0,10"
    $txtAddress.Text = $selected.FullAddress
    [System.Windows.Controls.Grid]::SetRow($txtAddress, 2)
    $grid.Children.Add($txtAddress)
    
    # Buttons
    $btnPanel = New-Object System.Windows.Controls.StackPanel
    $btnPanel.Orientation = 'Horizontal'
    $btnPanel.HorizontalAlignment = 'Right'
    [System.Windows.Controls.Grid]::SetRow($btnPanel, 3)
    
    $btnOK = New-Object System.Windows.Controls.Button
    $btnOK.Content = "Speichern"
    $btnOK.Width = 100
    $btnOK.Height = 30
    $btnOK.Margin = "0,0,10,0"
    $btnOK.Background = "#28A745"
    $btnOK.Foreground = "White"
    
    $btnCancel = New-Object System.Windows.Controls.Button
    $btnCancel.Content = "Abbrechen"
    $btnCancel.Width = 100
    $btnCancel.Height = 30
    
    $btnOK.Add_Click({
        $newProxy = $txtAddress.Text.Trim()
        
        if ([string]::IsNullOrWhiteSpace($newProxy)) {
            [System.Windows.MessageBox]::Show("Bitte geben Sie eine Adresse ein.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        if ($newProxy -eq $selected.FullAddress) {
            [System.Windows.MessageBox]::Show("Die Adresse wurde nicht geändert.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            $inputDialog.Close()
            return
        }
        
        # Alte löschen, neue hinzufügen
        if (Remove-ProxyAddressFromUser -User $script:CurrentUser -ProxyAddress $selected.FullAddress) {
            if (Add-ProxyAddressToUser -User $script:CurrentUser -ProxyAddress $newProxy) {
                Write-UserChangeLog -Action "ProxyAddress BEARBEITET" -TargetUser $script:CurrentUser -ProxyAddress $newProxy -OldValue $selected.FullAddress -NewValue $newProxy -Details "ProxyAddress wurde geändert"
                $inputDialog.DialogResult = $true
                $inputDialog.Close()
                Refresh-CurrentUser
            } else {
                # Fehler beim Hinzufügen - alte wieder hinzufügen
                Add-ProxyAddressToUser -User $script:CurrentUser -ProxyAddress $selected.FullAddress | Out-Null
            }
        }
    })
    
    $btnCancel.Add_Click({
        $inputDialog.Close()
    })
    
    $btnPanel.Children.Add($btnOK)
    $btnPanel.Children.Add($btnCancel)
    $grid.Children.Add($btnPanel)
    
    $inputDialog.Content = $grid
    [void]$inputDialog.ShowDialog()
})

# ProxyAddress löschen
$btnDeleteProxy.Add_Click({
    if (-not $script:CurrentUser -or -not $dgProxyAddresses.SelectedItem) {
        return
    }
    
    $selected = $dgProxyAddresses.SelectedItem
    
    $result = [System.Windows.MessageBox]::Show(
        "ProxyAddress wirklich löschen?`n`n$($selected.FullAddress)",
        "Löschen bestätigen",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        if (Remove-ProxyAddressFromUser -User $script:CurrentUser -ProxyAddress $selected.FullAddress) {
            Refresh-CurrentUser
        }
    }
})

# Als Primär setzen
$btnSetPrimary.Add_Click({
    if (-not $script:CurrentUser -or -not $dgProxyAddresses.SelectedItem) {
        return
    }
    
    $selected = $dgProxyAddresses.SelectedItem
    
    $result = [System.Windows.MessageBox]::Show(
        "Diese Adresse als primäre SMTP-Adresse setzen?`n`n$($selected.Address)`n`nDie aktuelle primäre Adresse wird zu einer sekundären Adresse.",
        "Primäre Adresse ändern",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        if (Set-PrimaryProxyAddress -User $script:CurrentUser -ProxyAddress $selected.FullAddress) {
            Refresh-CurrentUser
        }
    }
})

# Aktualisieren
$btnRefresh.Add_Click({
    Refresh-CurrentUser
})

#endregion Event Handlers

#region Window Events

# Beim Schließen des Fensters
$window.Add_Closing({
    Write-Log -Message "easyEXCH ProxyAddress Manager wird beendet" -Level INFO -Source "Shutdown"
    
    # Session-Footer schreiben
    $footer = @"

================================================================================
Session Ende: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Ausführender Benutzer: $script:CurrentExecutingUser@$env:USERDOMAIN
================================================================================
"@
    Add-Content -Path $script:LogFile -Value $footer -Encoding UTF8 -ErrorAction SilentlyContinue
    
    if ($script:EnableDebugLogging) {
        Add-Content -Path $script:DebugLogFile -Value $footer -Encoding UTF8 -ErrorAction SilentlyContinue
    }
})

#endregion Window Events

# Fenster anzeigen
Write-Status -Message "Bereit" -Level INFO
Write-Log -Message "GUI bereit - Warte auf Benutzerinteraktion" -Level INFO -Source "Startup"
[void]$window.ShowDialog()
