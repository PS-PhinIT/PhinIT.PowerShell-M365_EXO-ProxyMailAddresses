<#
.SYNOPSIS
    Easy Exchange ProxyAddress Manager - Verwaltung von ProxyMailAddresses für MS365 ADSync
.DESCRIPTION
    PowerShell Script mit WPF GUI zur Verwaltung von ProxyMailAddresses für Active Directory Benutzer.
    Lädt Benutzer aus lokalem AD, zeigt aktuelle ProxyAddresses und ermöglicht Hinzufügen/Ändern/Löschen.
    Prüft automatisch MS365 ADSync-Kompatibilität (berücksichtigt Legacy Exchange Attribute).
.VERSION
    1.0.1
.AUTHOR
    Andreas Hepp  |  PhinIT.DE
.DATE
    2026-10-14
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

# Script-Version (zentral, wird in Logs verwendet)
$script:AppVersion = "1.0.1"

# ActiveDirectory-Modul prüfen und laden (Voraussetzung für alle AD-Befehle)
try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    [System.Windows.MessageBox]::Show(
        "Das PowerShell-Modul 'ActiveDirectory' konnte nicht geladen werden.`n`nBitte RSAT installieren:`nAdd-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0`n`nDetails: $($_.Exception.Message)",
        "ActiveDirectory-Modul fehlt",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Error
    ) | Out-Null
    exit 1
}

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
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <!-- Titel -->
                <StackPanel Grid.Column="0">
                    <TextBlock Name="txtTitle" Text="easyEXCH ProxyAddress Manager" FontSize="20" FontWeight="Bold" Foreground="White"/>
                    <TextBlock Name="txtSubtitle" Text="Verwaltung von ProxyMailAddresses für MS365 ADSync" FontSize="12" Foreground="WhiteSmoke" Margin="0,5,0,0"/>
                </StackPanel>
                
                <!-- Sprach-Schalter -->
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="🌐" FontSize="16" Foreground="White" VerticalAlignment="Center" Margin="0,0,5,0"/>
                    <Button Name="btnLangDE" Content="DE" Width="40" Height="30" Margin="0,0,5,0" Background="#28A745" Foreground="White" FontWeight="Bold" ToolTip="Deutsch"/>
                    <Button Name="btnLangEN" Content="EN" Width="40" Height="30" Background="#6C757D" Foreground="White" FontWeight="Bold" ToolTip="English"/>
                </StackPanel>
            </Grid>
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
                            <TextBlock Name="lblValidationTitle" Text="⚠️ MS365 ADSync Validierung" FontWeight="Bold" FontSize="11" Foreground="#856404"/>
                            <TextBlock Name="txtValidationMessages" Text="" FontSize="10" Foreground="#856404" Margin="0,5,0,0" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>

                    <!-- ProxyAddresses List -->
                    <GroupBox Name="grpProxyAddresses" Header="Aktuelle ProxyAddresses" Grid.Row="2" Margin="0,0,0,10">
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
                            <Grid Grid.Row="1" Margin="0,10,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="5"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="5"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="5"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="5"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                
                                <Button Name="btnAddProxy" Grid.Column="0" Content="➕ Hinzufügen" Height="35" Background="#28A745" Foreground="White" FontWeight="Bold" IsEnabled="False" FontSize="11"/>
                                <Button Name="btnEditProxy" Grid.Column="2" Content="✏️ Bearbeiten" Height="35" Background="#FFC107" Foreground="Black" FontWeight="Bold" IsEnabled="False" FontSize="11"/>
                                <Button Name="btnDeleteProxy" Grid.Column="4" Content="🗑️ Löschen" Height="35" Background="#DC3545" Foreground="White" FontWeight="Bold" IsEnabled="False" FontSize="11"/>
                                <Button Name="btnSetPrimary" Grid.Column="6" Content="⭐ Als Primär" Height="35" Background="#0078D4" Foreground="White" FontWeight="Bold" IsEnabled="False" FontSize="11"/>
                                <Button Name="btnRefresh" Grid.Column="8" Content="🔄 Aktualisieren" Height="35" Background="#6C757D" Foreground="White" FontWeight="Bold" IsEnabled="False" FontSize="11"/>
                            </Grid>

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
                    <GroupBox Name="grpSearchUsers" Header="Benutzer suchen" Grid.Row="0" Margin="0,0,0,10">
                        <StackPanel Margin="5">
                            <TextBlock Name="lblSearchTerm" Text="Suchbegriff (min. 3 Zeichen - Vorname, Nachname, Email, SAM):" Margin="0,0,0,5" FontSize="11"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="5"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox Name="txtSearch" Grid.Column="0" Height="25" ToolTip="Automatische Suche ab 3 Zeichen"/>
                                <Button Name="btnSearch" Grid.Column="2" Content="🔍 Suchen" Height="25" Width="100" Background="#0078D4" Foreground="White" FontWeight="Bold"/>
                            </Grid>
                        </StackPanel>
                    </GroupBox>

                    <!-- Load All Button -->
                    <GroupBox Name="grpLoadUsers" Header="Alle Benutzer" Grid.Row="1" Margin="0,0,0,10">
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
                    <GroupBox Name="grpFoundUsers" Header="Gefundene Benutzer" Grid.Row="2">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            
                            <!-- DataGrid -->
                            <Grid Grid.Row="0">
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
                            
                            <!-- Action Buttons -->
                            <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,10,0,0" HorizontalAlignment="Left">
                                <Button Name="btnExport" Content="📤 Export CSV" Height="30" Width="110" Background="#17A2B8" Foreground="White" FontWeight="Bold" Margin="0,0,5,0" IsEnabled="False" ToolTip="Exportiert alle geladenen ProxyAddresses als CSV"/>
                                <ComboBox Name="cmbFilter" Width="150" Height="30" Margin="0,0,5,0" SelectedIndex="0">
                                    <ComboBoxItem Content="🔍 Alle anzeigen"/>
                                    <ComboBoxItem Content="✅ Nur kompatible"/>
                                    <ComboBoxItem Content="⚠️ Mit Warnungen"/>
                                    <ComboBoxItem Content="❌ Mit Fehlern"/>
                                </ComboBox>
                                <Button Name="btnHistory" Content="📜 History" Height="30" Width="100" Background="#6C757D" Foreground="White" FontWeight="Bold" IsEnabled="False" ToolTip="Zeigt Änderungsverlauf des ausgewählten Benutzers"/>
                                <Button Name="btnStatistics" Content="📊 Statistik" Height="30" Width="100" Background="#FFC107" Foreground="Black" FontWeight="Bold" Margin="5,0,0,0" IsEnabled="False" ToolTip="Zeigt Statistik-Dashboard"/>
                            </StackPanel>
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
                <TextBlock Text="Version 1.0.1 | Andreas Hepp  -  www.PhinIT.DE  2025" VerticalAlignment="Center" FontSize="10" Foreground="Gray" Grid.Column="1"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

#region Language Resources

# Aktuelle Sprache (Standard: Deutsch)
$script:CurrentLanguage = "DE"

# Sprachressourcen
$script:LanguageStrings = @{
    DE = @{
        # Header
        Title = "easyEXCH ProxyAddress Manager"
        Subtitle = "Verwaltung von ProxyMailAddresses für MS365 ADSync"
        
        # Buttons
        Search = "🔍 Suchen"
        LoadAll = "👥 Alle Benutzer laden"
        CheckADSync = "🔍 ADSync prüfen"
        Export = "📤 Export CSV"
        History = "📜 History"
        Statistics = "📊 Statistik"
        Add = "➕ Hinzufügen"
        Edit = "✏️ Bearbeiten"
        Delete = "🗑️ Löschen"
        SetPrimary = "⭐ Als Primär"
        Refresh = "🔄 Aktualisieren"
        
        # GroupBox Headers
        CurrentProxyAddresses = "Aktuelle ProxyAddresses"
        SearchUsers = "Benutzer suchen"
        LoadUsers = "Alle Benutzer"
        FoundUsers = "Gefundene Benutzer"
        
        # Labels
        SearchLabel = "Suchbegriff (min. 3 Zeichen - Vorname, Nachname, Email, SAM):"
        UserInfo = "Benutzer-Informationen"
        ValidationStatus = "Validierungsstatus"
        ValidationTitle = "⚠️ MS365 ADSync Validierung"
        Compatible = "✅ MS365 Kompatibel"
        NotCompatible = "❌ Nicht kompatibel"
        
        # DataGrid Columns - ProxyAddresses
        ColType = "Typ"
        ColAddress = "Adresse"
        ColPrimary = "Primär"
        ColStatus = "Status"
        
        # DataGrid Columns - Users
        ColGivenName = "Vorname"
        ColSurname = "Nachname"
        ColDisplayName = "Anzeigename"
        ColPrimarySMTP = "Primäre SMTP"
        
        # Tooltips
        TooltipAdd = "Hinzufügen"
        TooltipEdit = "Bearbeiten"
        TooltipDelete = "Löschen"
        TooltipSetPrimary = "Als Primär"
        TooltipRefresh = "Aktualisieren"
        TooltipGerman = "Deutsch"
        TooltipEnglish = "English"
        TooltipADSync = "Prüft ob alle geladenen Benutzer ADSync-kompatibel sind"
        TooltipSearch = "Automatische Suche ab 3 Zeichen"
        
        # Filter
        FilterAll = "🔍 Alle anzeigen"
        FilterValid = "✅ Nur kompatible"
        FilterWarnings = "⚠️ Mit Warnungen"
        FilterErrors = "❌ Mit Fehlern"
        
        # Messages
        NoUsersLoaded = "Keine Benutzer geladen.`nBitte suchen oder alle laden."
        NoUsersFound = "Keine Benutzer gefunden für '{0}'.`nBitte ändern Sie den Suchbegriff."
        NoUserSelected = "Kein Benutzer ausgewählt.`nBitte wählen Sie einen Benutzer aus der Liste."
        MinChars = "Mindestens 3 Zeichen eingeben..."
        Ready = "Bereit"
        
        # Status
        UsersLoaded = "{0} Benutzer geladen"
        UsersFoundCount = "{0} Benutzer gefunden"
        
        # Dynamic Messages
        LoadingUsers = "Lade Benutzer aus Active Directory..."
        SearchingFor = "Suche nach '{0}'..."
        CheckingADSync = "Prüfe {0} Benutzer auf ADSync-Kompatibilität..."
        ADSyncComplete = "ADSync-Prüfung abgeschlossen: {0}/{1} kompatibel"
        ProxyAddressAdded = "ProxyAddress erfolgreich hinzugefügt"
        ProxyAddressDeleted = "ProxyAddress erfolgreich gelöscht"
        ProxyAddressUpdated = "ProxyAddress erfolgreich geändert"
        PrimarySMTPSet = "Primäre SMTP-Adresse erfolgreich gesetzt"
        UserRefreshed = "Benutzer aktualisiert"
        ExportComplete = "Export abgeschlossen: {0}"
        FilterApplied = "Filter angewendet: {0} Benutzer"
        HistoryLoaded = "History geladen für {0}"
        StatisticsShown = "Statistik angezeigt"
        ErrorOccurred = "Fehler: {0}"
        
        # Dialog Titles
        StatisticsTitle = "Statistik"
        NoData = "Keine Daten"
        NoDataMessage = "Keine Benutzer geladen. Bitte laden Sie zuerst Benutzer."
        Close = "Schließen"
        
        # Validierung / Anzeige
        ProblemsFound = "❌ Probleme gefunden"
        ErrorsHeader = "FEHLER:"
        WarningsHeader = "WARNUNGEN:"
        NoProxyAddresses = "Keine ProxyAddresses vorhanden.`nFügen Sie eine hinzu."
        FilterCount = "{0} von {1} Benutzer ({2})"
        Yes = "Ja"
        No = "Nein"
    }
    
    EN = @{
        # Header
        Title = "easyEXCH ProxyAddress Manager"
        Subtitle = "Management of ProxyMailAddresses for MS365 ADSync"
        
        # Buttons
        Search = "🔍 Search"
        LoadAll = "👥 Load All Users"
        CheckADSync = "🔍 Check ADSync"
        Export = "📤 Export CSV"
        History = "📜 History"
        Statistics = "📊 Statistics"
        Add = "➕ Add"
        Edit = "✏️ Edit"
        Delete = "🗑️ Delete"
        SetPrimary = "⭐ Set Primary"
        Refresh = "🔄 Refresh"
        
        # GroupBox Headers
        CurrentProxyAddresses = "Current ProxyAddresses"
        SearchUsers = "Search Users"
        LoadUsers = "All Users"
        FoundUsers = "Found Users"
        
        # Labels
        SearchLabel = "Search term (min. 3 characters - First name, Last name, Email, SAM):"
        UserInfo = "User Information"
        ValidationStatus = "Validation Status"
        ValidationTitle = "⚠️ MS365 ADSync Validation"
        Compatible = "✅ MS365 Compatible"
        NotCompatible = "❌ Not compatible"
        
        # DataGrid Columns - ProxyAddresses
        ColType = "Type"
        ColAddress = "Address"
        ColPrimary = "Primary"
        ColStatus = "Status"
        
        # DataGrid Columns - Users
        ColGivenName = "First Name"
        ColSurname = "Last Name"
        ColDisplayName = "Display Name"
        ColPrimarySMTP = "Primary SMTP"
        
        # Tooltips
        TooltipAdd = "Add"
        TooltipEdit = "Edit"
        TooltipDelete = "Delete"
        TooltipSetPrimary = "Set as Primary"
        TooltipRefresh = "Refresh"
        TooltipGerman = "Deutsch"
        TooltipEnglish = "English"
        TooltipADSync = "Checks if all loaded users are ADSync compatible"
        TooltipSearch = "Automatic search from 3 characters"
        
        # Filter
        FilterAll = "🔍 Show All"
        FilterValid = "✅ Compatible Only"
        FilterWarnings = "⚠️ With Warnings"
        FilterErrors = "❌ With Errors"
        
        # Messages
        NoUsersLoaded = "No users loaded.`nPlease search or load all."
        NoUsersFound = "No users found for '{0}'.`nPlease change your search term."
        NoUserSelected = "No user selected.`nPlease select a user from the list."
        MinChars = "Enter at least 3 characters..."
        Ready = "Ready"
        
        # Status
        UsersLoaded = "{0} users loaded"
        UsersFoundCount = "{0} users found"
        
        # Dynamic Messages
        LoadingUsers = "Loading users from Active Directory..."
        SearchingFor = "Searching for '{0}'..."
        CheckingADSync = "Checking {0} users for ADSync compatibility..."
        ADSyncComplete = "ADSync check complete: {0}/{1} compatible"
        ProxyAddressAdded = "ProxyAddress successfully added"
        ProxyAddressDeleted = "ProxyAddress successfully deleted"
        ProxyAddressUpdated = "ProxyAddress successfully updated"
        PrimarySMTPSet = "Primary SMTP address successfully set"
        UserRefreshed = "User refreshed"
        ExportComplete = "Export complete: {0}"
        FilterApplied = "Filter applied: {0} users"
        HistoryLoaded = "History loaded for {0}"
        StatisticsShown = "Statistics displayed"
        ErrorOccurred = "Error: {0}"
        
        # Dialog Titles
        StatisticsTitle = "Statistics"
        NoData = "No Data"
        NoDataMessage = "No users loaded. Please load users first."
        Close = "Close"
        
        # Validation / Display
        ProblemsFound = "❌ Issues found"
        ErrorsHeader = "ERRORS:"
        WarningsHeader = "WARNINGS:"
        NoProxyAddresses = "No ProxyAddresses present.`nPlease add one."
        FilterCount = "{0} of {1} users ({2})"
        Yes = "Yes"
        No = "No"
    }
}

# Funktion zum Abrufen eines übersetzten Strings
function Get-LocalizedString {
    param(
        [string]$Key,
        $FormatArgs  # Nicht typisiert, damit es flexibel ist ($Args ist eine Automatik-Variable)
    )
    
    try {
        $text = $script:LanguageStrings[$script:CurrentLanguage][$Key]
        
        if ($null -eq $text) {
            Write-Log -Message "Übersetzung nicht gefunden für Key: $Key" -Level WARNING -Source "Language"
            return $Key
        }
        
        # Wenn keine Args übergeben wurden, Text direkt zurückgeben
        if ($null -eq $FormatArgs) {
            return $text
        }
        
        # Args in Array konvertieren wenn nötig
        $argsArray = if ($FormatArgs -is [Array]) { $FormatArgs } else { @($FormatArgs) }
        
        # String formatieren
        try {
            return $text -f $argsArray
        } catch {
            Write-Log -Message "Fehler beim Formatieren von '$Key': $($_.Exception.Message)" -Level ERROR -Source "Language"
            return $text
        }
    } catch {
        Write-Log -Message "Fehler in Get-LocalizedString für Key '$Key': $($_.Exception.Message)" -Level ERROR -Source "Language"
        return $Key
    }
}

# Funktion zum Umschalten der Sprache
function Switch-Language {
    param([string]$Language)
    
    $script:CurrentLanguage = $Language
    
    # UI-Texte aktualisieren
    Update-UILanguage
    
    # Angezeigten Benutzer (Validierungsstatus etc.) neu darstellen
    if ($script:CurrentUser) {
        Update-ProxyAddressDisplay
    }
    
    # Button-Status aktualisieren
    if ($Language -eq "DE") {
        $btnLangDE.Background = "#28A745"
        $btnLangEN.Background = "#6C757D"
    } else {
        $btnLangDE.Background = "#6C757D"
        $btnLangEN.Background = "#28A745"
    }
    
    Write-Log -Message "Sprache gewechselt zu: $Language" -Level INFO -Source "Language"
}

# Funktion zum Aktualisieren der UI-Texte
function Update-UILanguage {
    # Header
    $txtTitle.Text = Get-LocalizedString "Title"
    $txtSubtitle.Text = Get-LocalizedString "Subtitle"
    
    # GroupBox Headers
    $grpProxyAddresses.Header = Get-LocalizedString "CurrentProxyAddresses"
    $grpSearchUsers.Header = Get-LocalizedString "SearchUsers"
    $grpLoadUsers.Header = Get-LocalizedString "LoadUsers"
    $grpFoundUsers.Header = Get-LocalizedString "FoundUsers"
    
    # Labels
    $lblSearchTerm.Text = Get-LocalizedString "SearchLabel"
    $lblValidationTitle.Text = Get-LocalizedString "ValidationTitle"
    
    # Messages
    $txtNoProxyData.Text = Get-LocalizedString "NoUserSelected"
    if ($txtNoUsers.Visibility -eq 'Visible') {
        $txtNoUsers.Text = Get-LocalizedString "NoUsersLoaded"
    }
    
    # Buttons
    $btnSearch.Content = Get-LocalizedString "Search"
    $btnLoadAll.Content = Get-LocalizedString "LoadAll"
    $btnCheckADSync.Content = Get-LocalizedString "CheckADSync"
    $btnExport.Content = Get-LocalizedString "Export"
    $btnHistory.Content = Get-LocalizedString "History"
    $btnStatistics.Content = Get-LocalizedString "Statistics"
    $btnAddProxy.Content = Get-LocalizedString "Add"
    $btnEditProxy.Content = Get-LocalizedString "Edit"
    $btnDeleteProxy.Content = Get-LocalizedString "Delete"
    $btnSetPrimary.Content = Get-LocalizedString "SetPrimary"
    $btnRefresh.Content = Get-LocalizedString "Refresh"
    
    # Tooltips
    $btnAddProxy.ToolTip = Get-LocalizedString "TooltipAdd"
    $btnEditProxy.ToolTip = Get-LocalizedString "TooltipEdit"
    $btnDeleteProxy.ToolTip = Get-LocalizedString "TooltipDelete"
    $btnSetPrimary.ToolTip = Get-LocalizedString "TooltipSetPrimary"
    $btnRefresh.ToolTip = Get-LocalizedString "TooltipRefresh"
    $btnLangDE.ToolTip = Get-LocalizedString "TooltipGerman"
    $btnLangEN.ToolTip = Get-LocalizedString "TooltipEnglish"
    $btnCheckADSync.ToolTip = Get-LocalizedString "TooltipADSync"
    $txtSearch.ToolTip = Get-LocalizedString "TooltipSearch"
    
    # DataGrid Columns - ProxyAddresses
    $dgProxyAddresses.Columns[0].Header = Get-LocalizedString "ColType"
    $dgProxyAddresses.Columns[1].Header = Get-LocalizedString "ColAddress"
    $dgProxyAddresses.Columns[2].Header = Get-LocalizedString "ColPrimary"
    $dgProxyAddresses.Columns[3].Header = Get-LocalizedString "ColStatus"
    
    # DataGrid Columns - Users
    $dgUsers.Columns[0].Header = Get-LocalizedString "ColGivenName"
    $dgUsers.Columns[1].Header = Get-LocalizedString "ColSurname"
    $dgUsers.Columns[2].Header = Get-LocalizedString "ColDisplayName"
    $dgUsers.Columns[3].Header = Get-LocalizedString "ColPrimarySMTP"
    
    # Filter
    $cmbFilter.Items[0].Content = Get-LocalizedString "FilterAll"
    $cmbFilter.Items[1].Content = Get-LocalizedString "FilterValid"
    $cmbFilter.Items[2].Content = Get-LocalizedString "FilterWarnings"
    $cmbFilter.Items[3].Content = Get-LocalizedString "FilterErrors"
    
    # Status
    if ($txtStatus.Text -eq "Bereit" -or $txtStatus.Text -eq "Ready") {
        $txtStatus.Text = Get-LocalizedString "Ready"
    }
}

#endregion Language Resources

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
$script:DirectoryDomainsCache = $null
$script:FilteredUsers = @()
$script:CurrentFilter = "Alle"
$script:ValidationCache = @{}
$script:SearchTimer = $null
$script:ExportFolder = Join-Path $script:ScriptPath "Exports"
$script:HistoryFolder = Join-Path $script:ScriptPath "History"

# Einheitliche AD-Attributliste für alle Benutzerabfragen
$script:ADUserProperties = @(
    'DisplayName', 'GivenName', 'Surname', 'Mail', 'SamAccountName', 'UserPrincipalName',
    'ProxyAddresses', 'msExchRecipientTypeDetails', 'msExchRecipientDisplayType',
    'mailNickname', 'targetAddress', 'legacyExchangeDN', 'DistinguishedName'
)

# Logging-Konfiguration
$script:CurrentExecutingUser = $env:USERNAME
$script:CurrentDate = Get-Date -Format "yyyy-MM-dd"
$script:LogRootFolder = Join-Path $script:ScriptPath "Logs"
$script:UserLogFolder = Join-Path $script:LogRootFolder $script:CurrentExecutingUser
$script:LogFile = Join-Path $script:UserLogFolder "$($script:CurrentExecutingUser)_$($script:CurrentDate).log"
$script:DebugLogFile = Join-Path $script:UserLogFolder "DEBUG_$($script:CurrentExecutingUser)_$($script:CurrentDate).log"
$script:MaxLogAgeDays = 90
$script:EnableDebugLogging = $true

# Export- und History-Ordner initialisieren
function Initialize-ExportFolders {
    try {
        if (-not (Test-Path $script:ExportFolder)) {
            New-Item -Path $script:ExportFolder -ItemType Directory -Force | Out-Null
        }
        if (-not (Test-Path $script:HistoryFolder)) {
            New-Item -Path $script:HistoryFolder -ItemType Directory -Force | Out-Null
        }
        return $true
    } catch {
        return $false
    }
}

# CSV-/Formel-Injection verhindern (Werte, die in Excel als Formel interpretiert würden)
function ConvertTo-SafeCsvValue {
    param($Value)
    
    if ($null -eq $Value) { return '' }
    $text = [string]$Value
    if ($text -match '^[=+\-@\t\r]') {
        return "'" + $text
    }
    return $text
}

# Sonderzeichen für LDAP-Filter escapen (RFC 4515)
function ConvertTo-LdapFilterValue {
    param([string]$Value)
    
    $sb = New-Object System.Text.StringBuilder
    foreach ($char in $Value.ToCharArray()) {
        switch ($char) {
            '\'      { [void]$sb.Append('\5c') }
            '*'      { [void]$sb.Append('\2a') }
            '('      { [void]$sb.Append('\28') }
            ')'      { [void]$sb.Append('\29') }
            "`0"     { [void]$sb.Append('\00') }
            default  { [void]$sb.Append($char) }
        }
    }
    return $sb.ToString()
}

# Eindeutige Identität eines Benutzers (DN bevorzugt, Fallback SamAccountName)
function Get-UserIdentity {
    param([Parameter(Mandatory=$true)]$User)
    
    if ($User.DistinguishedName) { return [string]$User.DistinguishedName }
    return [string]$User.SamAccountName
}

# Aktuelle ProxyAddresses frisch aus dem AD lesen (als String-Array)
function Get-CurrentProxyAddresses {
    param([Parameter(Mandatory=$true)]$User)
    
    $adUser = Get-ADUser -Identity (Get-UserIdentity -User $User) -Properties ProxyAddresses -ErrorAction Stop
    # Komma-Operator verhindert das Entpacken (leeres/einelementiges Array bleibt ein Array)
    return ,([string[]]@($adUser.ProxyAddresses | ForEach-Object { [string]$_ }))
}

# Validierungsergebnis eines Benutzers (mit Cache für Performance)
function Get-UserValidation {
    param([Parameter(Mandatory=$true)]$User)
    
    $key = Get-UserIdentity -User $User
    if (-not $key) {
        return Test-UserMS365Compatibility -User $User
    }
    if (-not $script:ValidationCache.ContainsKey($key)) {
        $script:ValidationCache[$key] = Test-UserMS365Compatibility -User $User
    }
    return $script:ValidationCache[$key]
}

# Format einer neuen ProxyAddress prüfen - gibt Fehlermeldung oder $null zurück
function Test-NewProxyAddressFormat {
    param([string]$ProxyAddress)
    
    if ([string]::IsNullOrWhiteSpace($ProxyAddress)) {
        return "ProxyAddress darf nicht leer sein"
    }
    
    if ($ProxyAddress -notmatch '^([A-Za-z0-9]+):(.+)$') {
        return "Ungültiges Format. Erwartet: TYP:adresse (z.B. smtp:user@domain.com)"
    }
    
    $type = $matches[1]
    $address = $matches[2]
    
    switch -Regex ($type) {
        '^(smtp|sip)$' {
            if ($address -match '\s') {
                return "Die Adresse darf keine Leerzeichen enthalten"
            }
            if ($address -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$' -or $address.Length -gt 254) {
                return "Ungültige E-Mail-Adresse: $address"
            }
        }
        '^x500$' {
            if ($address -notmatch '^/o=') {
                return "Ungültige X500-Adresse. Erwartet wird ein Exchange-DN (z.B. /o=ExchangeLabs/ou=...)"
            }
        }
    }
    
    return $null
}

# Prüfen ob eine ProxyAddress bereits von einem anderen AD-Objekt verwendet wird
# (Doppelte ProxyAddresses verursachen ADSync-Fehler "AttributeValueMustBeUnique")
function Get-ProxyAddressConflict {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ProxyAddress,
        [string]$ExcludeDistinguishedName
    )
    
    # Vergleich im AD ist case-insensitive (SMTP:/smtp: derselben Adresse gelten als gleich)
    $ldapFilter = "(proxyAddresses=$(ConvertTo-LdapFilterValue -Value $ProxyAddress))"
    
    $objects = $null
    try {
        # Gesamter Forest über den Global Catalog
        $forestName = (Get-ADForest -ErrorAction Stop).Name
        $objects = @(Get-ADObject -LDAPFilter $ldapFilter -Server "$($forestName):3268" -SearchBase "" -ErrorAction Stop)
    } catch {
        Write-Log -Message "GC-Abfrage fehlgeschlagen, verwende aktuelle Domäne: $($_.Exception.Message)" -Level DEBUG -Source "ProxyAddress"
        $objects = @(Get-ADObject -LDAPFilter $ldapFilter -ErrorAction Stop)
    }
    
    $conflict = $objects | Where-Object { $_.DistinguishedName -ne $ExcludeDistinguishedName } | Select-Object -First 1
    return $conflict
}

# ProxyAddresses (und optional Mail) in EINER AD-Operation schreiben (atomar)
function Set-UserProxyAddressList {
    param(
        [Parameter(Mandatory=$true)]
        $User,
        [AllowEmptyCollection()]
        [string[]]$ProxyAddresses,
        [string]$Mail
    )
    
    $identity = Get-UserIdentity -User $User
    $replace = @{}
    $clear = @()
    
    if ($ProxyAddresses -and $ProxyAddresses.Count -gt 0) {
        $replace['proxyAddresses'] = [string[]]$ProxyAddresses
    } else {
        $clear += 'proxyAddresses'
    }
    
    if ($PSBoundParameters.ContainsKey('Mail') -and -not [string]::IsNullOrWhiteSpace($Mail)) {
        $replace['mail'] = $Mail
    }
    
    $params = @{ Identity = $identity; ErrorAction = 'Stop' }
    if ($replace.Count -gt 0) { $params['Replace'] = $replace }
    if ($clear.Count -gt 0) { $params['Clear'] = $clear }
    
    Set-ADUser @params
}

# ProxyAddresses exportieren (automatisch vor Änderungen)
function Export-ProxyAddressHistory {
    param(
        [Parameter(Mandatory=$true)]
        $User,
        [string]$Action = "Backup",
        # Optional: aktueller (frisch gelesener) Stand; Standard = zwischengespeicherter Stand
        [string[]]$ProxyAddresses
    )
    
    try {
        if (-not $PSBoundParameters.ContainsKey('ProxyAddresses')) {
            $ProxyAddresses = [string[]]@($User.ProxyAddresses)
        }
        
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $fileName = "$($User.SamAccountName)_$timestamp.csv"
        $filePath = Join-Path $script:HistoryFolder $fileName
        $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        
        $exportData = New-Object System.Collections.Generic.List[object]
        $values = if ($ProxyAddresses -and $ProxyAddresses.Count -gt 0) { $ProxyAddresses } else { @('(keine)') }
        foreach ($proxy in $values) {
            $exportData.Add([PSCustomObject]@{
                Timestamp = $now
                Action = $Action
                User = $User.SamAccountName
                DisplayName = ConvertTo-SafeCsvValue $User.DisplayName
                ProxyAddress = ConvertTo-SafeCsvValue $proxy
                ExecutedBy = $env:USERNAME
            })
        }
        
        $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
        Write-Log -Message "History-Export erstellt: $fileName" -Level INFO -Source "Export"
        return $filePath
    } catch {
        Write-Log -Message "Fehler beim History-Export: $($_.Exception.Message)" -Level ERROR -Source "Export"
        return $null
    }
}

# Statistik-Dashboard anzeigen
function Show-Statistics {
    param(
        [Parameter(Mandatory=$true)]
        $Users
    )
    
    try {
        Write-Log -Message "Erstelle Statistik-Dashboard" -Level INFO -Source "Statistics"
        
        # Statistiken berechnen
        $totalUsers = $Users.Count
        $usersWithProxy = @($Users | Where-Object { $_.ProxyAddresses }).Count
        $usersWithoutProxy = $totalUsers - $usersWithProxy
        
        $validUsers = 0
        $usersWithWarnings = 0
        $usersWithErrors = 0
        $totalProxyAddresses = 0
        $primarySMTP = 0
        $secondarySMTP = 0
        $sipAddresses = 0
        $x500Addresses = 0
        
        foreach ($user in $Users) {
            $validation = Get-UserValidation -User $user
            
            if ($validation.IsValid) {
                $validUsers++
            }
            if ($validation.Warnings.Count -gt 0) {
                $usersWithWarnings++
            }
            if ($validation.Issues.Count -gt 0) {
                $usersWithErrors++
            }
            
            if ($user.ProxyAddresses) {
                foreach ($proxy in $user.ProxyAddresses) {
                    $totalProxyAddresses++
                    if ($proxy -cmatch '^SMTP:') { $primarySMTP++ }
                    elseif ($proxy -cmatch '^smtp:') { $secondarySMTP++ }
                    elseif ($proxy -match '^SIP:') { $sipAddresses++ }
                    elseif ($proxy -match '^X500:') { $x500Addresses++ }
                }
            }
        }
        
        $averageProxy = if ($totalUsers -gt 0) { [math]::Round($totalProxyAddresses / $totalUsers, 2) } else { 0 }
        
        # Statistik-Fenster
        $statsWindow = New-Object System.Windows.Window
        $statsWindow.Title = Get-LocalizedString "StatisticsTitle"
        $statsWindow.Width = 700
        $statsWindow.Height = 600
        $statsWindow.WindowStartupLocation = 'CenterOwner'
        $statsWindow.Owner = $window
        
        $grid = New-Object System.Windows.Controls.Grid
        $grid.Margin = "20"
        
        $scrollViewer = New-Object System.Windows.Controls.ScrollViewer
        $scrollViewer.VerticalScrollBarVisibility = "Auto"
        
        $stackPanel = New-Object System.Windows.Controls.StackPanel
        
        # Titel
        $title = New-Object System.Windows.Controls.TextBlock
        $title.Text = "📊 Statistik-Dashboard"
        $title.FontSize = 24
        $title.FontWeight = "Bold"
        $title.Margin = "0,0,0,20"
        $stackPanel.Children.Add($title)
        
        # Benutzer-Statistiken
        $userStats = New-Object System.Windows.Controls.TextBlock
        $userStats.Text = @"
👥 BENUTZER-STATISTIKEN
═══════════════════════════════════════
Gesamt:                    $totalUsers
Mit ProxyAddresses:        $usersWithProxy
Ohne ProxyAddresses:       $usersWithoutProxy

✅ MS365-Kompatibel:       $validUsers
⚠️ Mit Warnungen:          $usersWithWarnings
❌ Mit Fehlern:            $usersWithErrors
"@
        $userStats.FontFamily = "Consolas"
        $userStats.FontSize = 12
        $userStats.Background = "#F0F0F0"
        $userStats.Padding = "15"
        $userStats.Margin = "0,0,0,20"
        $stackPanel.Children.Add($userStats)
        
        # ProxyAddress-Statistiken
        $proxyStats = New-Object System.Windows.Controls.TextBlock
        $proxyStats.Text = @"
📧 PROXYADDRESS-STATISTIKEN
═══════════════════════════════════════
Gesamt ProxyAddresses:     $totalProxyAddresses
Durchschnitt pro Benutzer: $averageProxy

SMTP: (Primär):            $primarySMTP
smtp: (Sekundär):          $secondarySMTP
SIP:                       $sipAddresses
X500:                      $x500Addresses
"@
        $proxyStats.FontFamily = "Consolas"
        $proxyStats.FontSize = 12
        $proxyStats.Background = "#F0F0F0"
        $proxyStats.Padding = "15"
        $proxyStats.Margin = "0,0,0,20"
        $stackPanel.Children.Add($proxyStats)
        
        # Prozentuale Verteilung
        $percentValid = if ($totalUsers -gt 0) { [math]::Round(($validUsers / $totalUsers) * 100, 1) } else { 0 }
        $percentWarnings = if ($totalUsers -gt 0) { [math]::Round(($usersWithWarnings / $totalUsers) * 100, 1) } else { 0 }
        $percentErrors = if ($totalUsers -gt 0) { [math]::Round(($usersWithErrors / $totalUsers) * 100, 1) } else { 0 }
        
        $distribution = New-Object System.Windows.Controls.TextBlock
        $distribution.Text = @"
📈 VERTEILUNG
═══════════════════════════════════════
Kompatibel:    $percentValid%
Warnungen:     $percentWarnings%
Fehler:        $percentErrors%
"@
        $distribution.FontFamily = "Consolas"
        $distribution.FontSize = 12
        $distribution.Background = "#F0F0F0"
        $distribution.Padding = "15"
        $stackPanel.Children.Add($distribution)
        
        $scrollViewer.Content = $stackPanel
        $grid.Children.Add($scrollViewer)
        $statsWindow.Content = $grid
        
        Write-Log -Message "Statistik-Dashboard angezeigt" -Level INFO -Source "Statistics"
        [void]$statsWindow.ShowDialog()
        
    } catch {
        Write-Log -Message "Fehler beim Erstellen der Statistik: $($_.Exception.Message)" -Level ERROR -Source "Statistics"
        [System.Windows.MessageBox]::Show(
            "Fehler beim Erstellen der Statistik:`n`n$($_.Exception.Message)",
            "Statistik-Fehler",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# History für Benutzer anzeigen
function Show-UserHistory {
    param(
        [Parameter(Mandatory=$true)]
        $User
    )
    
    try {
        Write-Log -Message "Lade History für $($User.SamAccountName)" -Level INFO -Source "History"
        
        # History-Dateien für diesen Benutzer suchen (exakter Namensabgleich,
        # damit z.B. "max" nicht die Dateien von "max_muster" anzeigt)
        $historyPattern = '^' + [regex]::Escape([string]$User.SamAccountName) + '_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.csv$'
        $historyFiles = @(Get-ChildItem -Path $script:HistoryFolder -Filter "$($User.SamAccountName)_*.csv" -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match $historyPattern } |
            Sort-Object LastWriteTime -Descending)
        
        if ($historyFiles.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "Keine History-Einträge für $($User.DisplayName) gefunden.",
                "Keine History",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            return
        }
        
        # History-Fenster
        $historyWindow = New-Object System.Windows.Window
        $historyWindow.Title = "History - $($User.DisplayName)"
        $historyWindow.Width = 900
        $historyWindow.Height = 600
        $historyWindow.WindowStartupLocation = 'CenterOwner'
        $historyWindow.Owner = $window
        
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
        
        # Header
        $header = New-Object System.Windows.Controls.TextBlock
        $header.Text = "📜 Änderungsverlauf für $($User.DisplayName) ($($User.SamAccountName))`n$($historyFiles.Count) History-Einträge gefunden"
        $header.FontSize = 14
        $header.FontWeight = "Bold"
        $header.Margin = "0,0,0,10"
        [System.Windows.Controls.Grid]::SetRow($header, 0)
        $grid.Children.Add($header)
        
        # Alle History-Einträge laden und kombinieren
        $allHistory = New-Object System.Collections.Generic.List[object]
        foreach ($file in $historyFiles) {
            foreach ($entry in @(Import-Csv -Path $file.FullName -Encoding UTF8)) {
                $allHistory.Add($entry)
            }
        }
        
        # DataGrid
        $dgHistory = New-Object System.Windows.Controls.DataGrid
        $dgHistory.AutoGenerateColumns = $false
        $dgHistory.IsReadOnly = $true
        $dgHistory.AlternatingRowBackground = "LightGray"
        $dgHistory.GridLinesVisibility = "All"
        $dgHistory.CanUserSortColumns = $true
        
        $col1 = New-Object System.Windows.Controls.DataGridTextColumn
        $col1.Header = "Zeitpunkt"
        $col1.Binding = New-Object System.Windows.Data.Binding("Timestamp")
        $col1.Width = 150
        $dgHistory.Columns.Add($col1)
        
        $col2 = New-Object System.Windows.Controls.DataGridTextColumn
        $col2.Header = "Aktion"
        $col2.Binding = New-Object System.Windows.Data.Binding("Action")
        $col2.Width = 150
        $dgHistory.Columns.Add($col2)
        
        $col3 = New-Object System.Windows.Controls.DataGridTextColumn
        $col3.Header = "ProxyAddress"
        $col3.Binding = New-Object System.Windows.Data.Binding("ProxyAddress")
        $col3.Width = "*"
        $dgHistory.Columns.Add($col3)
        
        $col4 = New-Object System.Windows.Controls.DataGridTextColumn
        $col4.Header = "Ausgeführt von"
        $col4.Binding = New-Object System.Windows.Data.Binding("ExecutedBy")
        $col4.Width = 120
        $dgHistory.Columns.Add($col4)
        
        $dgHistory.ItemsSource = $allHistory
        [System.Windows.Controls.Grid]::SetRow($dgHistory, 1)
        $grid.Children.Add($dgHistory)
        
        # Schließen-Button
        $btnClose = New-Object System.Windows.Controls.Button
        $btnClose.Content = "Schließen"
        $btnClose.Width = 100
        $btnClose.Height = 30
        $btnClose.Margin = "0,10,0,0"
        $btnClose.HorizontalAlignment = "Right"
        $btnClose.Add_Click({ $historyWindow.Close() })
        [System.Windows.Controls.Grid]::SetRow($btnClose, 2)
        $grid.Children.Add($btnClose)
        
        $historyWindow.Content = $grid
        
        Write-Log -Message "History angezeigt: $($allHistory.Count) Einträge" -Level INFO -Source "History"
        [void]$historyWindow.ShowDialog()
        
    } catch {
        Write-Log -Message "Fehler beim Laden der History: $($_.Exception.Message)" -Level ERROR -Source "History"
        [System.Windows.MessageBox]::Show(
            "Fehler beim Laden der History:`n`n$($_.Exception.Message)",
            "History-Fehler",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

# Alle ProxyAddresses exportieren
function Export-AllProxyAddresses {
    param(
        [Parameter(Mandatory=$true)]
        $Users
    )
    
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $fileName = "ProxyAddresses_Export_$timestamp.csv"
        $filePath = Join-Path $script:ExportFolder $fileName
        
        $exportData = New-Object System.Collections.Generic.List[object]
        foreach ($user in $Users) {
            if ($user.ProxyAddresses) {
                foreach ($proxy in $user.ProxyAddresses) {
                    $exportData.Add([PSCustomObject]@{
                        DisplayName = ConvertTo-SafeCsvValue $user.DisplayName
                        SamAccountName = ConvertTo-SafeCsvValue $user.SamAccountName
                        Mail = ConvertTo-SafeCsvValue $user.Mail
                        ProxyAddress = ConvertTo-SafeCsvValue $proxy
                        Type = if ($proxy -match '^([A-Za-z0-9]+):') { $matches[1] } else { "Unknown" }
                        IsPrimary = if ($proxy -cmatch '^[A-Z0-9]+:') { "Ja" } else { "Nein" }
                    })
                }
            }
        }
        
        $exportData | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8
        Write-Log -Message "Export erstellt: $fileName ($($exportData.Count) Einträge)" -Level INFO -Source "Export"
        
        [System.Windows.MessageBox]::Show(
            "Export erfolgreich erstellt:`n`n$filePath`n`n$($exportData.Count) ProxyAddresses exportiert",
            "Export erfolgreich",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
        
        return $filePath
    } catch {
        Write-Log -Message "Fehler beim Export: $($_.Exception.Message)" -Level ERROR -Source "Export"
        [System.Windows.MessageBox]::Show(
            "Fehler beim Export:`n`n$($_.Exception.Message)",
            "Export-Fehler",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return $null
    }
}

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
Script Version: $script:AppVersion
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
        
        Write-Status -Message (Get-LocalizedString "LoadingUsers") -Level INFO
        
        # Nur aktive Benutzer - Filterung serverseitig (deutlich schneller als clientseitig)
        $adFilter = if ($Filter) { "($Filter) -and (Enabled -eq `$true)" } else { 'Enabled -eq $true' }
        
        $params = @{
            Filter = $adFilter
            Properties = $script:ADUserProperties
            ErrorAction = 'Stop'
        }
        
        if ($DomainController) {
            $params['Server'] = $DomainController
        }
        
        if ($SearchBase) {
            $params['SearchBase'] = $SearchBase
        }
        
        $users = @(Get-ADUser @params | Select-Object -Property $script:ADUserProperties)
        $userCount = $users.Count
        
        Write-Log -Message "AD-Benutzer-Abfrage abgeschlossen: $userCount Benutzer gefunden" -Level INFO -Source "AD-Query"
        Write-Status -Message (Get-LocalizedString "UsersLoaded" -FormatArgs @($userCount)) -Level SUCCESS
        return $users
        
    } catch {
        Write-Status -Message "Fehler beim Laden der Benutzer: $($_.Exception.Message)" -Level ERROR
        [System.Windows.MessageBox]::Show("Fehler beim Laden der Benutzer:`n`n$($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return @()
    }
}

# Benutzer suchen - KOMPLETT NEU
function Search-ADUsers {
    param(
        [string]$SearchTerm
    )
    
    try {
        Write-Log -Message "Suche nach '$SearchTerm'" -Level DEBUG -Source "Search"
        Write-Status -Message (Get-LocalizedString "SearchingFor" -FormatArgs @($SearchTerm)) -Level INFO
        
        # Wildcard für Suche
        $wildcard = "*$SearchTerm*"
        
        Write-Log -Message "Verwende Wildcard: $wildcard" -Level DEBUG -Source "Search"
        
        # Direkte Get-ADUser Suche mit -Filter
        $users = @(Get-ADUser -Filter {
            (Enabled -eq $true) -and (
                (DisplayName -like $wildcard) -or
                (GivenName -like $wildcard) -or
                (Surname -like $wildcard) -or
                (Mail -like $wildcard) -or
                (SamAccountName -like $wildcard) -or
                (UserPrincipalName -like $wildcard)
            )
        } -Properties $script:ADUserProperties -ErrorAction Stop |
        Select-Object -Property $script:ADUserProperties)
        
        $userCount = $users.Count
        Write-Log -Message "Suche abgeschlossen: $userCount Benutzer gefunden (Array mit $($users.Count) Elementen)" -Level INFO -Source "Search"
        Write-Status -Message (Get-LocalizedString "UsersFoundCount" -FormatArgs @($userCount)) -Level SUCCESS
        
        # Explizit als Array zurückgeben
        return @($users)
        
    } catch {
        Write-Log -Message "Fehler bei der Suche: $($_.Exception.Message)" -Level ERROR -Source "Search"
        Write-Status -Message "Fehler bei der Suche: $($_.Exception.Message)" -Level ERROR
        [System.Windows.MessageBox]::Show(
            "Fehler bei der Suche:`n`n$($_.Exception.Message)",
            "Suchfehler",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return @()
    }
}

# ProxyAddresses parsen
function Get-ParsedProxyAddresses {
    param(
        [Parameter(Mandatory=$true)]
        $User
    )
    
    $proxyList = New-Object System.Collections.Generic.List[object]
    
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
            $validationStatus = Test-ProxyAddressMS365Compatibility -ProxyAddress $proxy
            
            $proxyList.Add([PSCustomObject]@{
                FullAddress = $proxy
                Type = $type
                Address = $address
                IsPrimary = if ($isPrimary) { Get-LocalizedString "Yes" } else { Get-LocalizedString "No" }
                IsPrimaryFlag = $isPrimary
                ValidationStatus = $validationStatus
            })
        }
    }
    
    return ,$proxyList.ToArray()
}

# MS365 Kompatibilität prüfen
function Test-ProxyAddressMS365Compatibility {
    param(
        [string]$ProxyAddress
    )
    
    $issues = @()
    
    # SMTP-Adressen prüfen (-match ist case-insensitive: SMTP: und smtp:)
    if ($ProxyAddress -match '^smtp:(.+)$') {
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
        # 2. Primäre SMTP-Adresse vorhanden? (@() erzwingt Array - sonst liefert [0] bei
        #    genau einer Adresse nur das erste ZEICHEN des Strings)
        $primarySMTP = @($User.ProxyAddresses | Where-Object { $_ -cmatch '^SMTP:' })
        if ($primarySMTP.Count -eq 0) {
            $issues += "Keine primäre SMTP-Adresse (SMTP: Großbuchstaben)"
        }
        
        # 3. Mehrere primäre SMTP-Adressen?
        if ($primarySMTP.Count -gt 1) {
            $issues += "Mehrere primäre SMTP-Adressen gefunden"
        }
        
        # 4. Mail-Attribut stimmt mit primärer SMTP überein?
        if ($primarySMTP.Count -gt 0 -and $User.Mail) {
            $primaryAddress = ([string]$primarySMTP[0] -replace '^SMTP:', '')
            if ($primaryAddress -ne $User.Mail) {
                $warnings += "Mail-Attribut ($($User.Mail)) stimmt nicht mit primärer SMTP-Adresse ($primaryAddress) überein"
            }
        }
        
        # 4a. Ungültige SMTP-Adressen (werden von ADSync/Exchange Online abgelehnt)
        # 4b. Doppelte Einträge (case-insensitive)
        $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $duplicates = @()
        foreach ($proxy in $User.ProxyAddresses) {
            if ($proxy -match '^smtp:' -and (Test-ProxyAddressMS365Compatibility -ProxyAddress $proxy) -like '❌*') {
                $issues += "Ungültige SMTP-Adresse: $proxy"
            }
            if (-not $seen.Add([string]$proxy)) {
                $duplicates += [string]$proxy
            }
        }
        if ($duplicates.Count -gt 0) {
            $issues += "Doppelte ProxyAddresses: " + ($duplicates -join ', ')
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

# AD-Domänen auslesen (nur öffentliche TLDs)
function Get-ADAcceptedDomains {
    # Liste der lokalen/internen TLDs die gefiltert werden sollen
    # (echte öffentliche TLDs wie .dev, .prod, .site, .network, .company, .office, .exchange, .ad
    #  werden bewusst NICHT gefiltert)
    $localTLDs = @(
        # Deutsch
        'local', 'loc', 'lokal', 'intern', 'internal', 'int', 'lan', 'intranet',
        
        # Englisch
        'localhost', 'localdomain', 'home', 'private', 'corp', 'corporate',
        'domain', 'test', 'testing', 'development',
        'production', 'stage', 'staging', 'lab', 'demo',
        
        # Weitere typische interne
        'ds', 'mail', 'smtp', 'workgroup', 'net-local', 'invalid', 'example'
    )
    
    $isPublicDomain = {
        param([string]$Name)
        if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
        if ($Name -notmatch '\.([^.]+)$') { return $false }
        return ($matches[1].ToLower() -notin $localTLDs)
    }
    
    # Forest-Domänen und UPN-Suffixe nur einmal pro Sitzung ermitteln (Performance)
    if ($null -eq $script:DirectoryDomainsCache) {
        $directoryDomains = New-Object System.Collections.Generic.List[string]
        try {
            $forest = Get-ADForest -ErrorAction Stop
            foreach ($name in @($forest.RootDomain) + @($forest.Domains) + @($forest.UPNSuffixes)) {
                $name = [string]$name
                if ((& $isPublicDomain $name) -and -not ($directoryDomains -contains $name)) {
                    $directoryDomains.Add($name)
                }
            }
            $script:DirectoryDomainsCache = $directoryDomains.ToArray()
        } catch {
            Write-Log -Message "Fehler beim Auslesen der AD-Domänen: $($_.Exception.Message)" -Level WARNING -Source "Domains"
            $script:DirectoryDomainsCache = @()
        }
    }
    
    $domains = New-Object System.Collections.Generic.List[string]
    foreach ($name in $script:DirectoryDomainsCache) { $domains.Add($name) }
    
    # Zusätzlich: Alle einzigartigen Domänen aus bestehenden SMTP-ProxyAddresses
    if ($script:AllUsers) {
        foreach ($user in $script:AllUsers) {
            foreach ($proxy in @($user.ProxyAddresses)) {
                if ([string]$proxy -match '^smtp:[^@]+@(.+)$') {
                    $domain = $matches[1]
                    if ((& $isPublicDomain $domain) -and -not ($domains -contains $domain)) {
                        $domains.Add($domain)
                    }
                }
            }
        }
    }
    
    Write-Log -Message "Öffentliche Domänen gefunden: $($domains.Count) (lokale TLDs gefiltert)" -Level DEBUG -Source "Domains"
    return ,[string[]]@($domains | Sort-Object)
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
        $formatError = Test-NewProxyAddressFormat -ProxyAddress $ProxyAddress
        if ($formatError) {
            throw $formatError
        }
        
        # Aktuellen Stand frisch aus dem AD lesen
        $currentProxies = Get-CurrentProxyAddresses -User $User
        
        # Prüfen ob bereits vorhanden (case-insensitive wie im AD)
        if ($currentProxies | Where-Object { $_ -ieq $ProxyAddress }) {
            throw "Diese ProxyAddress ist bereits vorhanden"
        }
        
        # Prüfen ob ein anderes AD-Objekt die Adresse bereits verwendet (ADSync-Konflikt)
        $conflict = Get-ProxyAddressConflict -ProxyAddress $ProxyAddress -ExcludeDistinguishedName $User.DistinguishedName
        if ($conflict) {
            throw "Diese ProxyAddress wird bereits von einem anderen AD-Objekt verwendet:`n$($conflict.DistinguishedName)"
        }
        
        $newList = New-Object System.Collections.Generic.List[string]
        $newMail = $null
        $oldPrimary = $null
        
        foreach ($proxy in $currentProxies) {
            # Bei neuer primärer SMTP-Adresse: alte primäre zu sekundärer machen
            if ($ProxyAddress -cmatch '^SMTP:' -and $proxy -cmatch '^SMTP:') {
                $oldPrimary = $proxy
                $newList.Add(($proxy -creplace '^SMTP:', 'smtp:'))
            } else {
                $newList.Add($proxy)
            }
        }
        $newList.Add($ProxyAddress)
        
        if ($ProxyAddress -cmatch '^SMTP:(.+)$') {
            $newMail = $matches[1]
        }
        
        # History-Export VOR der Änderung (aktueller AD-Stand)
        Export-ProxyAddressHistory -User $User -Action "Vor Hinzufügen" -ProxyAddresses $currentProxies | Out-Null
        
        # Änderung in einer einzigen (atomaren) AD-Operation schreiben
        if ($newMail) {
            Set-UserProxyAddressList -User $User -ProxyAddresses $newList.ToArray() -Mail $newMail
        } else {
            Set-UserProxyAddressList -User $User -ProxyAddresses $newList.ToArray()
        }
        
        # Änderung protokollieren
        $details = "ProxyAddress wurde erfolgreich hinzugefügt"
        if ($oldPrimary) {
            $details += ". Alte primäre Adresse ($oldPrimary) wurde zu sekundärer Adresse; Mail-Attribut aktualisiert"
        }
        Write-UserChangeLog -Action "ProxyAddress HINZUGEFÜGT" -TargetUser $User -ProxyAddress $ProxyAddress -OldValue $oldPrimary -Details $details
        
        Write-Status -Message (Get-LocalizedString "ProxyAddressAdded") -Level SUCCESS
        return $true
        
    } catch {
        Write-Status -Message "Fehler beim Hinzufügen: $($_.Exception.Message)" -Level ERROR
        [System.Windows.MessageBox]::Show("Fehler beim Hinzufügen der ProxyAddress:`n`n$($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        return $false
    }
}

# ProxyAddress bearbeiten (Ersetzen in einer atomaren Operation)
function Update-ProxyAddressOfUser {
    param(
        [Parameter(Mandatory=$true)]
        $User,
        [Parameter(Mandatory=$true)]
        [string]$OldProxyAddress,
        [Parameter(Mandatory=$true)]
        [string]$NewProxyAddress
    )
    
    try {
        $formatError = Test-NewProxyAddressFormat -ProxyAddress $NewProxyAddress
        if ($formatError) {
            throw $formatError
        }
        
        $currentProxies = Get-CurrentProxyAddresses -User $User
        
        if (-not ($currentProxies | Where-Object { $_ -ceq $OldProxyAddress })) {
            throw "Die ursprüngliche ProxyAddress '$OldProxyAddress' ist im AD nicht mehr vorhanden. Bitte aktualisieren."
        }
        
        # Duplikat innerhalb des Benutzers (außer der zu ersetzenden Adresse)
        if ($currentProxies | Where-Object { $_ -cne $OldProxyAddress -and $_ -ieq $NewProxyAddress }) {
            throw "Diese ProxyAddress ist bereits vorhanden"
        }
        
        # Konflikt mit anderen AD-Objekten nur prüfen, wenn sich die Adresse (nicht nur die Schreibweise) ändert
        if ($NewProxyAddress -ine $OldProxyAddress) {
            $conflict = Get-ProxyAddressConflict -ProxyAddress $NewProxyAddress -ExcludeDistinguishedName $User.DistinguishedName
            if ($conflict) {
                throw "Diese ProxyAddress wird bereits von einem anderen AD-Objekt verwendet:`n$($conflict.DistinguishedName)"
            }
        }
        
        # Primäre SMTP-Adresse würde entfernt -> bestätigen lassen
        if ($OldProxyAddress -cmatch '^SMTP:' -and $NewProxyAddress -cnotmatch '^SMTP:') {
            $result = [System.Windows.MessageBox]::Show(
                "Die primäre SMTP-Adresse wird geändert, ohne dass eine neue primäre Adresse gesetzt wird.`n`nWARNUNG: Der Benutzer hat dann keine primäre E-Mail-Adresse mehr!`n`nFortfahren?",
                "Primäre Adresse ändern?",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning
            )
            if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
                return $false
            }
        }
        
        $newList = New-Object System.Collections.Generic.List[string]
        foreach ($proxy in $currentProxies) {
            if ($proxy -ceq $OldProxyAddress) {
                $newList.Add($NewProxyAddress)
            } elseif ($NewProxyAddress -cmatch '^SMTP:' -and $proxy -cmatch '^SMTP:') {
                # Neue primäre Adresse -> bisherige primäre wird sekundär
                $newList.Add(($proxy -creplace '^SMTP:', 'smtp:'))
            } else {
                $newList.Add($proxy)
            }
        }
        
        Export-ProxyAddressHistory -User $User -Action "Vor Bearbeiten" -ProxyAddresses $currentProxies | Out-Null
        
        if ($NewProxyAddress -cmatch '^SMTP:(.+)$') {
            Set-UserProxyAddressList -User $User -ProxyAddresses $newList.ToArray() -Mail $matches[1]
        } else {
            Set-UserProxyAddressList -User $User -ProxyAddresses $newList.ToArray()
        }
        
        Write-UserChangeLog -Action "ProxyAddress BEARBEITET" -TargetUser $User -ProxyAddress $NewProxyAddress -OldValue $OldProxyAddress -NewValue $NewProxyAddress -Details "ProxyAddress wurde geändert"
        
        Write-Status -Message (Get-LocalizedString "ProxyAddressUpdated") -Level SUCCESS
        return $true
        
    } catch {
        Write-Status -Message "Fehler beim Bearbeiten: $($_.Exception.Message)" -Level ERROR
        [System.Windows.MessageBox]::Show("Fehler beim Bearbeiten der ProxyAddress:`n`n$($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
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
        # Primäre SMTP-Adresse nur nach zusätzlicher Bestätigung löschen
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
        
        $currentProxies = Get-CurrentProxyAddresses -User $User
        $newList = [string[]]@($currentProxies | Where-Object { $_ -ine $ProxyAddress })
        
        if ($newList.Count -eq $currentProxies.Count) {
            throw "Die ProxyAddress '$ProxyAddress' ist im AD nicht mehr vorhanden. Bitte aktualisieren."
        }
        
        # History-Export vor Änderung
        Export-ProxyAddressHistory -User $User -Action "Vor Löschen" -ProxyAddresses $currentProxies | Out-Null
        
        Set-UserProxyAddressList -User $User -ProxyAddresses $newList
        
        # Änderung protokollieren
        Write-UserChangeLog -Action "ProxyAddress GELÖSCHT" -TargetUser $User -ProxyAddress $ProxyAddress -Details "ProxyAddress wurde gelöscht"
        
        Write-Status -Message (Get-LocalizedString "ProxyAddressDeleted") -Level SUCCESS
        return $true
        
    } catch {
        Write-Status -Message "Fehler beim Löschen: $($_.Exception.Message)" -Level ERROR
        [System.Windows.MessageBox]::Show("Fehler beim Löschen der ProxyAddress:`n`n$($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
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
        $newPrimary = "SMTP:$emailAddress"
        
        # Aktuelle ProxyAddresses laden
        $currentProxies = Get-CurrentProxyAddresses -User $User
        
        if (-not ($currentProxies | Where-Object { $_ -ieq $ProxyAddress })) {
            throw "Die ProxyAddress '$ProxyAddress' ist im AD nicht mehr vorhanden. Bitte aktualisieren."
        }
        
        $oldPrimary = @($currentProxies | Where-Object { $_ -cmatch '^SMTP:' }) -join ', '
        
        # Neue Liste: alte primäre -> sekundär, ausgewählte -> primär
        $newList = New-Object System.Collections.Generic.List[string]
        foreach ($proxy in $currentProxies) {
            if ($proxy -ieq $ProxyAddress) {
                $newList.Add($newPrimary)
            } elseif ($proxy -cmatch '^SMTP:') {
                $newList.Add(($proxy -creplace '^SMTP:', 'smtp:'))
            } else {
                $newList.Add($proxy)
            }
        }
        
        # History-Export VOR der Änderung
        Export-ProxyAddressHistory -User $User -Action "Vor Primär-Änderung" -ProxyAddresses $currentProxies | Out-Null
        
        # ProxyAddresses und Mail-Attribut in einer atomaren Operation aktualisieren
        Set-UserProxyAddressList -User $User -ProxyAddresses $newList.ToArray() -Mail $emailAddress
        
        # Änderung protokollieren
        Write-UserChangeLog -Action "PRIMÄRE SMTP-ADRESSE GEÄNDERT" -TargetUser $User -ProxyAddress $newPrimary -OldValue $oldPrimary -NewValue $newPrimary -Details "Primäre SMTP-Adresse wurde geändert. Alte primäre Adresse wurde zu sekundärer Adresse konvertiert."
        
        Write-Status -Message (Get-LocalizedString "PrimarySMTPSet") -Level SUCCESS
        return $true
        
    } catch {
        Write-Status -Message "Fehler beim Setzen der primären Adresse: $($_.Exception.Message)" -Level ERROR
        [System.Windows.MessageBox]::Show("Fehler beim Setzen der primären Adresse:`n`n$($_.Exception.Message)", "Fehler", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
        return $false
    }
}

# Primäre SMTP-Adresse eines Benutzers ermitteln (Fallback: Mail-Attribut)
function Get-UserPrimarySmtp {
    param([Parameter(Mandatory=$true)]$User)
    
    $primarySMTP = ""
    if ($User.ProxyAddresses) {
        $primaryProxy = $User.ProxyAddresses | Where-Object { $_ -cmatch '^SMTP:' } | Select-Object -First 1
        if ($primaryProxy) {
            $primarySMTP = $primaryProxy -replace '^SMTP:', ''
        }
    }
    
    if ([string]::IsNullOrWhiteSpace($primarySMTP) -and $User.Mail) {
        $primarySMTP = $User.Mail
    }
    
    return $primarySMTP
}

# Benutzer-Liste aktualisieren
function Update-UserList {
    param(
        [Parameter(Mandatory=$true)]
        [AllowEmptyCollection()]
        $Users
    )
    
    Write-Log -Message "Update-UserList aufgerufen mit $(@($Users).Count) Benutzern" -Level DEBUG -Source "GUI"
    
    # Benutzer für DataGrid vorbereiten
    $userList = New-Object System.Collections.Generic.List[object]
    
    foreach ($user in $Users) {
        # Benutzer-Objekt für DataGrid erstellen
        $userList.Add([PSCustomObject]@{
            GivenName = $user.GivenName
            Surname = $user.Surname
            DisplayName = $user.DisplayName
            PrimarySMTP = Get-UserPrimarySmtp -User $user
            # Original-Objekt für spätere Verwendung
            OriginalUser = $user
        })
    }
    
    Write-Log -Message "UserList erstellt mit $($userList.Count) Einträgen" -Level DEBUG -Source "GUI"
    
    # GUI-Update im Dispatcher
    $dgUsers.Dispatcher.Invoke([Action]{
        $dgUsers.ItemsSource = $null
        $dgUsers.ItemsSource = $userList
    }, "Normal")
    
    Write-Log -Message "DataGrid aktualisiert" -Level DEBUG -Source "GUI"
}

# Aktuellen Filter auf die geladenen Benutzer anwenden und Liste aktualisieren
function Update-UserFilter {
    $allUsers = @($script:AllUsers)
    $selectedIndex = $cmbFilter.SelectedIndex
    $script:CurrentFilter = $cmbFilter.SelectedItem.Content
    
    switch ($selectedIndex) {
        1 { # Nur kompatible
            $script:FilteredUsers = @($allUsers | Where-Object { (Get-UserValidation -User $_).IsValid })
        }
        2 { # Mit Warnungen
            $script:FilteredUsers = @($allUsers | Where-Object { (Get-UserValidation -User $_).Warnings.Count -gt 0 })
        }
        3 { # Mit Fehlern
            $script:FilteredUsers = @($allUsers | Where-Object { (Get-UserValidation -User $_).Issues.Count -gt 0 })
        }
        default { # Alle anzeigen
            $script:FilteredUsers = $allUsers
        }
    }
    
    Update-UserList -Users $script:FilteredUsers
    
    if ($selectedIndex -gt 0) {
        $txtUserCount.Text = Get-LocalizedString "FilterCount" -FormatArgs @($script:FilteredUsers.Count, $allUsers.Count, $script:CurrentFilter)
    } else {
        $txtUserCount.Text = Get-LocalizedString "UsersLoaded" -FormatArgs @($allUsers.Count)
    }
}

# Neu geladene Benutzerliste übernehmen (Laden/Suchen)
function Set-LoadedUsers {
    param(
        [AllowEmptyCollection()]
        $Users,
        [string]$CountKey = "UsersLoaded",
        [string]$EmptyText
    )
    
    $script:AllUsers = @($Users)
    $script:ValidationCache = @{}
    $userCount = $script:AllUsers.Count
    
    if ($userCount -gt 0) {
        $txtNoUsers.Visibility = 'Collapsed'
        Update-UserFilter
        if ($cmbFilter.SelectedIndex -le 0) {
            $txtUserCount.Text = Get-LocalizedString $CountKey -FormatArgs @($userCount)
        }
    } else {
        $dgUsers.ItemsSource = $null
        $txtNoUsers.Visibility = 'Visible'
        $txtNoUsers.Text = if ($EmptyText) { $EmptyText } else { Get-LocalizedString "NoUsersLoaded" }
        $txtUserCount.Text = Get-LocalizedString $CountKey -FormatArgs @(0)
    }
    
    $hasUsers = ($userCount -gt 0)
    $btnCheckADSync.IsEnabled = $hasUsers
    $btnExport.IsEnabled = $hasUsers
    $btnStatistics.IsEnabled = $hasUsers
}

# Aktuell ausgewählten Benutzer neu aus dem AD laden
function Update-CurrentUser {
    if ($script:CurrentUser) {
        try {
            $identity = Get-UserIdentity -User $script:CurrentUser
            $refreshedUser = Get-ADUser -Identity $identity -Properties $script:ADUserProperties -ErrorAction Stop |
                Select-Object -Property $script:ADUserProperties
            
            # Cache invalidieren und Benutzer in der geladenen Liste ersetzen,
            # damit Liste, Filter, Export und Statistik den aktuellen Stand verwenden
            $script:ValidationCache.Remove($identity)
            for ($i = 0; $i -lt @($script:AllUsers).Count; $i++) {
                if ($script:AllUsers[$i].DistinguishedName -eq $refreshedUser.DistinguishedName) {
                    $script:AllUsers[$i] = $refreshedUser
                }
            }
            
            # DataGrid-Zeile aktualisieren (Auswahl bleibt erhalten)
            if ($dgUsers.ItemsSource) {
                foreach ($item in $dgUsers.ItemsSource) {
                    if ($item.OriginalUser.DistinguishedName -eq $refreshedUser.DistinguishedName) {
                        $item.OriginalUser = $refreshedUser
                        $item.PrimarySMTP = Get-UserPrimarySmtp -User $refreshedUser
                        $item.DisplayName = $refreshedUser.DisplayName
                    }
                }
                $dgUsers.Items.Refresh()
            }
            
            $script:CurrentUser = $refreshedUser
            Update-ProxyAddressDisplay
            return $true
            
        } catch {
            Write-Status -Message "Fehler beim Aktualisieren: $($_.Exception.Message)" -Level ERROR
        }
    }
    return $false
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
    $validation = Get-UserValidation -User $script:CurrentUser
    
    $messages = @()
    if ($validation.Issues.Count -gt 0) {
        $messages += Get-LocalizedString "ErrorsHeader"
        $messages += $validation.Issues | ForEach-Object { "  • $_" }
    }
    if ($validation.Warnings.Count -gt 0) {
        $messages += Get-LocalizedString "WarningsHeader"
        $messages += $validation.Warnings | ForEach-Object { "  • $_" }
    }
    
    if ($validation.IsValid) {
        $txtValidationStatus.Text = Get-LocalizedString "Compatible"
        $txtValidationStatus.Foreground = "Green"
    } else {
        $txtValidationStatus.Text = Get-LocalizedString "ProblemsFound"
        $txtValidationStatus.Foreground = "Red"
    }
    
    # Warnungen auch bei kompatiblen Benutzern anzeigen
    if ($messages.Count -gt 0) {
        $txtValidationMessages.Text = $messages -join "`n"
        $pnlValidationInfo.Visibility = 'Visible'
    } else {
        $txtValidationMessages.Text = ""
        $pnlValidationInfo.Visibility = 'Collapsed'
    }
    
    # ProxyAddresses laden
    $proxyList = Get-ParsedProxyAddresses -User $script:CurrentUser
    
    if ($proxyList.Count -gt 0) {
        $dgProxyAddresses.ItemsSource = $proxyList
        $txtNoProxyData.Visibility = 'Collapsed'
    } else {
        $dgProxyAddresses.ItemsSource = $null
        $txtNoProxyData.Visibility = 'Visible'
        $txtNoProxyData.Text = Get-LocalizedString "NoProxyAddresses"
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
# Header
$txtTitle = $window.FindName("txtTitle")
$txtSubtitle = $window.FindName("txtSubtitle")
$btnLangDE = $window.FindName("btnLangDE")
$btnLangEN = $window.FindName("btnLangEN")

# GroupBoxes
$grpProxyAddresses = $window.FindName("grpProxyAddresses")
$grpSearchUsers = $window.FindName("grpSearchUsers")
$grpLoadUsers = $window.FindName("grpLoadUsers")
$grpFoundUsers = $window.FindName("grpFoundUsers")

# Labels
$lblSearchTerm = $window.FindName("lblSearchTerm")
$lblValidationTitle = $window.FindName("lblValidationTitle")

# TextBlocks
$txtValidationStatus = $window.FindName("txtValidationStatus")
$txtNoProxyData = $window.FindName("txtNoProxyData")

# Search and Load
$txtSearch = $window.FindName("txtSearch")
$btnSearch = $window.FindName("btnSearch")
$btnLoadAll = $window.FindName("btnLoadAll")
$btnCheckADSync = $window.FindName("btnCheckADSync")
$dgUsers = $window.FindName("dgUsers")
$txtUserCount = $window.FindName("txtUserCount")
$txtNoUsers = $window.FindName("txtNoUsers")

# Footer Buttons
$btnExport = $window.FindName("btnExport")
$cmbFilter = $window.FindName("cmbFilter")
$btnHistory = $window.FindName("btnHistory")
$btnStatistics = $window.FindName("btnStatistics")

$pnlUserInfo = $window.FindName("pnlUserInfo")
$txtSelectedUserName = $window.FindName("txtSelectedUserName")
$txtSelectedUserMail = $window.FindName("txtSelectedUserMail")
$txtSelectedUserSAM = $window.FindName("txtSelectedUserSAM")
$txtValidationStatus = $window.FindName("txtValidationStatus")

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

# Export- und History-Ordner initialisieren
Initialize-ExportFolders | Out-Null

#region Event Handlers

# ADSync-Prüfung für alle geladenen Benutzer
$btnCheckADSync.Add_Click({
    if (-not $script:AllUsers -or $script:AllUsers.Count -eq 0) {
        return
    }
    
    try {
        Write-Status -Message (Get-LocalizedString "CheckingADSync" -FormatArgs @($script:AllUsers.Count)) -Level INFO
        Write-Log -Message "Starte ADSync-Prüfung für $($script:AllUsers.Count) Benutzer" -Level INFO -Source "ADSync-Check"
        
        $totalUsers = $script:AllUsers.Count
        $validUsers = 0
        $usersWithIssues = 0
        $usersWithWarnings = 0
        $detailedResults = New-Object System.Collections.Generic.List[object]
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
        
        foreach ($user in $script:AllUsers) {
            $validation = Get-UserValidation -User $user
            
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
            
            $detailedResults.Add([PSCustomObject]@{
                DisplayName = $user.DisplayName
                SamAccountName = $user.SamAccountName
                Status = $status
                Fehler = if ($validation.Issues.Count -gt 0) { $validation.Issues -join "; " } else { "-" }
                Warnungen = if ($validation.Warnings.Count -gt 0) { $validation.Warnings -join "; " } else { "-" }
            })
        }
        [System.Windows.Input.Mouse]::OverrideCursor = $null
        
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
        Write-Status -Message (Get-LocalizedString "ADSyncComplete" -FormatArgs @($validUsers,$totalUsers)) -Level SUCCESS
        
        [void]$resultWindow.ShowDialog()
        
    } catch {
        Write-Status -Message "Fehler bei ADSync-Prüfung: $($_.Exception.Message)" -Level ERROR
        Write-Log -Message "Fehler bei ADSync-Prüfung: $($_.Exception.Message)" -Level ERROR -Source "ADSync-Check"
    } finally {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
    }
})

# Alle Benutzer laden
$btnLoadAll.Add_Click({
    try {
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
        $users = @(Get-ADUsersWithProxy -DomainController $null -SearchBase $null)
        Set-LoadedUsers -Users $users -CountKey "UsersLoaded"
    } catch {
        Write-Status -Message "Fehler: $($_.Exception.Message)" -Level ERROR
    } finally {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
    }
})

# Benutzer suchen (Funktion für Button, Enter-Taste und Live-Suche)
$performSearch = {
    try {
        if ($script:SearchTimer) {
            $script:SearchTimer.Stop()
        }
        
        $searchTerm = $txtSearch.Text.Trim()
        
        # Mindestens 3 Zeichen erforderlich
        if ($searchTerm.Length -lt 3) {
            if ($searchTerm.Length -eq 0) {
                # Leeres Suchfeld - Liste leeren
                Set-LoadedUsers -Users @() -CountKey "UsersLoaded"
            } else {
                # Weniger als 3 Zeichen
                $txtUserCount.Text = Get-LocalizedString "MinChars"
            }
            return
        }
        
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
        $searchResult = @(Search-ADUsers -SearchTerm $searchTerm)
        
        Write-Log -Message "Suchergebnis: Count=$($searchResult.Count)" -Level DEBUG -Source "GUI"
        
        Set-LoadedUsers -Users $searchResult -CountKey "UsersFoundCount" -EmptyText (Get-LocalizedString "NoUsersFound" -FormatArgs @($searchTerm))
        
    } catch {
        Write-Status -Message "Fehler: $($_.Exception.Message)" -Level ERROR
    } finally {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
    }
}

# Benutzer suchen - Button Click
$btnSearch.Add_Click($performSearch)

# Live-Suche: ein einziger Timer (500 ms Verzögerung) verhindert zu viele AD-Anfragen
$script:SearchTimer = New-Object System.Windows.Threading.DispatcherTimer
$script:SearchTimer.Interval = [TimeSpan]::FromMilliseconds(500)
$script:SearchTimer.Add_Tick({
    $script:SearchTimer.Stop()
    & $performSearch
})

# Automatische Suche beim Tippen (ab 3 Zeichen)
$txtSearch.Add_TextChanged({
    $searchTerm = $txtSearch.Text.Trim()
    $script:SearchTimer.Stop()
    
    # Nur suchen wenn mindestens 3 Zeichen
    if ($searchTerm.Length -ge 3) {
        $script:SearchTimer.Start()
    } elseif ($searchTerm.Length -eq 0) {
        # Suchfeld geleert - Liste leeren
        Set-LoadedUsers -Users @() -CountKey "UsersLoaded"
    } else {
        # Weniger als 3 Zeichen
        $txtUserCount.Text = Get-LocalizedString "MinChars"
    }
})

# Enter-Taste in Suchfeld
$txtSearch.Add_KeyDown({
    param($eventSender, $e)
    if ($e.Key -eq 'Return') {
        & $performSearch
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
            $btnHistory.IsEnabled = $true
        }
    } else {
        $btnHistory.IsEnabled = $false
    }
})

# ProxyAddress ausgewählt
$dgProxyAddresses.Add_SelectionChanged({
    if ($dgProxyAddresses.SelectedItem) {
        $btnEditProxy.IsEnabled = $true
        $btnDeleteProxy.IsEnabled = $true
        
        # Nur SMTP-Adressen können primär gesetzt werden
        $selected = $dgProxyAddresses.SelectedItem
        if ($selected.Type -match '^smtp$' -and -not $selected.IsPrimaryFlag) {
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
    
    # AD-Domänen laden
    $adDomains = Get-ADAcceptedDomains
    
    # Input Dialog
    $inputDialog = New-Object System.Windows.Window
    $inputDialog.Title = "ProxyAddress hinzufügen - $($script:CurrentUser.DisplayName)"
    $inputDialog.Width = 650
    $inputDialog.Height = 200
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
    
    # Label
    $lblAddress = New-Object System.Windows.Controls.TextBlock
    $lblAddress.Text = "ProxyAddress:"
    $lblAddress.Margin = "0,0,0,5"
    [System.Windows.Controls.Grid]::SetRow($lblAddress, 0)
    $grid.Children.Add($lblAddress)
    
    # Eingabe-Zeile
    $inputPanel = New-Object System.Windows.Controls.StackPanel
    $inputPanel.Orientation = 'Horizontal'
    $inputPanel.Margin = "0,0,0,10"
    [System.Windows.Controls.Grid]::SetRow($inputPanel, 1)
    
    # Typ ComboBox
    $cmbType = New-Object System.Windows.Controls.ComboBox
    $cmbType.Height = 30
    $cmbType.Width = 120
    $cmbType.Margin = "0,0,5,0"
    $cmbType.Items.Add("SMTP: (Primär)")
    $cmbType.Items.Add("smtp: (Sekundär)")
    $cmbType.Items.Add("SIP:")
    $cmbType.Items.Add("X500:")
    $cmbType.SelectedIndex = 1
    [void]$inputPanel.Children.Add($cmbType)
    
    # Benutzername TextBox
    $txtUsername = New-Object System.Windows.Controls.TextBox
    $txtUsername.Height = 30
    $txtUsername.Width = 262
    $txtUsername.Margin = "0,0,5,0"
    $txtUsername.Text = $script:CurrentUser.SamAccountName
    $txtUsername.VerticalContentAlignment = "Center"
    $inputPanel.Children.Add($txtUsername)
    
    # @ Symbol
    $lblAt = New-Object System.Windows.Controls.TextBlock
    $lblAt.Text = "@"
    $lblAt.FontSize = 16
    $lblAt.FontWeight = "Bold"
    $lblAt.Margin = "0,0,5,0"
    $lblAt.VerticalAlignment = "Center"
    $inputPanel.Children.Add($lblAt)
    
    # Domäne ComboBox
    $cmbDomain = New-Object System.Windows.Controls.ComboBox
    $cmbDomain.Height = 30
    $cmbDomain.Width = 200
    $cmbDomain.IsEditable = $true
    $cmbDomain.VerticalContentAlignment = "Center"
    foreach ($domain in $adDomains) {
        $cmbDomain.Items.Add($domain)
    }
    if ($cmbDomain.Items.Count -gt 0) {
        $cmbDomain.SelectedIndex = 0
    }
    $inputPanel.Children.Add($cmbDomain)
    
    $grid.Children.Add($inputPanel)
    
    # Hinweis
    $lblHint = New-Object System.Windows.Controls.TextBlock
    $lblHint.Text = "ⓘ Manuelle Domänen (nicht in der Liste) müssen bestätigt werden."
    $lblHint.FontSize = 10
    $lblHint.Foreground = "Gray"
    $lblHint.TextWrapping = "Wrap"
    [System.Windows.Controls.Grid]::SetRow($lblHint, 2)
    $grid.Children.Add($lblHint)
    
    # Typ-Wechsel: X500 benötigt einen vollständigen Exchange-DN statt name@domain
    $cmbType.Add_SelectionChanged({
        $isX500 = ($cmbType.SelectedItem -like 'X500:*')
        if ($isX500) {
            $txtUsername.Text = if ($script:CurrentUser.legacyExchangeDN) { [string]$script:CurrentUser.legacyExchangeDN } else { "/o=" }
            $txtUsername.Width = 467
            $lblAt.Visibility = 'Collapsed'
            $cmbDomain.Visibility = 'Collapsed'
            $lblHint.Text = "ⓘ X500: Vollständigen Exchange-DN eingeben (z.B. Wert aus legacyExchangeDN, beginnt mit /o=)."
        } else {
            if ($txtUsername.Text -like '/o=*') {
                $txtUsername.Text = $script:CurrentUser.SamAccountName
            }
            $txtUsername.Width = 262
            $lblAt.Visibility = 'Visible'
            $cmbDomain.Visibility = 'Visible'
            $lblHint.Text = "ⓘ Manuelle Domänen (nicht in der Liste) müssen bestätigt werden."
        }
    })
    # Buttons
    $btnPanel = New-Object System.Windows.Controls.StackPanel
    $btnPanel.Orientation = 'Horizontal'
    $btnPanel.HorizontalAlignment = 'Right'
    [System.Windows.Controls.Grid]::SetRow($btnPanel, 3)
    
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
        $username = $txtUsername.Text.Trim()
        $domain = $cmbDomain.Text.Trim()
        
        if ([string]::IsNullOrWhiteSpace($username)) {
            [System.Windows.MessageBox]::Show("Bitte geben Sie einen Benutzernamen ein.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        # X500: vollständiger Exchange-DN, keine Domäne
        if ($type -eq 'X500:') {
            if (Add-ProxyAddressToUser -User $script:CurrentUser -ProxyAddress "X500:$username") {
                $inputDialog.DialogResult = $true
                $inputDialog.Close()
                Update-CurrentUser | Out-Null
            }
            return
        }
        
        if ([string]::IsNullOrWhiteSpace($domain)) {
            [System.Windows.MessageBox]::Show("Bitte geben Sie eine Domäne ein.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        # Prüfen ob Domäne in AD-Liste ist
        if ($domain -notin $adDomains) {
            $result = [System.Windows.MessageBox]::Show(
                "Die Domäne '$domain' ist nicht in den konfigurierten AD-Domänen.`n`nMöchten Sie diese Domäne trotzdem verwenden?",
                "Manuelle Domäne bestätigen",
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Question
            )
            
            if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
                return
            }
            
            Write-Log -Message "Manuelle Domäne verwendet: $domain" -Level WARNING -Source "ProxyAddress"
        }
        
        # Vollständige Adresse erstellen (SMTP:/smtp:/SIP:)
        $fullProxy = "$type$username@$domain"
        
        if (Add-ProxyAddressToUser -User $script:CurrentUser -ProxyAddress $fullProxy) {
            $inputDialog.DialogResult = $true
            $inputDialog.Close()
            Update-CurrentUser | Out-Null
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
        
        if ($newProxy -ceq $selected.FullAddress) {
            [System.Windows.MessageBox]::Show("Die Adresse wurde nicht geändert.", "Hinweis", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            $inputDialog.Close()
            return
        }
        
        # Alte Adresse in EINER atomaren AD-Operation durch die neue ersetzen
        if (Update-ProxyAddressOfUser -User $script:CurrentUser -OldProxyAddress $selected.FullAddress -NewProxyAddress $newProxy) {
            $inputDialog.DialogResult = $true
            $inputDialog.Close()
            Update-CurrentUser | Out-Null
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
            Update-CurrentUser | Out-Null
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
            Update-CurrentUser | Out-Null
        }
    }
})

# Aktualisieren
$btnRefresh.Add_Click({
    if (Update-CurrentUser) {
        Write-Status -Message (Get-LocalizedString "UserRefreshed") -Level INFO
    }
})

# Export CSV
$btnExport.Add_Click({
    if (-not $script:AllUsers -or $script:AllUsers.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Keine Benutzer geladen. Bitte laden oder suchen Sie zuerst Benutzer.",
            "Keine Daten",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
        return
    }
    
    Export-AllProxyAddresses -Users $script:AllUsers
})

# Filter ändern
$cmbFilter.Add_SelectionChanged({
    if (-not $script:AllUsers -or @($script:AllUsers).Count -eq 0) {
        return
    }
    
    try {
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
        Update-UserFilter
        Write-Log -Message "Filter geändert: $($script:CurrentFilter)" -Level INFO -Source "Filter"
        Write-Status -Message (Get-LocalizedString "FilterApplied" -FormatArgs @($script:FilteredUsers.Count)) -Level INFO
    } catch {
        Write-Status -Message "Fehler: $($_.Exception.Message)" -Level ERROR
    } finally {
        [System.Windows.Input.Mouse]::OverrideCursor = $null
    }
})

# History anzeigen
$btnHistory.Add_Click({
    if (-not $script:CurrentUser) {
        [System.Windows.MessageBox]::Show(
            "Kein Benutzer ausgewählt. Bitte wählen Sie zuerst einen Benutzer aus.",
            "Kein Benutzer",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
        return
    }
    
    Show-UserHistory -User $script:CurrentUser
})

# Statistik anzeigen
$btnStatistics.Add_Click({
    if (-not $script:AllUsers -or $script:AllUsers.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            (Get-LocalizedString "NoDataMessage"),
            (Get-LocalizedString "NoData"),
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
        return
    }
    
    Show-Statistics -Users $script:AllUsers
})

# Sprach-Schalter Event-Handler
$btnLangDE.Add_Click({
    Switch-Language -Language "DE"
})

$btnLangEN.Add_Click({
    Switch-Language -Language "EN"
})

#endregion Event Handlers

#region Window Events

# Beim Schließen des Fensters
$window.Add_Closing({
    if ($script:SearchTimer) {
        $script:SearchTimer.Stop()
    }
    
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
Write-Status -Message (Get-LocalizedString "Ready") -Level INFO
Write-Log -Message "GUI bereit - Warte auf Benutzerinteraktion" -Level INFO -Source "Startup"
[void]$window.ShowDialog()
