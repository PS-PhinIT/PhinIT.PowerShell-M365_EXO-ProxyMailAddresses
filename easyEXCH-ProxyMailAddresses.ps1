#Requires -Version 5.1
#Requires -Modules ActiveDirectory

<#
.SYNOPSIS
    easyEXCH ProxyAddress Manager - Manage ProxyMailAddresses in Active Directory
.DESCRIPTION
    A PowerShell-based tool with modern WPF GUI to manage ProxyMailAddresses in Active Directory.
    Designed for environments without a local Exchange Server but using MS365 ADSync.
.NOTES
    Author: PS-easyIT
    Version: 1.0.0
    Date: 2025-12-08
#>

[CmdletBinding()]
param()

# Add required assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Import Active Directory module
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    [System.Windows.MessageBox]::Show("Active Directory module is not installed. Please install RSAT tools.", "Missing Module", "OK", "Error")
    exit
}

Import-Module ActiveDirectory

#region XAML Definition
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="easyEXCH ProxyAddress Manager" Height="700" Width="1100"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="MinWidth" Value="100"/>
            <Setter Property="Background" Value="#FF0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#FF005A9E"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#FFB0B0B0"/>
                    <Setter Property="Foreground" Value="#FF808080"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="5"/>
            <Setter Property="BorderBrush" Value="#FF0078D4"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Padding" Value="5"/>
            <Setter Property="BorderBrush" Value="#FF0078D4"/>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="BorderBrush" Value="#FF0078D4"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#FFE0E0E0"/>
            <Setter Property="AlternatingRowBackground" Value="#FFF5F5F5"/>
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="SelectionMode" Value="Single"/>
            <Setter Property="CanUserResizeRows" Value="False"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#FF0078D4" Padding="15" Margin="0,0,0,10">
            <StackPanel>
                <TextBlock Text="easyEXCH ProxyAddress Manager" FontSize="24" FontWeight="Bold" Foreground="White"/>
                <TextBlock Text="Manage ProxyMailAddresses for MS365 ADSync environments" FontSize="12" Foreground="White" Opacity="0.9" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- User Search Section -->
        <GroupBox Grid.Row="1" Header="User Search" Padding="10" Margin="0,0,0,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" Text="Search User:" VerticalAlignment="Center" Margin="5"/>
                <TextBox Grid.Column="1" Name="txtSearch" VerticalAlignment="Center"/>
                <Button Grid.Column="2" Name="btnSearch" Content="Search" Width="100"/>
                <Button Grid.Column="3" Name="btnClear" Content="Clear" Width="100" Background="#FF6C757D"/>
            </Grid>
        </GroupBox>
        
        <!-- User Info and ProxyAddresses Section -->
        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="300"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            
            <!-- User List -->
            <GroupBox Grid.Column="0" Header="Users" Margin="0,0,5,0" Padding="5">
                <DataGrid Name="dgUsers">
                    <DataGrid.Columns>
                        <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="*"/>
                    </DataGrid.Columns>
                </DataGrid>
            </GroupBox>
            
            <!-- ProxyAddresses Management -->
            <GroupBox Grid.Column="1" Header="ProxyAddresses" Margin="5,0,0,0" Padding="10">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- User Details -->
                    <StackPanel Grid.Row="0" Margin="0,0,0,10">
                        <TextBlock Name="txtUserInfo" FontWeight="Bold" FontSize="14" Margin="5"/>
                        <TextBlock Name="txtUserEmail" Foreground="#FF0078D4" Margin="5,0,5,5"/>
                    </StackPanel>
                    
                    <!-- ProxyAddresses List -->
                    <DataGrid Grid.Row="1" Name="dgProxyAddresses" Margin="0,0,0,10">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="80"/>
                            <DataGridTextColumn Header="Address" Binding="{Binding Address}" Width="*"/>
                            <DataGridTextColumn Header="Primary" Binding="{Binding IsPrimary}" Width="80"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    
                    <!-- Action Buttons -->
                    <Grid Grid.Row="2">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        
                        <StackPanel Grid.Column="0" Orientation="Horizontal">
                            <Button Name="btnAdd" Content="Add Address" Width="120"/>
                            <Button Name="btnSetPrimary" Content="Set Primary" Width="120"/>
                            <Button Name="btnRemove" Content="Remove" Width="120" Background="#FFDC3545"/>
                        </StackPanel>
                        
                        <Button Grid.Column="1" Name="btnRefresh" Content="Refresh" Width="100" Background="#FF28A745"/>
                    </Grid>
                </Grid>
            </GroupBox>
        </Grid>
        
        <!-- Add ProxyAddress Section -->
        <GroupBox Grid.Row="3" Header="Add New ProxyAddress" Padding="10" Margin="0,10,0,0" Name="grpAddAddress" Visibility="Collapsed">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="150"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                
                <TextBlock Grid.Column="0" Text="Type:" VerticalAlignment="Center" Margin="5"/>
                <ComboBox Grid.Column="1" Name="cmbAddressType" SelectedIndex="0">
                    <ComboBoxItem Content="SMTP (Primary)"/>
                    <ComboBoxItem Content="smtp (Secondary)"/>
                    <ComboBoxItem Content="SIP"/>
                    <ComboBoxItem Content="X400"/>
                    <ComboBoxItem Content="X500"/>
                </ComboBox>
                <TextBox Grid.Column="2" Name="txtNewAddress" VerticalAlignment="Center"/>
                <Button Grid.Column="3" Name="btnSaveAddress" Content="Save" Width="100" Background="#FF28A745"/>
                <Button Grid.Column="4" Name="btnCancelAdd" Content="Cancel" Width="100" Background="#FF6C757D"/>
            </Grid>
        </GroupBox>
        
        <!-- Status Bar -->
        <Border Grid.Row="4" Background="#FFF0F0F0" Padding="10" Margin="0,10,0,0" BorderBrush="#FFD0D0D0" BorderThickness="0,1,0,0">
            <TextBlock Name="txtStatus" Text="Ready" FontSize="12"/>
        </Border>
    </Grid>
</Window>
"@
#endregion

#region Helper Functions

function Write-Status {
    param([string]$Message)
    $window.FindName("txtStatus").Text = "$Message - $(Get-Date -Format 'HH:mm:ss')"
    $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Background)
}

function Get-ADUserDetails {
    param([string]$SearchTerm)
    
    try {
        Write-Status "Searching for users..."
        
        $filter = "DisplayName -like '*$SearchTerm*' -or SamAccountName -like '*$SearchTerm*' -or Mail -like '*$SearchTerm*'"
        $users = Get-ADUser -Filter $filter -Properties DisplayName, Mail, ProxyAddresses, SamAccountName, UserPrincipalName -ResultSetSize 50
        
        Write-Status "Found $($users.Count) user(s)"
        return $users
    }
    catch {
        [System.Windows.MessageBox]::Show("Error searching users: $_", "Error", "OK", "Error")
        Write-Status "Error searching users"
        return @()
    }
}

function Get-ProxyAddressDetails {
    param([Microsoft.ActiveDirectory.Management.ADUser]$User)
    
    $proxyList = @()
    
    foreach ($proxy in $User.ProxyAddresses) {
        $type = ""
        $address = ""
        $isPrimary = "No"
        
        if ($proxy -match "^([^:]+):(.+)$") {
            $prefix = $matches[1]
            $address = $matches[2]
            
            # Determine if primary based on uppercase prefix
            if ($prefix -ceq $prefix.ToUpper()) {
                $isPrimary = "Yes"
                $type = $prefix
            }
            else {
                $type = $prefix.ToLower()
            }
        }
        
        $proxyList += [PSCustomObject]@{
            Type = $type
            Address = $address
            IsPrimary = $isPrimary
            FullAddress = $proxy
        }
    }
    
    return $proxyList | Sort-Object Type, Address
}

function Add-ProxyAddressToUser {
    param(
        [string]$Identity,
        [string]$ProxyAddress
    )
    
    try {
        $user = Get-ADUser -Identity $Identity -Properties ProxyAddresses
        
        # Check if address already exists
        if ($user.ProxyAddresses -contains $ProxyAddress) {
            [System.Windows.MessageBox]::Show("This proxy address already exists for this user.", "Duplicate Address", "OK", "Warning")
            return $false
        }
        
        # If setting as primary SMTP, demote existing primary
        if ($ProxyAddress -like "SMTP:*") {
            $newAddresses = @()
            foreach ($addr in $user.ProxyAddresses) {
                if ($addr -like "SMTP:*") {
                    # Convert existing primary to secondary
                    $newAddresses += $addr -replace "^SMTP:", "smtp:"
                }
                else {
                    $newAddresses += $addr
                }
            }
            $newAddresses += $ProxyAddress
            Set-ADUser -Identity $Identity -Replace @{ProxyAddresses=$newAddresses}
        }
        else {
            Set-ADUser -Identity $Identity -Add @{ProxyAddresses=$ProxyAddress}
        }
        
        Write-Status "Proxy address added successfully"
        return $true
    }
    catch {
        [System.Windows.MessageBox]::Show("Error adding proxy address: $_", "Error", "OK", "Error")
        Write-Status "Error adding proxy address"
        return $false
    }
}

function Remove-ProxyAddressFromUser {
    param(
        [string]$Identity,
        [string]$ProxyAddress
    )
    
    try {
        Set-ADUser -Identity $Identity -Remove @{ProxyAddresses=$ProxyAddress}
        Write-Status "Proxy address removed successfully"
        return $true
    }
    catch {
        [System.Windows.MessageBox]::Show("Error removing proxy address: $_", "Error", "OK", "Error")
        Write-Status "Error removing proxy address"
        return $false
    }
}

function Set-PrimaryProxyAddress {
    param(
        [string]$Identity,
        [string]$NewPrimaryAddress
    )
    
    try {
        $user = Get-ADUser -Identity $Identity -Properties ProxyAddresses
        $newAddresses = @()
        
        foreach ($addr in $user.ProxyAddresses) {
            if ($addr -eq $NewPrimaryAddress) {
                # Convert to primary (uppercase)
                if ($addr -match "^([^:]+):(.+)$") {
                    $newAddresses += "$($matches[1].ToUpper()):$($matches[2])"
                }
            }
            elseif ($addr -like "SMTP:*" -and $NewPrimaryAddress -notlike "SMTP:*") {
                # Keep existing primary as secondary if new primary is not SMTP
                $newAddresses += $addr -replace "^SMTP:", "smtp:"
            }
            elseif ($addr -match "^([^:]+):(.+)$") {
                # Convert any existing primary of same type to secondary
                $prefix = $matches[1]
                $address = $matches[2]
                if ($NewPrimaryAddress -match "^$prefix:") {
                    $newAddresses += "$($prefix.ToLower()):$address"
                }
                else {
                    $newAddresses += $addr
                }
            }
            else {
                $newAddresses += $addr
            }
        }
        
        Set-ADUser -Identity $Identity -Replace @{ProxyAddresses=$newAddresses}
        Write-Status "Primary address set successfully"
        return $true
    }
    catch {
        [System.Windows.MessageBox]::Show("Error setting primary address: $_", "Error", "OK", "Error")
        Write-Status "Error setting primary address"
        return $false
    }
}

function Validate-EmailAddress {
    param([string]$Email)
    
    try {
        [void][System.Net.Mail.MailAddress]::new($Email)
        return $true
    }
    catch {
        return $false
    }
}

#endregion

#region UI Event Handlers

# Load XAML
[xml]$xamlContent = $xaml
$reader = New-Object System.Xml.XmlNodeReader $xamlContent
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get UI elements
$txtSearch = $window.FindName("txtSearch")
$btnSearch = $window.FindName("btnSearch")
$btnClear = $window.FindName("btnClear")
$dgUsers = $window.FindName("dgUsers")
$dgProxyAddresses = $window.FindName("dgProxyAddresses")
$txtUserInfo = $window.FindName("txtUserInfo")
$txtUserEmail = $window.FindName("txtUserEmail")
$btnAdd = $window.FindName("btnAdd")
$btnSetPrimary = $window.FindName("btnSetPrimary")
$btnRemove = $window.FindName("btnRemove")
$btnRefresh = $window.FindName("btnRefresh")
$grpAddAddress = $window.FindName("grpAddAddress")
$cmbAddressType = $window.FindName("cmbAddressType")
$txtNewAddress = $window.FindName("txtNewAddress")
$btnSaveAddress = $window.FindName("btnSaveAddress")
$btnCancelAdd = $window.FindName("btnCancelAdd")
$txtStatus = $window.FindName("txtStatus")

# Store current user
$script:currentUser = $null

# Search button click
$btnSearch.Add_Click({
    $searchTerm = $txtSearch.Text.Trim()
    if ([string]::IsNullOrEmpty($searchTerm)) {
        [System.Windows.MessageBox]::Show("Please enter a search term.", "Search Required", "OK", "Warning")
        return
    }
    
    $users = Get-ADUserDetails -SearchTerm $searchTerm
    $dgUsers.ItemsSource = $users
    
    # Clear proxy addresses
    $dgProxyAddresses.ItemsSource = $null
    $txtUserInfo.Text = ""
    $txtUserEmail.Text = ""
    $script:currentUser = $null
    
    # Disable action buttons
    $btnAdd.IsEnabled = $false
    $btnSetPrimary.IsEnabled = $false
    $btnRemove.IsEnabled = $false
    $btnRefresh.IsEnabled = $false
})

# Clear button click
$btnClear.Add_Click({
    $txtSearch.Text = ""
    $dgUsers.ItemsSource = $null
    $dgProxyAddresses.ItemsSource = $null
    $txtUserInfo.Text = ""
    $txtUserEmail.Text = ""
    $script:currentUser = $null
    $grpAddAddress.Visibility = "Collapsed"
    
    # Disable action buttons
    $btnAdd.IsEnabled = $false
    $btnSetPrimary.IsEnabled = $false
    $btnRemove.IsEnabled = $false
    $btnRefresh.IsEnabled = $false
    
    Write-Status "Ready"
})

# User selection changed
$dgUsers.Add_SelectionChanged({
    $selectedUser = $dgUsers.SelectedItem
    if ($null -eq $selectedUser) { return }
    
    try {
        $script:currentUser = Get-ADUser -Identity $selectedUser.SamAccountName -Properties DisplayName, Mail, ProxyAddresses, SamAccountName, UserPrincipalName
        
        $txtUserInfo.Text = "User: $($script:currentUser.DisplayName) ($($script:currentUser.SamAccountName))"
        $txtUserEmail.Text = "Email: $($script:currentUser.Mail)"
        
        $proxyList = Get-ProxyAddressDetails -User $script:currentUser
        $dgProxyAddresses.ItemsSource = $proxyList
        
        # Enable action buttons
        $btnAdd.IsEnabled = $true
        $btnSetPrimary.IsEnabled = $true
        $btnRemove.IsEnabled = $true
        $btnRefresh.IsEnabled = $true
        
        Write-Status "Loaded proxy addresses for $($script:currentUser.DisplayName)"
    }
    catch {
        [System.Windows.MessageBox]::Show("Error loading user details: $_", "Error", "OK", "Error")
    }
})

# Add button click
$btnAdd.Add_Click({
    if ($null -eq $script:currentUser) { return }
    
    $grpAddAddress.Visibility = "Visible"
    $txtNewAddress.Text = ""
    $cmbAddressType.SelectedIndex = 0
    $txtNewAddress.Focus()
})

# Save new address
$btnSaveAddress.Add_Click({
    if ($null -eq $script:currentUser) { return }
    
    $address = $txtNewAddress.Text.Trim()
    if ([string]::IsNullOrEmpty($address)) {
        [System.Windows.MessageBox]::Show("Please enter an address.", "Input Required", "OK", "Warning")
        return
    }
    
    $selectedType = $cmbAddressType.SelectedItem.Content
    $prefix = ""
    
    switch ($selectedType) {
        "SMTP (Primary)" { $prefix = "SMTP" }
        "smtp (Secondary)" { $prefix = "smtp" }
        "SIP" { $prefix = "sip" }
        "X400" { $prefix = "X400" }
        "X500" { $prefix = "X500" }
    }
    
    # Validate email format for SMTP and SIP
    if (($prefix -eq "SMTP" -or $prefix -eq "smtp" -or $prefix -eq "sip") -and -not (Validate-EmailAddress -Email $address)) {
        [System.Windows.MessageBox]::Show("Please enter a valid email address format.", "Invalid Format", "OK", "Warning")
        return
    }
    
    $fullAddress = "${prefix}:${address}"
    
    $result = Add-ProxyAddressToUser -Identity $script:currentUser.SamAccountName -ProxyAddress $fullAddress
    
    if ($result) {
        $grpAddAddress.Visibility = "Collapsed"
        
        # Refresh the proxy addresses
        $script:currentUser = Get-ADUser -Identity $script:currentUser.SamAccountName -Properties DisplayName, Mail, ProxyAddresses, SamAccountName, UserPrincipalName
        $proxyList = Get-ProxyAddressDetails -User $script:currentUser
        $dgProxyAddresses.ItemsSource = $proxyList
    }
})

# Cancel add
$btnCancelAdd.Add_Click({
    $grpAddAddress.Visibility = "Collapsed"
    $txtNewAddress.Text = ""
})

# Set Primary button click
$btnSetPrimary.Add_Click({
    if ($null -eq $script:currentUser) { return }
    
    $selectedProxy = $dgProxyAddresses.SelectedItem
    if ($null -eq $selectedProxy) {
        [System.Windows.MessageBox]::Show("Please select a proxy address to set as primary.", "Selection Required", "OK", "Warning")
        return
    }
    
    if ($selectedProxy.IsPrimary -eq "Yes") {
        [System.Windows.MessageBox]::Show("This address is already the primary address.", "Already Primary", "OK", "Information")
        return
    }
    
    $confirm = [System.Windows.MessageBox]::Show(
        "Set '$($selectedProxy.FullAddress)' as primary address?",
        "Confirm Set Primary",
        "YesNo",
        "Question"
    )
    
    if ($confirm -eq "Yes") {
        $result = Set-PrimaryProxyAddress -Identity $script:currentUser.SamAccountName -NewPrimaryAddress $selectedProxy.FullAddress
        
        if ($result) {
            # Refresh the proxy addresses
            $script:currentUser = Get-ADUser -Identity $script:currentUser.SamAccountName -Properties DisplayName, Mail, ProxyAddresses, SamAccountName, UserPrincipalName
            $proxyList = Get-ProxyAddressDetails -User $script:currentUser
            $dgProxyAddresses.ItemsSource = $proxyList
        }
    }
})

# Remove button click
$btnRemove.Add_Click({
    if ($null -eq $script:currentUser) { return }
    
    $selectedProxy = $dgProxyAddresses.SelectedItem
    if ($null -eq $selectedProxy) {
        [System.Windows.MessageBox]::Show("Please select a proxy address to remove.", "Selection Required", "OK", "Warning")
        return
    }
    
    if ($selectedProxy.IsPrimary -eq "Yes" -and $selectedProxy.Type -eq "SMTP") {
        $confirm = [System.Windows.MessageBox]::Show(
            "You are about to remove the primary SMTP address. This may affect email delivery. Continue?",
            "Warning: Primary Address",
            "YesNo",
            "Warning"
        )
        if ($confirm -ne "Yes") { return }
    }
    
    $confirm = [System.Windows.MessageBox]::Show(
        "Remove proxy address '$($selectedProxy.FullAddress)'?",
        "Confirm Removal",
        "YesNo",
        "Question"
    )
    
    if ($confirm -eq "Yes") {
        $result = Remove-ProxyAddressFromUser -Identity $script:currentUser.SamAccountName -ProxyAddress $selectedProxy.FullAddress
        
        if ($result) {
            # Refresh the proxy addresses
            $script:currentUser = Get-ADUser -Identity $script:currentUser.SamAccountName -Properties DisplayName, Mail, ProxyAddresses, SamAccountName, UserPrincipalName
            $proxyList = Get-ProxyAddressDetails -User $script:currentUser
            $dgProxyAddresses.ItemsSource = $proxyList
        }
    }
})

# Refresh button click
$btnRefresh.Add_Click({
    if ($null -eq $script:currentUser) { return }
    
    try {
        $script:currentUser = Get-ADUser -Identity $script:currentUser.SamAccountName -Properties DisplayName, Mail, ProxyAddresses, SamAccountName, UserPrincipalName
        $proxyList = Get-ProxyAddressDetails -User $script:currentUser
        $dgProxyAddresses.ItemsSource = $proxyList
        Write-Status "Proxy addresses refreshed"
    }
    catch {
        [System.Windows.MessageBox]::Show("Error refreshing proxy addresses: $_", "Error", "OK", "Error")
    }
})

# Allow Enter key in search box
$txtSearch.Add_KeyDown({
    if ($_.Key -eq "Return") {
        $btnSearch.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
    }
})

# Allow Enter key in add address box
$txtNewAddress.Add_KeyDown({
    if ($_.Key -eq "Return") {
        $btnSaveAddress.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
    }
})

#endregion

# Initialize UI state
$btnAdd.IsEnabled = $false
$btnSetPrimary.IsEnabled = $false
$btnRemove.IsEnabled = $false
$btnRefresh.IsEnabled = $false

# Show window
Write-Status "Ready - Enter search criteria to begin"
$window.ShowDialog() | Out-Null
