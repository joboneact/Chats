<#
GCC_Graph_WPF_Tool.ps1.Twist.ps1
GCC Graph / Power Platform WPF Query Tool - PowerShell 5.1
 
Purpose
- Simple WPF desktop tool for delegated Microsoft Graph and Power Platform REST queries.
- Defaults target Microsoft 365 GCC standard, which uses worldwide Graph endpoints.
- Includes GCC High/DoD cloud choices for Graph endpoint testing, but Power Platform examples are aimed at GCC standard.
 
Requirements
- Windows PowerShell 5.1 on Windows.
- Run in STA mode. This script relaunches itself with -STA when possible.
- An Entra ID app registration configured as a public client / device code capable app.
- Delegated API permissions consented for the commands you run.
 
Suggested delegated permissions by command
- Get my profile: Microsoft Graph User.Read
- Unread mail: Microsoft Graph Mail.Read
- Copilot Chat API preview: Microsoft Graph Sites.Read.All, Mail.Read, People.Read.All,
  OnlineMeetingTranscript.Read.All, Chat.Read, ChannelMessage.Read.All, ExternalItem.Read.All
- Power Platform examples: Power Platform API delegated access, and tenant/admin roles as required
- Dataverse workflow query: Dataverse delegated user_impersonation for the target org URL
 
Notes
- The Power Platform Admin Center shows Last activity for environments, but the supported
  List Environments API does not currently document a last-activity field. The related command
  below is a best-effort helper that reports whether any last-activity-like field is present.
- Microsoft 365 Copilot Chat API is beta/preview and can change. It also requires a Microsoft 365 Copilot license.
#>
 
param()
 
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    if ($PSCommandPath) {
        Start-Process -FilePath powershell.exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',"`"$PSCommandPath`"")
        exit
    }
}
 
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Web
 
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
 
$script:TokenCache = @{}
$script:LastRawJson = ''
$script:LastResponseObject = $null
$script:Commands = @()
$script:CommandByName = @{}
 
function New-CommandDefinition {
    param(
        [string]$Name,
        [string]$Title,
        [string]$Kind,
        [string]$DefaultParameters,
        [string]$Notes
    )
    $cmd = [pscustomobject]@{
        Name = $Name
        Title = $Title
        Kind = $Kind
        DefaultParameters = $DefaultParameters
        Notes = $Notes
    }
    $script:Commands += $cmd
    $script:CommandByName[$Title] = $cmd
}
 
New-CommandDefinition -Name 'graph_profile' -Title 'Graph: Get my user profile' -Kind 'Graph' -DefaultParameters @'
# No parameters required.
'@ -Notes @'
GET /v1.0/me with a conservative $select list.
Required delegated permission: User.Read.
'@
 
New-CommandDefinition -Name 'graph_unread_mail' -Title 'Graph: Get top unread inbox messages' -Kind 'Graph' -DefaultParameters @'
Top=10
'@ -Notes @'
GET /v1.0/me/mailFolders/inbox/messages filtered to unread messages.
Required delegated permission: Mail.Read.
If this fails with a query/orderby error, remove the orderby portion in the command handler.
'@
 
New-CommandDefinition -Name 'copilot_chat' -Title 'Graph beta: Invoke Microsoft 365 Copilot Chat API' -Kind 'Graph beta' -DefaultParameters @'
Prompt=Summarize my unread email at a high level.
TimeZone=America/New_York
WebSearchEnabled=false
'@ -Notes @'
Creates a beta Copilot conversation and posts one synchronous chat message.
Requires Microsoft 365 Copilot license and broad delegated permissions listed in the script header.
The documented Chat API endpoint is global Graph only; this command blocks GCC High/DoD Graph endpoints.
'@
 
New-CommandDefinition -Name 'pp_list_environments' -Title 'Power Platform: List environments for user/admin' -Kind 'Power Platform API' -DefaultParameters @'
# No parameters required.
'@ -Notes @'
GET https://api.powerplatform.com/environmentmanagement/environments?api-version=2022-03-01-preview
Returns environments available to the signed-in user. Admin users may see more.
'@
 
New-CommandDefinition -Name 'pp_inactive_environments' -Title 'Power Platform: Environments inactive > 6 months - best effort' -Kind 'Power Platform API' -DefaultParameters @'
Months=6
'@ -Notes @'
Best-effort helper. It calls the supported List Environments API and looks for a last-activity-like field.
If the field is not returned, the output explains that this API response does not expose true environment last activity.
In GCC, Power Platform Inventory is currently not available, so this script does not rely on Inventory for this task.
'@
 
New-CommandDefinition -Name 'pp_list_cloud_flows' -Title 'Power Platform: List cloud flows in environment' -Kind 'Power Platform API' -DefaultParameters @'
EnvironmentId=00000000-0000-0000-0000-000000000000
Top=100
# Optional filters supported by the Power Platform API:
# OwnerId=00000000-0000-0000-0000-000000000000
# CreatedBy=00000000-0000-0000-0000-000000000000
# ModifiedOnStartDate=2026-01-01
# ModifiedOnEndDate=2026-05-06
'@ -Notes @'
GET /powerautomate/environments/{environmentId}/cloudFlows.
Requires access to the environment. Some environments without Dataverse may not return flow metadata.
'@
 
New-CommandDefinition -Name 'dataverse_my_workflows' -Title 'Dataverse: Get solution-aware workflows for current user' -Kind 'Dataverse Web API' -DefaultParameters @'
DataverseOrgUrl=https://yourorg.crm.dynamics.com
Top=100
'@ -Notes @'
Uses Dataverse WhoAmI(), then queries /workflows for category eq 5 and ownerid eq current Dataverse user.
This covers solution-aware cloud flows in the specified Dataverse environment. Microsoft documentation says managing My Flows outside Solutions is not supported with code.
'@
 
New-CommandDefinition -Name 'custom_graph' -Title 'Custom: Microsoft Graph REST request' -Kind 'Custom Graph' -DefaultParameters @'
Method=GET
RelativePath=/v1.0/me
Scopes=User.Read
Body={}
'@ -Notes @'
Runs a custom Graph request against the selected cloud's Graph base URL.
RelativePath can be a full https URL or a relative path like /v1.0/me.
Scopes is a space-separated delegated scope list without the resource prefix, for example: User.Read Mail.Read.
'@
 
function Get-Xaml {
@"
<Window xmlns=http://schemas.microsoft.com/winfx/2006/xaml/presentation
        xmlns:x=http://schemas.microsoft.com/winfx/2006/xaml
        Title="GCC Graph / Power Platform Query Tool" Height="860" Width="1260" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="2*"/>
            <RowDefinition Height="8"/>
            <RowDefinition Height="3*"/>
        </Grid.RowDefinitions>
 
        <Border Grid.Row="0" BorderBrush="#CCCCCC" BorderThickness="1" Padding="10" CornerRadius="4">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="180"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="280"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="340"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" VerticalAlignment="Center" Margin="0,0,6,0" Text="Cloud:"/>
                <ComboBox Grid.Column="1" Name="cmbCloud" SelectedIndex="0" Margin="0,0,14,0">
                    <ComboBoxItem Content="GCC Standard / Commercial"/>
                    <ComboBoxItem Content="GCC High"/>
                    <ComboBoxItem Content="DoD"/>
                </ComboBox>
                <TextBlock Grid.Column="2" VerticalAlignment="Center" Margin="0,0,6,0" Text="Tenant ID/name:"/>
                <TextBox Grid.Column="3" Name="txtTenant" Text="organizations" Margin="0,0,14,0"/>
                <TextBlock Grid.Column="4" VerticalAlignment="Center" Margin="0,0,6,0" Text="Client ID:"/>
                <TextBox Grid.Column="5" Name="txtClientId" Margin="0,0,14,0"/>
                <StackPanel Grid.Column="6" Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button Name="btnSignIn" Content="Test sign-in" Width="100" Margin="0,0,6,0"/>
                    <Button Name="btnClearTokens" Content="Clear tokens" Width="100"/>
                </StackPanel>
            </Grid>
        </Border>
 
        <Grid Grid.Row="2">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="350"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0" BorderBrush="#CCCCCC" BorderThickness="1" Padding="8" CornerRadius="4">
                <DockPanel>
                    <TextBlock DockPanel.Dock="Top" FontWeight="Bold" Text="Commands" Margin="0,0,0,6"/>
                    <ListBox Name="lstCommands" DisplayMemberPath="Title"/>
                </DockPanel>
            </Border>
            <Border Grid.Column="2" BorderBrush="#CCCCCC" BorderThickness="1" Padding="8" CornerRadius="4">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="90"/>
                    </Grid.RowDefinitions>
                    <TextBlock Grid.Row="0" FontWeight="Bold" Text="Parameters" Margin="0,0,0,6"/>
                    <TextBox Grid.Row="1" Name="txtParams" AcceptsReturn="True" AcceptsTab="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="13"/>
                    <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,8">
                        <Button Name="btnRun" Content="Run selected command" Width="160" Margin="0,0,8,0"/>
                        <Button Name="btnClearOutput" Content="Clear output" Width="100"/>
                    </StackPanel>
                    <TextBox Grid.Row="3" Name="txtCommandHelp" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
                </Grid>
            </Border>
        </Grid>
 
        <Grid Grid.Row="4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="2*"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="2*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="2*"/>
                <RowDefinition Height="8"/>
                <RowDefinition Height="1*"/>
            </Grid.RowDefinitions>
 
            <Border Grid.Column="0" Grid.Row="0" BorderBrush="#CCCCCC" BorderThickness="1" Padding="8" CornerRadius="4">
                <DockPanel>
                    <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,6">
                        <TextBlock FontWeight="Bold" Text="Output key / value pairs" VerticalAlignment="Center"/>
                        <Button Name="btnCopySelected" Content="Copy selected" Width="100" Margin="16,0,6,0"/>
                        <Button Name="btnCopyAll" Content="Copy all" Width="80" Margin="0,0,6,0"/>
                        <Button Name="btnCopyRawJson" Content="Copy raw JSON" Width="105"/>
                    </StackPanel>
                    <ListBox Name="lstOutput" SelectionMode="Extended" FontFamily="Consolas" FontSize="12"/>
                </DockPanel>
            </Border>
 
            <Border Grid.Column="2" Grid.Row="0" BorderBrush="#CCCCCC" BorderThickness="1" Padding="8" CornerRadius="4">
                <DockPanel>
                    <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,6">
                        <TextBlock FontWeight="Bold" Text="Console log" VerticalAlignment="Center"/>
                        <Button Name="btnClearLog" Content="Clear log" Width="80" Margin="16,0,0,0"/>
                    </StackPanel>
                    <TextBox Name="txtLog" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" IsReadOnly="True" FontFamily="Consolas" FontSize="12"/>
                </DockPanel>
            </Border>
 
            <Border Grid.Column="0" Grid.ColumnSpan="3" Grid.Row="2" BorderBrush="#CCCCCC" BorderThickness="1" Padding="8" CornerRadius="4">
                <DockPanel>
                    <TextBlock DockPanel.Dock="Top" FontWeight="Bold" Text="What might have caused the error / result note" Margin="0,0,0,6"/>
                    <TextBox Name="txtCause" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" IsReadOnly="True"/>
                </DockPanel>
            </Border>
        </Grid>
    </Grid>
</Window>
"@
}
 
[xml]$xaml = Get-Xaml
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
 
$cmbCloud = $window.FindName('cmbCloud')
$txtTenant = $window.FindName('txtTenant')
$txtClientId = $window.FindName('txtClientId')
$btnSignIn = $window.FindName('btnSignIn')
$btnClearTokens = $window.FindName('btnClearTokens')
$lstCommands = $window.FindName('lstCommands')
$txtParams = $window.FindName('txtParams')
$txtCommandHelp = $window.FindName('txtCommandHelp')
$btnRun = $window.FindName('btnRun')
$btnClearOutput = $window.FindName('btnClearOutput')
$lstOutput = $window.FindName('lstOutput')
$btnCopySelected = $window.FindName('btnCopySelected')
$btnCopyAll = $window.FindName('btnCopyAll')
$btnCopyRawJson = $window.FindName('btnCopyRawJson')
$txtLog = $window.FindName('txtLog')
$btnClearLog = $window.FindName('btnClearLog')
$txtCause = $window.FindName('txtCause')
 
function Add-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $txtLog.AppendText($line + [Environment]::NewLine)
    $txtLog.ScrollToEnd()
}
 
function Set-Cause {
    param([string]$Message)
    $txtCause.Text = $Message
}
 
function Clear-ResultState {
    $lstOutput.Items.Clear()
    $script:LastRawJson = ''
    $script:LastResponseObject = $null
    Set-Cause ''
}
 
function Get-CloudConfig {
    $selected = [string]$cmbCloud.Text
    try {
        if ($cmbCloud.SelectedItem -and $cmbCloud.SelectedItem.Content) {
            $selected = [string]$cmbCloud.SelectedItem.Content
        }
    } catch { }
    switch -Wildcard ($selected) {
        'GCC High*' {
            return [pscustomobject]@{
                Name = 'GCC High'
                AuthBase = 'https://login.microsoftonline.us'
                GraphBase = 'https://graph.microsoft.us'
                PowerPlatformBase = 'https://api.powerplatform.com'
                Notes = 'GCC High Graph uses login.microsoftonline.us and graph.microsoft.us. Power Platform API endpoint support may differ by service.'
            }
        }
        'DoD*' {
            return [pscustomobject]@{
                Name = 'DoD'
                AuthBase = 'https://login.microsoftonline.us'
                GraphBase = 'https://dod-graph.microsoft.us'
                PowerPlatformBase = 'https://api.powerplatform.com'
                Notes = 'DoD Graph uses login.microsoftonline.us and dod-graph.microsoft.us. Power Platform API endpoint support may differ by service.'
            }
        }
        default {
            return [pscustomobject]@{
                Name = 'GCC Standard / Commercial'
                AuthBase = 'https://login.microsoftonline.com'
                GraphBase = 'https://graph.microsoft.com'
                PowerPlatformBase = 'https://api.powerplatform.com'
                Notes = 'Microsoft 365 GCC standard continues to use worldwide Graph endpoints.'
            }
        }
    }
}
 
function Get-RequiredTextBoxValue {
    param([System.Windows.Controls.TextBox]$TextBox, [string]$Name)
    $value = $TextBox.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw "$Name is required."
    }
    return $value
}
 
function ConvertFrom-ParameterText {
    param([string]$Text)
    $hash = @{}
    $lines = $Text -split "`r?`n"
    foreach ($line in $lines) {
        $trim = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trim)) { continue }
        if ($trim.StartsWith('#')) { continue }
        $idx = $trim.IndexOf('=')
        if ($idx -lt 1) { continue }
        $key = $trim.Substring(0, $idx).Trim()
        $value = $trim.Substring($idx + 1).Trim()
        $hash[$key] = $value
    }
    return $hash
}
 
function ConvertTo-BoolValue {
    param([object]$Value, [bool]$Default = $false)
    if ($null -eq $Value) { return $Default }
    $s = ([string]$Value).Trim().ToLowerInvariant()
    if ($s -in @('true','1','yes','y','on')) { return $true }
    if ($s -in @('false','0','no','n','off')) { return $false }
    return $Default
}
 
function Join-Url {
    param([string]$Base, [string]$Path)
    if ($Path -match '^https?://') { return $Path }
    if ($Base.EndsWith('/') -and $Path.StartsWith('/')) { return $Base.TrimEnd('/') + $Path }
    if ((-not $Base.EndsWith('/')) -and (-not $Path.StartsWith('/'))) { return $Base + '/' + $Path }
    return $Base + $Path
}
 
function Get-ExceptionBody {
    param([System.Exception]$Exception)
    try {
        if ($Exception.Response -and $Exception.Response.GetResponseStream()) {
            $stream = $Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            return $reader.ReadToEnd()
        }
    } catch { }
    return ''
}
 
function Get-HttpStatusDescription {
    param([System.Exception]$Exception)
    try {
        if ($Exception.Response) {
            return ('{0} {1}' -f ([int]$Exception.Response.StatusCode), $Exception.Response.StatusDescription)
        }
    } catch { }
    return ''
}
 
function Get-LikelyCause {
    param([System.Exception]$Exception, [string]$Body)
    $status = Get-HttpStatusDescription -Exception $Exception
    $message = $Exception.Message
    $combined = ($status + ' ' + $message + ' ' + $Body).ToLowerInvariant()
 
    if ($combined -match 'aadsts700016|application.*not.*found') {
        return 'Likely cause: the Client ID is wrong for this tenant/cloud, or the app registration is not available in the selected cloud.'
    }
    if ($combined -match 'aadsts65001|consent') {
        return 'Likely cause: admin/user consent has not been granted for one or more delegated permissions requested by this command.'
    }
    if ($combined -match 'aadsts500011|resource principal') {
        return 'Likely cause: the requested resource is wrong for this cloud, or the API is not available in the tenant/cloud.'
    }
    if ($combined -match 'unauthorized|401') {
        return 'Likely cause: token is missing/expired, the wrong cloud endpoint was selected, or the app lacks delegated permission for this API.'
    }
    if ($combined -match 'forbidden|403') {
        return 'Likely cause: the signed-in user lacks the required role/license, admin consent is missing, conditional access blocked the request, or the API is not enabled in this tenant.'
    }
    if ($combined -match 'not found|404') {
        return 'Likely cause: wrong URL, wrong Graph cloud, wrong environment ID, wrong Dataverse org URL, or the resource does not exist for the signed-in user.'
    }
    if ($combined -match 'bad request|400') {
        return 'Likely cause: a query parameter/body is invalid, a beta API changed, or the API rejected an OData expression.'
    }
    if ($combined -match 'license|copilot') {
        return 'Likely cause: Copilot APIs require a Microsoft 365 Copilot license and the preview API must be available in your tenant.'
    }
    if ($combined -match 'name resolution|could not establish trust|secure channel|tls|proxy') {
        return 'Likely cause: network/proxy/TLS inspection issue, blocked endpoint, or selected cloud endpoint is unreachable from this workstation.'
    }
    return 'Likely cause: see the HTTP status and response body in the console log. Check endpoint cloud, delegated permissions, admin consent, user role/license, and parameter values.'
}
 
function Invoke-DeviceCodeTokenV2 {
    param([string]$Scope)
 
    $tenant = Get-RequiredTextBoxValue -TextBox $txtTenant -Name 'Tenant ID/name'
    $clientId = Get-RequiredTextBoxValue -TextBox $txtClientId -Name 'Client ID'
    $cloud = Get-CloudConfig
    $cacheKey = 'v2|' + $cloud.AuthBase + '|' + $tenant + '|' + $clientId + '|' + $Scope
 
    if ($script:TokenCache.ContainsKey($cacheKey)) {
        $cached = $script:TokenCache[$cacheKey]
        if ($cached.ExpiresOn -gt (Get-Date).AddMinutes(5)) {
            Add-Log "Using cached v2 token for scopes: $Scope"
            return $cached.AccessToken
        }
    }
 
    Add-Log "Starting device-code sign-in against $($cloud.AuthBase) for scopes: $Scope"
    $deviceUrl = "$($cloud.AuthBase)/$tenant/oauth2/v2.0/devicecode"
    $tokenUrl = "$($cloud.AuthBase)/$tenant/oauth2/v2.0/token"
 
    $deviceBody = @{
        client_id = $clientId
        scope = $Scope
    }
    $device = Invoke-RestMethod -Method Post -Uri $deviceUrl -Body $deviceBody -ContentType 'application/x-www-form-urlencoded'
    Add-Log $device.message 'AUTH'
    Set-Cause $device.message
 
    $interval = 5
    if ($device.interval) { $interval = [int]$device.interval }
    $expiresAt = (Get-Date).AddSeconds([int]$device.expires_in)
 
    while ((Get-Date) -lt $expiresAt) {
        Start-Sleep -Seconds $interval
        try {
            $tokenBody = @{
                grant_type = 'urn:ietf:params:oauth:grant-type:device_code'
                client_id = $clientId
                device_code = $device.device_code
            }
            $token = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $tokenBody -ContentType 'application/x-www-form-urlencoded'
            $expiresOn = (Get-Date).AddSeconds([int]$token.expires_in)
            $script:TokenCache[$cacheKey] = [pscustomobject]@{ AccessToken = $token.access_token; ExpiresOn = $expiresOn }
            Add-Log "Token acquired. Expires $expiresOn" 'AUTH'
            Set-Cause ''
            return $token.access_token
        } catch {
            $body = Get-ExceptionBody -Exception $_.Exception
            if ($body -match 'authorization_pending') { continue }
            if ($body -match 'slow_down') { $interval += 5; continue }
            throw
        }
    }
    throw 'Device-code sign-in timed out before completion.'
}
 
function Invoke-DeviceCodeTokenV1 {
    param([string]$Resource)
 
    $tenant = Get-RequiredTextBoxValue -TextBox $txtTenant -Name 'Tenant ID/name'
    $clientId = Get-RequiredTextBoxValue -TextBox $txtClientId -Name 'Client ID'
    $cloud = Get-CloudConfig
    $cacheKey = 'v1|' + $cloud.AuthBase + '|' + $tenant + '|' + $clientId + '|' + $Resource
 
    if ($script:TokenCache.ContainsKey($cacheKey)) {
        $cached = $script:TokenCache[$cacheKey]
        if ($cached.ExpiresOn -gt (Get-Date).AddMinutes(5)) {
            Add-Log "Using cached v1 token for resource: $Resource"
            return $cached.AccessToken
        }
    }
 
    Add-Log "Starting device-code sign-in against $($cloud.AuthBase) for resource: $Resource"
    $deviceUrl = "$($cloud.AuthBase)/$tenant/oauth2/devicecode"
    $tokenUrl = "$($cloud.AuthBase)/$tenant/oauth2/token"
 
    $deviceBody = @{
        client_id = $clientId
        resource = $Resource
    }
    $device = Invoke-RestMethod -Method Post -Uri $deviceUrl -Body $deviceBody -ContentType 'application/x-www-form-urlencoded'
    Add-Log $device.message 'AUTH'
    Set-Cause $device.message
 
    $interval = 5
    if ($device.interval) { $interval = [int]$device.interval }
    $expiresAt = (Get-Date).AddSeconds([int]$device.expires_in)
 
    while ((Get-Date) -lt $expiresAt) {
        Start-Sleep -Seconds $interval
        try {
            $tokenBody = @{
                grant_type = 'urn:ietf:params:oauth:grant-type:device_code'
                client_id = $clientId
                code = $device.device_code
            }
            $token = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body $tokenBody -ContentType 'application/x-www-form-urlencoded'
            $expiresOn = (Get-Date).AddSeconds([int]$token.expires_in)
            $script:TokenCache[$cacheKey] = [pscustomobject]@{ AccessToken = $token.access_token; ExpiresOn = $expiresOn }
            Add-Log "Token acquired. Expires $expiresOn" 'AUTH'
            Set-Cause ''
            return $token.access_token
        } catch {
            $body = Get-ExceptionBody -Exception $_.Exception
            if ($body -match 'authorization_pending') { continue }
            if ($body -match 'slow_down') { $interval += 5; continue }
            throw
        }
    }
    throw 'Device-code sign-in timed out before completion.'
}
 
function Get-GraphToken {
    param([string[]]$Scopes)
    $cloud = Get-CloudConfig
    $scoped = @()
    foreach ($scope in $Scopes) {
        if ([string]::IsNullOrWhiteSpace($scope)) { continue }
        if ($scope -match '^https?://') { $scoped += $scope }
        elseif ($scope -eq 'offline_access' -or $scope -eq 'openid' -or $scope -eq 'profile') { $scoped += $scope }
        else { $scoped += ($cloud.GraphBase.TrimEnd('/') + '/' + $scope) }
    }
    if ($scoped -notcontains 'offline_access') { $scoped += 'offline_access' }
    $scopeString = ($scoped -join ' ')
    return Invoke-DeviceCodeTokenV2 -Scope $scopeString
}
 
function Invoke-ApiRequest {
    param(
        [ValidateSet('GET','POST','PATCH','PUT','DELETE')][string]$Method,
        [string]$Uri,
        [string]$AccessToken,
        [object]$Body = $null,
        [string]$ContentType = 'application/json'
    )
 
    Add-Log "$Method $Uri" 'HTTP'
    $headers = @{
        Authorization = "Bearer $AccessToken"
        Accept = 'application/json'
    }
 
    try {
        if ($null -ne $Body -and $Method -ne 'GET') {
            if ($Body -is [string]) { $bodyText = $Body } else { $bodyText = ($Body | ConvertTo-Json -Depth 30) }
            Add-Log "Request body: $bodyText" 'HTTP'
            $response = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $bodyText -ContentType $ContentType
        } else {
            $response = Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
        }
        $script:LastResponseObject = $response
        if ($null -eq $response) { $script:LastRawJson = '' } else { $script:LastRawJson = ($response | ConvertTo-Json -Depth 50) }
        Add-Log 'Request succeeded.' 'HTTP'
        return $response
    } catch {
        $body = Get-ExceptionBody -Exception $_.Exception
        $status = Get-HttpStatusDescription -Exception $_.Exception
        Add-Log "Request failed. $status $($_.Exception.Message)" 'ERROR'
        if (-not [string]::IsNullOrWhiteSpace($body)) { Add-Log "Response body: $body" 'ERROR' }
        Set-Cause (Get-LikelyCause -Exception $_.Exception -Body $body)
        throw
    }
}
 
function Add-OutputLine {
    param([string]$Line)
    [void]$lstOutput.Items.Add($Line)
}
 
function Convert-ValueToString {
    param([object]$Value)
    if ($null -eq $Value) { return '<null>' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [datetime]) { return $Value.ToString('o') }
    if ($Value -is [bool] -or $Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) { return [string]$Value }
    return ($Value | ConvertTo-Json -Depth 10 -Compress)
}
 
function Write-FlattenedObject {
    param([object]$Object, [string]$Prefix = '')
 
    if ($null -eq $Object) {
        Add-OutputLine ("{0} = <null>" -f $Prefix)
        return
    }
 
    if ($Object -is [System.Array]) {
        $i = 0
        foreach ($item in $Object) {
            Write-FlattenedObject -Object $item -Prefix ("{0}[{1}]" -f $Prefix, $i)
            $i++
        }
        if ($i -eq 0) { Add-OutputLine ("{0} = []" -f $Prefix) }
        return
    }
 
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($key in $Object.Keys) {
            $newPrefix = if ($Prefix) { "$Prefix.$key" } else { [string]$key }
            Write-FlattenedObject -Object $Object[$key] -Prefix $newPrefix
        }
        return
    }
 
    $properties = @($Object.PSObject.Properties | Where-Object { $_.MemberType -match 'Property|NoteProperty|AliasProperty' })
    if ($properties.Count -gt 0 -and -not ($Object -is [string])) {
        foreach ($prop in $properties) {
            $newPrefix = if ($Prefix) { "$Prefix.$($prop.Name)" } else { $prop.Name }
            Write-FlattenedObject -Object $prop.Value -Prefix $newPrefix
        }
        return
    }
 
    Add-OutputLine ("{0} = {1}" -f $Prefix, (Convert-ValueToString $Object))
}
 
function Show-ObjectResult {
    param([object]$Object, [string]$Heading = '')
    $lstOutput.Items.Clear()
    if (-not [string]::IsNullOrWhiteSpace($Heading)) {
        Add-OutputLine $Heading
        Add-OutputLine ('-' * 80)
    }
    if ($null -eq $Object) {
        Add-OutputLine '<no response body>'
        return
    }
    Write-FlattenedObject -Object $Object
}
 
function Copy-LinesToClipboard {
    param([string[]]$Lines)
    try {
        if ($Lines.Count -gt 0) {
            [System.Windows.Clipboard]::SetText(($Lines -join [Environment]::NewLine))
            Add-Log "Copied $($Lines.Count) line(s) to clipboard."
        }
    } catch {
        Add-Log "Clipboard copy failed: $($_.Exception.Message)" 'ERROR'
        Set-Cause 'Clipboard operations require the process to run in STA mode and may be blocked by desktop/session restrictions.'
    }
}
 
function Get-OutputLines {
    $lines = @()
    foreach ($item in $lstOutput.Items) { $lines += [string]$item }
    return $lines
}
 
function Get-SelectedOutputLines {
    $lines = @()
    foreach ($item in $lstOutput.SelectedItems) { $lines += [string]$item }
    return $lines
}
 
function Normalize-OrgUrl {
    param([string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { throw 'DataverseOrgUrl is required.' }
    return $Url.Trim().TrimEnd('/')
}
 
function Get-FirstPropertyValue {
    param([object]$Object, [string[]]$Names)
    if ($null -eq $Object) { return $null }
    foreach ($name in $Names) {
        $prop = $Object.PSObject.Properties[$name]
        if ($prop -and $null -ne $prop.Value -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) { return $prop.Value }
    }
    return $null
}
 
function Run-GraphProfile {
    $cloud = Get-CloudConfig
    $token = Get-GraphToken -Scopes @('User.Read')
    $url = $cloud.GraphBase + '/v1.0/me?$select=displayName,mail,userPrincipalName,id,jobTitle,department,officeLocation,mobilePhone,businessPhones'
    $result = Invoke-ApiRequest -Method GET -Uri $url -AccessToken $token
    Show-ObjectResult -Object $result -Heading 'Current user profile'
}
 
function Run-GraphUnreadMail {
    $p = ConvertFrom-ParameterText -Text $txtParams.Text
    $top = 10
    if ($p.ContainsKey('Top')) { [void][int]::TryParse($p['Top'], [ref]$top) }
    if ($top -lt 1) { $top = 10 }
    if ($top -gt 50) { $top = 50 }
 
    $cloud = Get-CloudConfig
    $token = Get-GraphToken -Scopes @('Mail.Read')
    $query = '?$filter=isRead%20eq%20false&$top=' + $top + '&$select=subject,from,receivedDateTime,webLink,isRead&$orderby=receivedDateTime%20desc'
    $url = $cloud.GraphBase + '/v1.0/me/mailFolders/inbox/messages' + $query
    $result = Invoke-ApiRequest -Method GET -Uri $url -AccessToken $token
    Show-ObjectResult -Object $result -Heading "Top $top unread inbox messages"
}
 
function Run-CopilotChat {
    $cloud = Get-CloudConfig
    if ($cloud.GraphBase -ne 'https://graph.microsoft.com') {
        Set-Cause 'The Microsoft 365 Copilot Chat API documentation lists the create-conversation endpoint as available in Global service but not US Government L4/L5 or China. GCC standard uses global Graph; GCC High/DoD do not.'
        throw 'Copilot Chat API command is blocked for this selected cloud endpoint.'
    }
 
    $p = ConvertFrom-ParameterText -Text $txtParams.Text
    $prompt = 'Summarize my unread email at a high level.'
    if ($p.ContainsKey('Prompt')) { $prompt = $p['Prompt'] }
    $timeZone = 'America/New_York'
    if ($p.ContainsKey('TimeZone')) { $timeZone = $p['TimeZone'] }
    $webEnabled = $false
    if ($p.ContainsKey('WebSearchEnabled')) { $webEnabled = ConvertTo-BoolValue -Value $p['WebSearchEnabled'] -Default $false }
 
    $scopes = @(
        'Sites.Read.All',
        'Mail.Read',
        'People.Read.All',
        'OnlineMeetingTranscript.Read.All',
        'Chat.Read',
        'ChannelMessage.Read.All',
        'ExternalItem.Read.All'
    )
    $token = Get-GraphToken -Scopes $scopes
 
    $conversationUrl = $cloud.GraphBase + '/beta/copilot/conversations'
    $conversation = Invoke-ApiRequest -Method POST -Uri $conversationUrl -AccessToken $token -Body @{}
    $conversationId = $conversation.id
    if ([string]::IsNullOrWhiteSpace($conversationId)) { throw 'Copilot conversation was created but no id was returned.' }
    Add-Log "Copilot conversation id: $conversationId"
 
    $chatUrl = $cloud.GraphBase + '/beta/copilot/conversations/' + $conversationId + '/chat'
    $body = @{
        message = @{ text = $prompt }
        locationHint = @{ timeZone = $timeZone }
        contextualResources = @{ webContext = @{ isWebEnabled = $webEnabled } }
    }
    $result = Invoke-ApiRequest -Method POST -Uri $chatUrl -AccessToken $token -Body $body
    Show-ObjectResult -Object $result -Heading 'Copilot Chat API response'
 
    try {
        $messages = @($result.messages)
        if ($messages.Count -gt 0) {
            $last = $messages[$messages.Count - 1]
            if ($last.text) {
                Add-OutputLine ''
                Add-OutputLine 'copilot.finalText ='
                Add-OutputLine ([string]$last.text)
            }
        }
    } catch { }
}
 
function Run-PowerPlatformListEnvironments {
    $cloud = Get-CloudConfig
    $token = Invoke-DeviceCodeTokenV1 -Resource $cloud.PowerPlatformBase
    $url = $cloud.PowerPlatformBase + '/environmentmanagement/environments?api-version=2022-03-01-preview'
    $result = Invoke-ApiRequest -Method GET -Uri $url -AccessToken $token
    Show-ObjectResult -Object $result -Heading 'Power Platform environments available to signed-in user'
}
 
function Run-PowerPlatformInactiveEnvironments {
    $p = ConvertFrom-ParameterText -Text $txtParams.Text
    $months = 6
    if ($p.ContainsKey('Months')) { [void][int]::TryParse($p['Months'], [ref]$months) }
    if ($months -lt 1) { $months = 6 }
    $cutoff = (Get-Date).AddMonths(-1 * $months)
 
    $cloud = Get-CloudConfig
    $token = Invoke-DeviceCodeTokenV1 -Resource $cloud.PowerPlatformBase
    $url = $cloud.PowerPlatformBase + '/environmentmanagement/environments?api-version=2022-03-01-preview'
    $result = Invoke-ApiRequest -Method GET -Uri $url -AccessToken $token
 
    $lstOutput.Items.Clear()
    Add-OutputLine "Inactive environment best-effort check. Cutoff: $($cutoff.ToString('yyyy-MM-dd'))"
    Add-OutputLine ('-' * 80)
 
    $items = @()
    if ($result.value) { $items = @($result.value) } else { $items = @($result) }
 
    $activityFieldNames = @('lastActivityDateTime','lastActivity','lastUsedDateTime','lastUsedOn','lastAccessedDateTime','lastModifiedDateTime','modifiedDateTime')
    $foundAnyActivityField = $false
    $inactiveCount = 0
 
    foreach ($env in $items) {
        $activity = Get-FirstPropertyValue -Object $env -Names $activityFieldNames
        $name = Get-FirstPropertyValue -Object $env -Names @('displayName','name','id')
        $id = Get-FirstPropertyValue -Object $env -Names @('id','name','environmentId')
        $type = Get-FirstPropertyValue -Object $env -Names @('type')
        $state = Get-FirstPropertyValue -Object $env -Names @('state')
        $created = Get-FirstPropertyValue -Object $env -Names @('createdDateTime','createdTime','createdOn')
 
        if ($activity) {
            $foundAnyActivityField = $true
            $activityDate = $null
            if ([datetime]::TryParse([string]$activity, [ref]$activityDate)) {
                if ($activityDate -le $cutoff) {
                    $inactiveCount++
                    Add-OutputLine ("inactive[{0}].displayName = {1}" -f $inactiveCount, $name)
                    Add-OutputLine ("inactive[{0}].id = {1}" -f $inactiveCount, $id)
                    Add-OutputLine ("inactive[{0}].type = {1}" -f $inactiveCount, $type)
                    Add-OutputLine ("inactive[{0}].state = {1}" -f $inactiveCount, $state)
                    Add-OutputLine ("inactive[{0}].lastActivityObserved = {1}" -f $inactiveCount, $activity)
                    Add-OutputLine ''
                }
            }
        } else {
            Add-OutputLine ("environment.displayName = {0}" -f $name)
            Add-OutputLine ("environment.id = {0}" -f $id)
            Add-OutputLine ("environment.type = {0}" -f $type)
            Add-OutputLine ("environment.state = {0}" -f $state)
            Add-OutputLine ("environment.createdDateTime = {0}" -f $created)
            Add-OutputLine 'environment.lastActivityObserved = <not returned by supported List Environments API>'
            Add-OutputLine ''
        }
    }
 
    if ($foundAnyActivityField) {
        Add-OutputLine "inactive.count = $inactiveCount"
        if ($inactiveCount -eq 0) { Add-OutputLine 'No environments older than the cutoff were found using the activity field returned by the API.' }
    } else {
        Set-Cause 'The supported Power Platform List Environments API response did not include a documented Last activity field. In GCC, Power Platform Inventory is currently not available, so true environment inactivity usually has to be checked in PPAC Last activity or via audit/activity logging data rather than this REST response.'
    }
}
 
function Run-PowerPlatformListCloudFlows {
    $p = ConvertFrom-ParameterText -Text $txtParams.Text
    if (-not $p.ContainsKey('EnvironmentId') -or [string]::IsNullOrWhiteSpace($p['EnvironmentId']) -or $p['EnvironmentId'] -eq '00000000-0000-0000-0000-000000000000') {
        throw 'EnvironmentId is required. Run the environment list command first, then paste the id here.'
    }
 
    $cloud = Get-CloudConfig
    $token = Invoke-DeviceCodeTokenV1 -Resource $cloud.PowerPlatformBase
    $envId = [System.Web.HttpUtility]::UrlEncode($p['EnvironmentId'])
    $queryParts = New-Object System.Collections.Generic.List[string]
    $queryParts.Add('api-version=2022-03-01-preview')
 
    foreach ($key in @('OwnerId','CreatedBy','ModifiedOnStartDate','ModifiedOnEndDate')) {
        if ($p.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace($p[$key])) {
            $apiKey = $key.Substring(0,1).ToLowerInvariant() + $key.Substring(1)
            $queryParts.Add($apiKey + '=' + [System.Web.HttpUtility]::UrlEncode($p[$key]))
        }
    }
 
    $url = $cloud.PowerPlatformBase + '/powerautomate/environments/' + $envId + '/cloudFlows?' + ($queryParts -join '&')
    $result = Invoke-ApiRequest -Method GET -Uri $url -AccessToken $token
    Show-ObjectResult -Object $result -Heading 'Cloud flows in environment'
}
 
function Run-DataverseMyWorkflows {
    $p = ConvertFrom-ParameterText -Text $txtParams.Text
    if (-not $p.ContainsKey('DataverseOrgUrl')) { throw 'DataverseOrgUrl is required.' }
    $orgUrl = Normalize-OrgUrl -Url $p['DataverseOrgUrl']
    $top = 100
    if ($p.ContainsKey('Top')) { [void][int]::TryParse($p['Top'], [ref]$top) }
    if ($top -lt 1) { $top = 100 }
    if ($top -gt 500) { $top = 500 }
 
    $token = Invoke-DeviceCodeTokenV1 -Resource $orgUrl
    $who = Invoke-ApiRequest -Method GET -Uri ($orgUrl + '/api/data/v9.2/WhoAmI()') -AccessToken $token
    $userId = [string]$who.UserId
    if ([string]::IsNullOrWhiteSpace($userId)) { throw 'Dataverse WhoAmI() did not return a UserId.' }
    Add-Log "Dataverse current user id: $userId"
 
    $filter = [System.Uri]::EscapeDataString("category eq 5 and _ownerid_value eq $userId")
    $orderby = [System.Uri]::EscapeDataString('modifiedon desc')
    $url = $orgUrl + '/api/data/v9.2/workflows?$select=name,workflowid,workflowidunique,createdon,modifiedon,statecode,category,_ownerid_value&$filter=' + $filter + '&$top=' + $top + '&$orderby=' + $orderby
    $result = Invoke-ApiRequest -Method GET -Uri $url -AccessToken $token
    Show-ObjectResult -Object $result -Heading 'Solution-aware cloud flows owned by current Dataverse user'
}
 
function Run-CustomGraph {
    $p = ConvertFrom-ParameterText -Text $txtParams.Text
    $method = 'GET'
    if ($p.ContainsKey('Method')) { $method = $p['Method'].Trim().ToUpperInvariant() }
    if ($method -notin @('GET','POST','PATCH','PUT','DELETE')) { throw 'Method must be GET, POST, PATCH, PUT, or DELETE.' }
    $relativePath = '/v1.0/me'
    if ($p.ContainsKey('RelativePath')) { $relativePath = $p['RelativePath'].Trim() }
    $scopeText = 'User.Read'
    if ($p.ContainsKey('Scopes')) { $scopeText = $p['Scopes'] }
    $scopes = @($scopeText -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $bodyText = $null
    if ($p.ContainsKey('Body')) { $bodyText = $p['Body'] }
 
    $cloud = Get-CloudConfig
    $url = Join-Url -Base $cloud.GraphBase -Path $relativePath
    $token = Get-GraphToken -Scopes $scopes
 
    $body = $null
    if ($bodyText -and $bodyText.Trim() -ne '{}' -and $method -ne 'GET') {
        try { $body = $bodyText | ConvertFrom-Json } catch { $body = $bodyText }
    } elseif ($method -ne 'GET' -and $bodyText) {
        $body = @{}
    }
 
    $result = Invoke-ApiRequest -Method $method -Uri $url -AccessToken $token -Body $body
    Show-ObjectResult -Object $result -Heading 'Custom Graph response'
}
 
function Run-SelectedCommand {
    Clear-ResultState
    $cmd = $lstCommands.SelectedItem
    if ($null -eq $cmd) { throw 'Select a command first.' }
    Add-Log "Running command: $($cmd.Title)"
    Add-Log (Get-CloudConfig).Notes
 
    switch ($cmd.Name) {
        'graph_profile' { Run-GraphProfile }
        'graph_unread_mail' { Run-GraphUnreadMail }
        'copilot_chat' { Run-CopilotChat }
        'pp_list_environments' { Run-PowerPlatformListEnvironments }
        'pp_inactive_environments' { Run-PowerPlatformInactiveEnvironments }
        'pp_list_cloud_flows' { Run-PowerPlatformListCloudFlows }
        'dataverse_my_workflows' { Run-DataverseMyWorkflows }
        'custom_graph' { Run-CustomGraph }
        default { throw "Unknown command: $($cmd.Name)" }
    }
}
 
foreach ($cmd in $script:Commands) { [void]$lstCommands.Items.Add($cmd) }
 
$lstCommands.Add_SelectionChanged({
    $cmd = $lstCommands.SelectedItem
    if ($null -ne $cmd) {
        $txtParams.Text = $cmd.DefaultParameters
        $txtCommandHelp.Text = ($cmd.Kind + [Environment]::NewLine + [Environment]::NewLine + $cmd.Notes)
        Set-Cause ''
    }
})
 
$cmbCloud.Add_SelectionChanged({
    $cloud = Get-CloudConfig
    Add-Log "Cloud selected: $($cloud.Name). Graph base: $($cloud.GraphBase)"
})
 
$btnSignIn.Add_Click({
    Clear-ResultState
    try {
        Add-Log 'Testing sign-in with Microsoft Graph User.Read.'
        [void](Get-GraphToken -Scopes @('User.Read'))
        Add-OutputLine 'signIn = success'
        Add-OutputLine ('cloud = ' + (Get-CloudConfig).Name)
        Add-OutputLine ('graphBase = ' + (Get-CloudConfig).GraphBase)
        Set-Cause ''
    } catch {
        Add-Log $_.Exception.Message 'ERROR'
        $body = Get-ExceptionBody -Exception $_.Exception
        Set-Cause (Get-LikelyCause -Exception $_.Exception -Body $body)
    }
})
 
$btnClearTokens.Add_Click({
    $script:TokenCache.Clear()
    Add-Log 'Token cache cleared.'
})
 
$btnRun.Add_Click({
    try {
        Run-SelectedCommand
    } catch {
        Add-Log $_.Exception.Message 'ERROR'
        if ([string]::IsNullOrWhiteSpace($txtCause.Text)) {
            $body = Get-ExceptionBody -Exception $_.Exception
            Set-Cause (Get-LikelyCause -Exception $_.Exception -Body $body)
        }
    }
})
 
$btnClearOutput.Add_Click({ Clear-ResultState })
 
$btnCopySelected.Add_Click({
    $lines = Get-SelectedOutputLines
    if ($lines.Count -eq 0) { $lines = Get-OutputLines }
    Copy-LinesToClipboard -Lines $lines
})
 
$btnCopyAll.Add_Click({
    Copy-LinesToClipboard -Lines (Get-OutputLines)
})
 
$btnCopyRawJson.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($script:LastRawJson)) {
        [System.Windows.Clipboard]::SetText($script:LastRawJson)
        Add-Log 'Copied raw JSON to clipboard.'
    } else {
        Add-Log 'No raw JSON available to copy.'
    }
})
 
$btnClearLog.Add_Click({ $txtLog.Clear() })
 
if ($lstCommands.Items.Count -gt 0) {
    $lstCommands.SelectedIndex = -1
    $lstCommands.SelectedIndex = 0
}
 
Add-Log 'Tool loaded. Enter Tenant ID/name and Client ID, select a command, then run.'
Add-Log 'For GCC standard, use GCC Standard / Commercial. For GCC High, switch cloud before signing in.'
Set-Cause 'Tip: Create a public-client app registration, enable device-code/public-client flow, and grant delegated permissions needed by the command you select.'
 
[void]$window.ShowDialog()


