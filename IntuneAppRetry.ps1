<#
.SYNOPSIS
    Intune App Retry 1.3 - resets and re-triggers failed Intune Win32 app installations locally.

.DESCRIPTION
    WPF tool (single file, UI + engine) for IT support. Workflow:
      1. Scan      - reads the IME registry (Win32Apps) and lists failed app contexts with a readable status
                     (Waiting / Pending / Installing / Installed / Failed), error code and enforcement state.
      2. Names     - optional: resolves app names via Microsoft Graph (device code sign-in, read-only).
      3. Retry     - backs up the app and GRS registry keys (.reg + SHA256 manifest), deletes them,
                     restarts the Intune Management Extension and triggers an MDM sync.
      4. Monitor   - live monitor after a retry (every 15 s, max. 30 min); reads all registry contexts
                     and new IME log entries (AppWorkload.log) and writes the outcome to the manifest.
      5. History   - lists all retries from the backup manifests (time, operator, app, outcome).
    Data, backups and audit log: C:\ProgramData\IntuneRetryCenter (admins and SYSTEM only).
    Graph is optional; the local retry works offline.
    Optional installation cache capture: watches only <AppId>[_version] folders in
    C:\Windows\IMECache during retry and monitoring, copying available files to
    C:\Temp\IntuneAppRetry-Cache\<AppId>_<timestamp>_<unique>. Copy is best effort.
    Stop monitoring / cache does not cancel an Intune installation or service recovery.

.EXAMPLE
    powershell.exe -STA -File .\IntuneAppRetry.ps1

.NOTES
    Name      : IntuneAppRetry.ps1
    Version   : 1.3
    Developer : Sahiti Ditor
    Requires  : Windows PowerShell 5.1 (64-bit, STA), local administrator rights
    Disclaimer: Deleting IME registry keys is an unsupported troubleshooting workaround.
                A reset does not guarantee installation success.
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitProcess) { throw 'Please use 64-bit Windows PowerShell.' }
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') { throw 'Please start with powershell.exe -STA -File.' }
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Secure diagnostic records and exported registry keys (admins and SYSTEM only).
$script:DataRoot = Join-Path $env:ProgramData 'IntuneRetryCenter'
if (-not (Test-Path -LiteralPath $DataRoot)) { New-Item -ItemType Directory -Path $DataRoot -Force | Out-Null }
$acl = New-Object Security.AccessControl.DirectorySecurity
$acl.SetAccessRuleProtection($true, $false)
foreach ($sid in @('S-1-5-18','S-1-5-32-544')) {
    $identity = New-Object Security.Principal.SecurityIdentifier($sid)
    $rule = New-Object Security.AccessControl.FileSystemAccessRule($identity, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
    $acl.AddAccessRule($rule)
}
Set-Acl -LiteralPath $DataRoot -AclObject $acl
$script:AuditFile = Join-Path $DataRoot ('session-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.log')
$script:Shared = [hashtable]::Synchronized(@{
    Messages = (New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'); GraphConnected = $false
})
$script:Busy = $false
$script:Started = $false
$script:Job = $null
$script:Rows = @()
$script:Monitor = $null

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Intune App Retry 1.3" Width="1300" Height="900" MinWidth="1000" MinHeight="700" WindowStartupLocation="CenterScreen" Background="#F4F6FA" FontFamily="Segoe UI">
 <Grid Margin="24">
  <Grid.RowDefinitions>
   <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
   <RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="138"/>
  </Grid.RowDefinitions>
  <Border Grid.Row="0" Padding="18" Background="#152B4A" CornerRadius="10" Margin="0,0,0,16">
   <Grid>
    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
    <StackPanel>
     <TextBlock Text="Intune App Retry" FontSize="26" FontWeight="SemiBold" Foreground="White"/>
     <TextBlock Text="Review and retry failed Win32 apps - Version 1.3" Margin="0,4,0,0" FontSize="13" Foreground="#C8D6EA"/>
    </StackPanel>
    <Border Grid.Column="1" Background="#274468" CornerRadius="6" Padding="12,8" VerticalAlignment="Center">
     <TextBlock x:Name="GraphState" Foreground="White" Text="Graph: not connected (optional)"/>
    </Border>
   </Grid>
  </Border>
  <WrapPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,12">
   <Button x:Name="Refresh" Content="Scan failed apps locally" Padding="16,10" Margin="0,0,10,0" Background="#FFFFFF"/>
   <Button x:Name="Connect" Padding="16,4" Margin="0,0,10,0" Background="#FFFFFF" ToolTip="Sign in with Microsoft Graph to resolve app names">
    <StackPanel>
     <TextBlock Text="Resolve app names" HorizontalAlignment="Center"/>
     <TextBlock Text="(Sign in with Microsoft Graph)" FontSize="10" Foreground="#526376" HorizontalAlignment="Center"/>
    </StackPanel>
   </Button>
   <Button x:Name="Retry" Content="Retry selected app" Padding="16,10" Margin="0,0,10,0" Background="#1E5BBA" Foreground="White" FontWeight="SemiBold" IsEnabled="False"/>
   <Button x:Name="Logs" Content="Backups / Logs" Padding="12,10" Margin="0,0,10,0"/>
   <Button x:Name="History" Content="Retry history" Padding="12,10" Margin="0,0,10,0"/>
   <Button x:Name="SyncHelp" Content="MDM sync now" Padding="12,10"/>
   <Button x:Name="StopObservation" Content="Stop monitoring / cache" Padding="12,10" Margin="10,0,0,0" IsEnabled="False" ToolTip="Stops observation and cache copying. Does not cancel the installation or registry/service recovery."/>
   <CheckBox x:Name="CaptureCache" Content="Get Installation Cache from Package" Margin="12,8,0,8" VerticalAlignment="Center" ToolTip="During Retry, copy only this app's IMECache folders to C:\Temp\IntuneAppRetry-Cache. Files may be partial while extraction or cleanup is running."/>
  </WrapPanel>
  <Border x:Name="AuthPanel" Grid.Row="2" Margin="0,0,0,12" Padding="12" Visibility="Collapsed" CornerRadius="8" Background="#DFEDFF" BorderBrush="#ACC6ED" BorderThickness="1">
   <StackPanel>
    <TextBlock Text="Microsoft Graph - sign-in code" FontSize="15" FontWeight="SemiBold" Foreground="#17365C"/>
    <TextBlock Text="1. Copy code  2. Open Microsoft login  3. Enter the code in the browser. The code is used for this sign-in only." Margin="0,4,0,8" TextWrapping="Wrap" Foreground="#344C6E"/>
    <StackPanel Orientation="Horizontal">
     <TextBox x:Name="DeviceCode" IsReadOnly="True" Width="195" Height="36" FontSize="21" FontFamily="Consolas" FontWeight="Bold" Padding="6,3" BorderBrush="#7A9CCB" TextAlignment="Center" />
     <Button x:Name="CopyCode" Content="Copy code" Padding="14,7" Margin="10,0,8,0"/>
     <Button x:Name="OpenLogin" Content="Open Microsoft login" Padding="14,7"/>
    </StackPanel>
    <TextBlock Text="Important: only sign in if you started the Graph sign-in yourself. Do not share the code." Margin="0,6,0,0" Foreground="#385579" FontSize="11"/>
   </StackPanel>
  </Border>
  <StackPanel Grid.Row="3" Margin="0,0,0,10">
   <TextBlock x:Name="Status" Text="Reading local Intune state ..." FontWeight="SemiBold" Foreground="#264E83"/>
   <ProgressBar x:Name="BusyBar" IsIndeterminate="True" Height="6" Margin="0,6,0,0" Foreground="#1E5BBA" Background="#DCE5F2" BorderThickness="0" Visibility="Collapsed"/>
   <TextBlock x:Name="BusyText" Margin="0,4,0,0" FontSize="11" Foreground="#1E5BBA" TextTrimming="CharacterEllipsis" Visibility="Collapsed"/>
   <TextBlock Text="Failed Win32 app contexts only. Graph is used for name resolution only and is not required for the local retry." FontSize="12" Foreground="#526376" Margin="0,4,0,0" TextWrapping="Wrap"/>
  </StackPanel>
  <DataGrid x:Name="Apps" ClipboardCopyMode="IncludeHeader" Grid.Row="4" IsReadOnly="True" AutoGenerateColumns="False" SelectionMode="Single" SelectionUnit="FullRow" CanUserAddRows="False" RowHeight="32" HeadersVisibility="Column" GridLinesVisibility="Horizontal" Background="White" BorderBrush="#D7DFEA" AlternatingRowBackground="#F2F6FB" AlternationCount="2">
   <DataGrid.Columns>
    <DataGridTextColumn Header="App" Binding="{Binding AppName}" Width="240"/>
    <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="90"/>
    <DataGridTextColumn Header="Error code" Binding="{Binding ErrorCode}" Width="110"/>
    <DataGridTextColumn Header="Enforcement" Binding="{Binding EnforcementState}" Width="100"/>
    <DataGridTextColumn Header="App-ID" Binding="{Binding AppId}" Width="260"/>
    <DataGridTextColumn Header="Context" Binding="{Binding Context}" Width="270"/>
    <DataGridTextColumn Header="Name via" Binding="{Binding NameSource}" Width="120"/>
    <DataGridTextColumn Header="GRS key" Binding="{Binding GRSRegistry}" Width="410"/>
   </DataGrid.Columns>
  </DataGrid>
  <Border Grid.Row="5" Margin="0,10,0,10" Padding="10" CornerRadius="6" Background="#E7EEF7">
   <StackPanel>
    <TextBlock x:Name="SelectedDetails" Text="Please select an app." FontSize="12" TextWrapping="Wrap" Foreground="#1B3554"/>
    <Border x:Name="MonitorPanel" Visibility="Collapsed" Margin="0,8,0,0" Padding="10,6" CornerRadius="6" Background="#DFEDFF" BorderBrush="#ACC6ED" BorderThickness="1">
     <DockPanel>
      <Button x:Name="StopMonitor" DockPanel.Dock="Right" Content="Stop monitoring" Padding="12,4" Margin="8,0,0,0"/>
      <StackPanel VerticalAlignment="Center">
       <TextBlock x:Name="MonitorTitle" FontWeight="SemiBold" Foreground="#17365C"/>
       <TextBlock x:Name="MonitorSteps" FontSize="11" Foreground="#344C6E" TextWrapping="Wrap"/>
      </StackPanel>
     </DockPanel>
    </Border>
    <TextBlock Margin="0,6,0,0" Text="Retry: back up and reset 2 registry keys; restart IME (Intune check-in) and trigger an MDM sync. No installation guarantee." TextWrapping="Wrap" FontSize="11" Foreground="#526376"/>
   </StackPanel>
  </Border>
  <TextBox x:Name="Activity" Grid.Row="6" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Background="#172B4D" Foreground="#E4ECFA" FontFamily="Consolas" FontSize="12" Padding="10"/>
 </Grid>
</Window>
'@
$reader = New-Object System.Xml.XmlNodeReader($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)
foreach ($name in @('CaptureCache','StopObservation','Status','BusyBar','BusyText','History','MonitorPanel','StopMonitor','MonitorTitle','MonitorSteps','GraphState','Connect','Refresh','Retry','Logs','SyncHelp','Apps','Activity','SelectedDetails','AuthPanel','DeviceCode','CopyCode','OpenLogin')) {
    Set-Variable -Name $name -Value $window.FindName($name) -Scope Script
}
# Shared registry and logging helpers keep scanning and monitoring consistent.
$script:CommonSource = {
$script:RegistryRoot = 'HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps'
$script:RegistryNative = 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps'
$script:LogsRoot = Join-Path $env:ProgramData 'Microsoft\IntuneManagementExtension\Logs'
$script:GuidPattern = '^[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$'
$script:AppPattern = '^(?<id>[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12})(?:_(?<version>\d+))?$'

function Write-Activity {
    param([Parameter(Mandatory)][string]$Message)
    $line = ('{0:HH:mm:ss}  {1}' -f (Get-Date), $Message)
    $Shared.Messages.Enqueue($line)
    # Audit output is protected by the ACL configured in the GUI startup.
    try { Add-Content -LiteralPath $AuditFile -Value $line -Encoding UTF8 -ErrorAction Stop } catch { }
}

function Get-RegistryValue {
    param([string]$Path, [string]$Value)
    try { return (Get-Item -LiteralPath $Path -ErrorAction Stop).GetValue($Value, $null) }
    catch { return $null }
}

function Read-RegistryJson {
    param([string]$Key, [string]$Value)
    $raw = Get-RegistryValue -Path $Key -Value $Value
    if ($raw -is [string] -and $raw.TrimStart().StartsWith('{')) {
        try { return ($raw | ConvertFrom-Json -ErrorAction Stop) } catch { return $null }
    }
    return $null
}

function Get-IntegerProperty {
    param($Object, [string]$Member)
    if ($null -ne $Object -and $null -ne $Object.PSObject.Properties[$Member]) {
        try { return [int]$Object.$Member } catch { }
    }
    return $null
}

function Get-AppRegistryState {
    # Use one classification rule for the scan and live monitor; local state can lag behind Intune.
    param([string]$Path)
    $cObj = Read-RegistryJson -Key (Join-Path $Path 'ComplianceStateMessage') -Value 'ComplianceStateMessage'
    $eObj = Read-RegistryJson -Key (Join-Path $Path 'EnforcementStateMessage') -Value 'EnforcementStateMessage'
    $resultObj = Read-RegistryJson -Key $Path -Value 'Result'
    $c = Get-IntegerProperty $cObj 'ComplianceState'
    $e = Get-IntegerProperty $eObj 'EnforcementState'
    if ($null -eq $c) { $c = Get-IntegerProperty $resultObj 'ComplianceState' }
    if ($null -eq $e) { $e = Get-IntegerProperty $resultObj 'EnforcementState' }
    $codes = @()
    foreach ($obj in @($eObj, $cObj, $resultObj)) {
        if ($null -ne $obj -and $null -ne $obj.PSObject.Properties['ErrorCode']) {
            $code = [string]$obj.ErrorCode
            if (-not [string]::IsNullOrWhiteSpace($code)) { $codes += $code.Trim() }
        }
    }
    $realCodes = @($codes | Where-Object { $_ -notmatch '^(?:0|0x0+|3010|0x0*BC2)$' })
    $errorCode = if ($realCodes.Count -gt 0) { $realCodes[0] } elseif ($codes.Count -gt 0) { $codes[0] } else { '-' }
    $running = $null -ne $e -and $e -ge 2000 -and $e -lt 3000
    $success = $c -eq 1 -or ($null -ne $e -and $e -ge 1000 -and $e -lt 2000)
    $failed = -not $running -and -not $success -and (
        $c -eq 4 -or ($null -ne $e -and $e -ge 5000 -and $e -lt 6000) -or $realCodes.Count -gt 0
    )
    $status = if ($running) { 'Installing' } elseif ($success) { 'Installed' } elseif ($failed) { 'Failed' } elseif ($null -ne $e -or $null -ne $c) { 'Pending' } else { 'Unknown' }
    [pscustomobject]@{ Compliance=$c; Enforcement=$e; ErrorCode=$errorCode; Running=[bool]$running; Failed=[bool]$failed; Status=$status; IsFinal=[bool]($success -or $failed); Codes=$codes }
}

function Test-AppContext {
    param([string]$Context)
    return ($Context -match $script:GuidPattern -or $Context -match '^S-1-[0-9-]+$')
}

function Get-SpecificAppKeys {
    param([string]$Context, [string]$AppId)
    if (-not (Test-AppContext $Context) -or $AppId -notmatch $script:GuidPattern) { throw 'Invalid app ID / context identifier.' }
    $contextPath = Join-Path $script:RegistryRoot $Context
    if (-not (Test-Path -LiteralPath $contextPath)) { return @() }
    return @(Get-ChildItem -LiteralPath $contextPath -ErrorAction Stop | Where-Object {
        $_.PSChildName -match ('^' + [regex]::Escape($AppId) + '(?:_\d+)?$')
    })
}
}.ToString()
. ([scriptblock]::Create($script:CommonSource))

# Worker engine

$script:EngineSource = {
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:GraphNames = @{}
$script:GraphErrors = @{}
if (-not $Shared.ContainsKey('GraphConnected')) { $Shared.GraphConnected = $false }

function Get-MatchingGrs {
    # A GRS container can belong to several apps. Never reset a container marked as shared.
    param([string]$Context, [string]$AppId)
    if (-not (Test-AppContext $Context) -or $AppId -notmatch $script:GuidPattern) { throw 'Invalid GRS search parameters.' }
    $root = Join-Path (Join-Path $script:RegistryRoot $Context) 'GRS'
    if (-not (Test-Path -LiteralPath $root)) { return @() }
    $found = @()
    foreach ($key in @(Get-ChildItem -LiteralPath $root -ErrorAction Stop)) {
        $valueNames = @()
        try { $valueNames = @((Get-Item -LiteralPath $key.PSPath -ErrorAction Stop).GetValueNames()) }
        catch { continue }
        $matched = @($valueNames | Where-Object { $_ -match ('^' + [regex]::Escape($AppId) + '(?:_\d+)?$') })
        if ($matched.Count -eq 0) { continue }
        $others = @($valueNames | Where-Object {
            if ($_ -match $script:AppPattern) { return ($Matches['id'] -ine $AppId) }
            return $false
        })
        $found += [pscustomobject]@{
            Path = [string]$key.Name
            Shared = [bool]($others.Count -gt 0)
            Evidence = ($matched -join ', ')
        }
    }
    return @($found)
}

function Get-AllAppEntries {
    if (-not (Test-Path -LiteralPath $script:RegistryRoot)) { throw 'Intune Win32Apps registry key does not exist.' }
    $rows = @()
    foreach ($ctx in @(Get-ChildItem -LiteralPath $script:RegistryRoot -ErrorAction Stop)) {
        $context = [string]$ctx.PSChildName
        if (-not (Test-AppContext $context)) { continue }
        $keys = @(Get-ChildItem -LiteralPath $ctx.PSPath -ErrorAction Stop | Where-Object { $_.PSChildName -match $script:AppPattern })
        foreach ($group in @($keys | Group-Object { $_.PSChildName -replace '_\d+$', '' })) {
            $sorted = @($group.Group | Sort-Object @{Expression={
                if ($_.PSChildName -match '_(\d+)$') { [long]$Matches[1] } else { [long]0 }
            }} -Descending)
            if ($sorted.Count -eq 0) { continue }
            $key = $sorted[0]
            $appId = ($key.PSChildName -replace '_\d+$', '').ToLowerInvariant()
            $state = Get-AppRegistryState $key.PSPath
            $grs = @(Get-MatchingGrs $context $appId)
            $regName = Get-RegistryValue $key.PSPath 'DisplayName'
            if ([string]::IsNullOrWhiteSpace([string]$regName)) { $regName = Get-RegistryValue $key.PSPath 'AppName' }
            $name = if ([string]::IsNullOrWhiteSpace([string]$regName)) { '(name unknown)' } else { [string]$regName }
            $source = if ($name -eq '(name unknown)') { '-' } else { 'Registry' }
            $grsText = if ($grs.Count -eq 0) { '(no unique match)' }
                elseif ($grs.Count -gt 1) { '(multiple GRS hits: blocked)' }
                elseif ($grs[0].Shared) { '(GRS shared with another app: blocked)' }
                else { $grs[0].Path }
            $rows += [pscustomobject]@{
                AppName=$name; NameSource=$source; AppId=$appId; Context=$context;
                ComplianceState=$state.Compliance; EnforcementState=$state.Enforcement;
                ErrorCode=$state.ErrorCode; Status=$state.Status; Failed=$state.Failed; Running=$state.Running;
                AppRegistry=([string]$key.Name); GRSRegistry=$grsText
            }
        }
    }
    return @($rows)
}

function Get-LocalAppNames {
    # Only explicit policy ID/name pairs are trusted; nearby log lines may describe other apps.
    param([string[]]$AppIds)
    $names = @{}
    if ($AppIds.Count -eq 0 -or -not (Test-Path -LiteralPath $script:LogsRoot)) { return $names }
    $wanted = @{}
    foreach ($id in $AppIds) { $wanted[$id.ToLowerInvariant()] = $true }
    $files = @(Get-ChildItem -LiteralPath $script:LogsRoot -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^(?:_?IntuneManagementExtension|AppWorkload|AppActionProcessor).*\.log(?:\.\d+)?$' } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 18)
    foreach ($file in $files) {
        $stream = $null; $reader = $null
        try {
            $stream = [IO.File]::Open($file.FullName, 'Open', 'Read', 'ReadWrite')
            $reader = New-Object IO.StreamReader($stream, [Text.Encoding]::UTF8, $true)
            while ($null -ne ($line = $reader.ReadLine())) {
                $marker = $line.IndexOf('Get policies =', [StringComparison]::OrdinalIgnoreCase)
                if ($marker -lt 0) { continue }
                $start = $line.IndexOf('[', $marker + 'Get policies ='.Length)
                if ($start -lt 0) { continue }
                $end = $line.IndexOf(']LOG]!', $start, [StringComparison]::Ordinal)
                if ($end -lt 0) { $end = $line.LastIndexOf(']') + 1 }
                if ($end -le $start) { continue }
                try {
                    $policies = @($line.Substring($start, $end-$start).Trim() | ConvertFrom-Json -ErrorAction Stop)
                    foreach ($policy in $policies) {
                        $id = [string]$policy.Id
                        $name = [string]$policy.Name
                        if ($id -and $name -and $wanted.ContainsKey($id.ToLowerInvariant()) -and -not $names.ContainsKey($id.ToLowerInvariant())) {
                            $names[$id.ToLowerInvariant()] = $name.Trim()
                        }
                    }
                } catch { }
            }
        } catch { }
        finally {
            if ($null -ne $reader) { $reader.Dispose() }
            elseif ($null -ne $stream) { $stream.Dispose() }
        }
    }
    return $names
}

function Resolve-AppNames {
    # Graph names take precedence; local names remain usable when Graph is unavailable.
    param([object[]]$Rows)
    $ids = @($Rows | Select-Object -ExpandProperty AppId -Unique)
    if ($ids.Count -eq 0) { return }
    $local = Get-LocalAppNames $ids
    foreach ($row in $Rows) {
        if ($row.NameSource -eq '-' -and $local.ContainsKey($row.AppId)) {
            $row.AppName = $local[$row.AppId]
            $row.NameSource = 'IME log'
        }
    }
    if (-not [bool]$Shared.GraphConnected) { return }
    foreach ($id in $ids) {
        # Retry unresolved names on every scan; prior failures must not block a new lookup.
        if (-not $script:GraphNames.ContainsKey($id)) {
            $script:GraphErrors.Remove($id)
            try {
                $token = [string]$Shared.GraphAccessToken
                if ([string]::IsNullOrWhiteSpace($token)) { throw 'No Graph access token available.' }
                $headers = @{ Authorization = ('Bearer ' + $token) }
                $response = $null
                $lastError = ''
                foreach ($version in @('v1.0', 'beta')) {
                    try {
                        $uri = 'https://graph.microsoft.com/' + $version + '/deviceAppManagement/mobileApps/' + $id + '?$select=id,displayName'
                        $response = Invoke-RestMethod -Method Get -Uri $uri -Headers $headers -ErrorAction Stop
                        break
                    } catch {
                        $lastError = [string]$_.Exception.Message
                        try {
                            $webResponse = $_.Exception.Response
                            if ($null -ne $webResponse) {
                                $wwwAuth = [string]$webResponse.Headers['WWW-Authenticate']
                                if (-not [string]::IsNullOrWhiteSpace($wwwAuth)) {
                                    $wwwAuth = $wwwAuth -replace 'authorization_uri="[^"]*",?\s*', ''
                                    $lastError += ' [' + $version + ' WWW-Authenticate: ' + $wwwAuth + ']'
                                }
                            }
                        } catch { }
                        if ($null -ne $_.ErrorDetails -and -not [string]::IsNullOrWhiteSpace([string]$_.ErrorDetails.Message)) {
                            $detail = [string]$_.ErrorDetails.Message
                            if ($detail.Length -gt 300) { $detail = $detail.Substring(0, 300) }
                            $lastError += ' ' + $detail
                        }
                    }
                }
                if ($null -eq $response) { throw $lastError }
                if ($null -ne $response.PSObject.Properties['displayName'] -and -not [string]::IsNullOrWhiteSpace([string]$response.displayName)) {
                    $script:GraphNames[$id] = [string]$response.displayName
                } else {
                    throw 'Response contained no displayName.'
                }
            } catch {
                $script:GraphErrors[$id] = [string]$_.Exception.Message
                Write-Activity "Graph: name for $id could not be retrieved ($($_.Exception.Message))."
            }
        }
    }
    Write-Activity ("Graph: resolved $($script:GraphNames.Count) of $($ids.Count) app names.")
    foreach ($row in $Rows) {
        if ($script:GraphNames.ContainsKey($row.AppId)) {
            $row.AppName = $script:GraphNames[$row.AppId]
            $row.NameSource = 'Microsoft Graph'
        }
    }
}

function Get-FailedApps {
    Write-Activity 'Scanning Intune Win32 registry ...'
    $all = @(Get-AllAppEntries)
    $failed = @($all | Where-Object { $_.Failed } | Sort-Object AppName, AppId, Context)
    Resolve-AppNames $failed
    $unique = @($failed | Select-Object -ExpandProperty AppId -Unique).Count
    Write-Activity "Found: $($failed.Count) failed contexts ($unique distinct app IDs)."
    return @($failed)
}

function Get-AuthErrorDetails {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)
    $message = [string]$ErrorRecord.Exception.Message
    # PowerShell 5.1 consumes the response stream; the body is kept in ErrorDetails.
    if ($null -ne $ErrorRecord.ErrorDetails -and -not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.ErrorDetails.Message)) {
        try {
            $body = [string]$ErrorRecord.ErrorDetails.Message | ConvertFrom-Json -ErrorAction Stop
            if ($null -ne $body.PSObject.Properties['error']) {
                return [pscustomobject]@{ Code = [string]$body.error; Message = [string]$body.error_description }
            }
        } catch { }
    }
    $resp = $null
    if ($ErrorRecord.Exception -is [System.Net.WebException]) { $resp = $ErrorRecord.Exception.Response }
    if ($null -ne $resp) {
        try {
            $stream = $resp.GetResponseStream()
            if ($null -ne $stream) {
                $reader = New-Object -TypeName System.IO.StreamReader -ArgumentList $stream
                try {
                    $body = $reader.ReadToEnd() | ConvertFrom-Json -ErrorAction Stop
                    return [pscustomobject]@{
                        Code = [string]$body.error
                        Message = [string]$body.error_description
                    }
                } finally { $reader.Dispose() }
            }
        } catch { }
    }
    return [pscustomobject]@{ Code = ''; Message = $message }
}

function Connect-Intune {
    $Shared.GraphConnected = $false
    $Shared.GraphAccessToken = $null
    $Shared.DeviceCodeDisplay = $null
    Write-Activity 'Starting Microsoft Graph sign-in (device code shown in the GUI).'

    $clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
    $tenant = 'organizations'
    $authBase = 'https://login.microsoftonline.com/' + $tenant + '/oauth2/v2.0/'
    $scope = 'https://graph.microsoft.com/DeviceManagementApps.Read.All'
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        $device = Invoke-RestMethod -Method Post -Uri ($authBase + 'devicecode') -Body @{
            client_id = $clientId
            scope = $scope
        } -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop

        if ([string]::IsNullOrWhiteSpace([string]$device.user_code) -or
            [string]::IsNullOrWhiteSpace([string]$device.device_code) -or
            [string]::IsNullOrWhiteSpace([string]$device.verification_uri)) {
            throw 'The identity provider did not return a complete device code.'
        }
        # Only the short USER code is presented. Secret device_code and tokens
        # never enter the GUI message log, disk, or clipboard.
        $Shared.DeviceCodeDisplay = [pscustomobject]@{
            Code = [string]$device.user_code
            Url  = [string]$device.verification_uri
        }
        Write-Activity 'Graph: one-time code is shown in the GUI. Enter only that code in the browser.'
        $interval = [math]::Max(5, [int]$device.interval)
        $expiresAt = (Get-Date).AddSeconds([math]::Max(1, [int]$device.expires_in))
        $tokenUrl = $authBase + 'token'
        while ((Get-Date) -lt $expiresAt) {
            Start-Sleep -Seconds $interval
            try {
                $auth = Invoke-RestMethod -Method Post -Uri $tokenUrl -Body @{
                    grant_type = 'urn:ietf:params:oauth:grant-type:device_code'
                    client_id  = $clientId
                    device_code = [string]$device.device_code
                } -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
                if ([string]::IsNullOrWhiteSpace([string]$auth.access_token)) {
                    throw 'Graph sign-in returned no access token.'
                }
                $grantedScopes = @(([string]$auth.scope) -split ' ' | ForEach-Object { ($_ -replace '^https://graph\.microsoft\.com/', '') })
                if (-not ($grantedScopes -contains 'DeviceManagementApps.Read.All')) {
                    throw 'Token is missing DeviceManagementApps.Read.All.'
                }
                $Shared.GraphAccessToken = [string]$auth.access_token
                try {
                    # Diagnostics only: log non-secret claims, never the token itself.
                    $payload = ([string]$auth.access_token).Split('.')[1].Replace('-', '+').Replace('_', '/')
                    switch ($payload.Length % 4) { 2 { $payload += '==' } 3 { $payload += '=' } }
                    $claims = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload)) | ConvertFrom-Json
                    Write-Activity ("Graph token: aud=" + [string]$claims.aud + " | scp=" + [string]$claims.scp + " | upn=" + [string]$claims.upn)
                } catch { Write-Activity 'Graph token: claims could not be read.' }
                $Shared.GraphConnected = $true
                $script:GraphNames.Clear()
                $script:GraphErrors.Clear()
                Write-Activity 'Graph sign-in successful. Loading names of failed Intune apps.'
                return
            } catch {
                $errorInfo = Get-AuthErrorDetails $_
                switch ([string]$errorInfo.Code) {
                    'authorization_pending' { continue }
                    'slow_down' { $interval += 5; continue }
                    'authorization_declined' { throw 'Sign-in was declined.' }
                    'expired_token' { throw 'Sign-in code has expired.' }
                    default {
                        $details = [string]$errorInfo.Message
                        if ($details.Length -gt 700) { $details = $details.Substring(0,700) }
                        throw ('Graph sign-in denied: ' + $details)
                    }
                }
            }
        }
        throw 'Sign-in code has expired. Click Graph sign-in again if needed.'
    } catch {
        $Shared.GraphConnected = $false
        $Shared.GraphAccessToken = $null
        $details = [string]$_.Exception.Message
        # Do not include authentication materials in diagnostic logs.
        Write-Activity ('Graph sign-in not possible: ' + $details + ' Local diagnostics remain available.')
    } finally {
        $Shared.DeviceCodeDisplay = $null
    }
}

function ConvertTo-RegistryProviderPath {
    param([string]$Native)
    if (-not $Native.StartsWith('HKEY_LOCAL_MACHINE\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Non-HKLM registry path blocked.' }
    return ('HKLM:\' + $Native.Substring('HKEY_LOCAL_MACHINE\'.Length))
}

function Test-RepairPath {
    param([string]$Native, [string]$Context, [string]$AppId, [bool]$GRS)
    $prefix = $script:RegistryNative + '\' + $Context + '\'
    if (-not $Native.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { return $false }
    $child = $Native.Substring($prefix.Length)
    if ($GRS) { return ($child -match '^GRS\\[^\\]+$') }
    return ($child -match ('^' + [regex]::Escape($AppId) + '(?:_\d+)?$'))
}

function Get-RepairPlan {
    # Re-read the registry instead of trusting paths captured by an earlier GUI scan.
    param($Record)
    $appId = [string]$Record.AppId
    $context = [string]$Record.Context
    if ($appId -notmatch $script:GuidPattern -or -not (Test-AppContext $context)) { throw 'Invalid app ID or context.' }
    $apps = @(Get-SpecificAppKeys $context $appId)
    if ($apps.Count -ne 1) { throw "Expected exactly one app key; found: $($apps.Count). No changes made." }
    $state = Get-AppRegistryState $apps[0].PSPath
    if (-not $state.Failed -or $state.Running) { throw 'App is no longer clearly failed or is currently running. Please rescan.' }
    $grs = @(Get-MatchingGrs $context $appId)
    if ($grs.Count -ne 1) { throw "GRS not uniquely matched: $($grs.Count) hits. No changes made." }
    if ($grs[0].Shared) { throw 'GRS key is shared with another app. No changes made.' }
    if (-not (Test-RepairPath $apps[0].Name $context $appId $false)) { throw 'App path outside the allowed scope.' }
    if (-not (Test-RepairPath $grs[0].Path $context $appId $true)) { throw 'GRS path outside the allowed scope.' }
    $paths = @([string]$apps[0].Name, [string]$grs[0].Path)
    [pscustomobject]@{ AppId=$appId; Context=$context; Paths=$paths; AppPath=$paths[0]; GrsPath=$paths[1] }
}

function Save-RegistryBackups {
    # Every export must succeed before the caller is allowed to remove either key.
    param($Plan, $Record)
    $label = ('{0}_{1}_{2}_{3}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), $Plan.AppId,
        $Plan.Context, ([guid]::NewGuid().ToString('N').Substring(0,8)))
    $dir = Join-Path $DataRoot (Join-Path 'Backups' $label)
    $null = New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop
    $exports = @()
    $regExe = Join-Path $env:SystemRoot 'System32\reg.exe'
    for ($i = 0; $i -lt $Plan.Paths.Count; $i++) {
        $regPath = [string]$Plan.Paths[$i]
        $destination = Join-Path $dir ('{0:D2}.reg' -f ($i+1))
        & $regExe export $regPath $destination '/y' 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $destination) -or (Get-Item -LiteralPath $destination).Length -eq 0) {
            throw "Registry export failed for $regPath. No keys deleted. Folder: $dir"
        }
        $exports += [pscustomobject]@{ RegistryKey=$regPath; FileName=(Split-Path $destination -Leaf); SHA256=(Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash }
    }
    [pscustomobject]@{
        Time=(Get-Date).ToString('o'); Computer=$env:COMPUTERNAME; Operator=$env:USERNAME;
        AppName=[string]$Record.AppName; AppId=$Plan.AppId; Context=$Plan.Context; Files=$exports
    } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $dir 'manifest.json') -Encoding UTF8
    return $dir
}

function Invoke-AppRetry {
    # Keep service recovery in finally so a failed reset does not leave IME stopped.
    param($Record)
    # Only called following an explicit confirmation from the GUI.
    if ($null -eq $Record) { throw 'No app selected.' }
    $service = Get-Service -Name 'IntuneManagementExtension' -ErrorAction Stop
    $plan = Get-RepairPlan $Record
    Write-Activity "Retry selected: $($Record.AppName) / $($plan.AppId) / $($plan.Context)."
    $others = @(Get-AllAppEntries | Where-Object {
        $_.Running -and -not ($_.AppId -eq $plan.AppId -and $_.Context -eq $plan.Context)
    })
    if ($others.Count -gt 0) { throw "Other app installations appear to be active ($($others.Count)). Please try again later." }
    $shouldStart = $false
    $backupDir = '(none yet)'
    $deleted = 0
    try {
        $service.Refresh()
        if ($service.Status -eq 'Running') {
            $shouldStart = $true
            Write-Activity 'Stopping Intune Management Extension ...'
            Stop-Service -Name 'IntuneManagementExtension' -ErrorAction Stop
            $service.WaitForStatus([ServiceProcess.ServiceControllerStatus]::Stopped, [TimeSpan]::FromSeconds(45))
        } elseif ($service.Status -eq 'Stopped') {
            $shouldStart = $true
        } else { throw "IME service state is not safe: $($service.Status)." }
        # Re-read on stopped agent to guard against state changes since the GUI scan.
        $fresh = Get-RepairPlan $Record
        if ((@($fresh.Paths | Sort-Object) -join '|') -ne (@($plan.Paths | Sort-Object) -join '|')) {
            throw 'Registry keys changed in the meantime. No changes made.'
        }
        Write-Activity 'Exporting both registry keys (.reg + SHA256 manifest) ...'
        $backupDir = Save-RegistryBackups -Plan $fresh -Record $Record
        Write-Activity "Backup created: $backupDir"
        $Shared.LastBackupDir = $backupDir
        foreach ($key in $fresh.Paths) {
            if (-not (Test-RepairPath $key $fresh.Context $fresh.AppId ($key -eq $fresh.GrsPath))) {
                throw "Unsafe path: $key"
            }
            $psPath = ConvertTo-RegistryProviderPath $key
            if (-not (Test-Path -LiteralPath $psPath)) { throw "Registry key disappeared unexpectedly: $key" }
            Remove-Item -LiteralPath $psPath -Force -Recurse -ErrorAction Stop
            $deleted++
            Write-Activity "Registry key deleted ($deleted/2): $key"
        }
    } finally {
        if ($shouldStart) {
            Write-Activity 'Starting Intune Management Extension / IME check-in ...'
            try {
                $s = Get-Service -Name 'IntuneManagementExtension' -ErrorAction Stop
                $s.Refresh()
                if ($s.Status -ne 'Running') {
                    if ($s.Status -ne 'Stopped') { $s.WaitForStatus([ServiceProcess.ServiceControllerStatus]::Stopped, [TimeSpan]::FromSeconds(40)) }
                    Start-Service -Name 'IntuneManagementExtension' -ErrorAction Stop
                    (Get-Service -Name 'IntuneManagementExtension').WaitForStatus([ServiceProcess.ServiceControllerStatus]::Running, [TimeSpan]::FromSeconds(40))
                }
                Write-Activity 'IME service running; new IME check-in requested.'
            } catch {
                Write-Activity "CRITICAL: IME could not be started: $($_.Exception.Message)"
                throw "IME restart failed! Backups: $backupDir. Please check the service manually. $($_.Exception.Message)"
            }
        }
    }
    if ($deleted -ne 2) { throw "Registry reset incomplete ($deleted/2); IME was restarted. Backups: $backupDir" }
    Write-Activity "Done: 2 keys reset. Backup: $backupDir. Installation has not yet been confirmed."
    Invoke-MdmSync
}

function Invoke-MdmSync {
    # Task submission is not proof that synchronization completed.
    Write-Activity 'Triggering MDM sync ...'
    try {
        $tasks = @(Get-ScheduledTask -TaskPath '\Microsoft\Windows\EnterpriseMgmt\*' -ErrorAction Stop |
            Where-Object { $_.TaskName -eq 'Schedule #3 created by enrollment client' })
        if ($tasks.Count -eq 0) {
            Write-Activity 'MDM sync: no enrollment sync task found (device not MDM-enrolled?). Use Company Portal > Settings > Sync.'
            return
        }
        foreach ($task in $tasks) {
            Start-ScheduledTask -TaskPath $task.TaskPath -TaskName $task.TaskName -ErrorAction Stop
            Write-Activity ("MDM sync triggered: " + $task.TaskPath)
        }
    } catch {
        Write-Activity ('MDM sync could not be triggered: ' + $_.Exception.Message + ' Use Company Portal > Settings > Sync.')
    }
}
}.ToString()
$script:Runspace = [RunspaceFactory]::CreateRunspace()
$Runspace.ApartmentState = 'STA'
$Runspace.ThreadOptions = 'ReuseThread'
$Runspace.Open()
$Runspace.SessionStateProxy.SetVariable('Shared', $Shared)
$Runspace.SessionStateProxy.SetVariable('DataRoot', $DataRoot)
$Runspace.SessionStateProxy.SetVariable('AuditFile', $AuditFile)
$Runspace.SessionStateProxy.SetVariable('CommonSource', $script:CommonSource)
$Runspace.SessionStateProxy.SetVariable('EngineSource', $script:EngineSource)
$init = [PowerShell]::Create()
$init.Runspace = $Runspace
$null = $init.AddScript('. ([scriptblock]::Create($CommonSource)); . ([scriptblock]::Create($EngineSource))', $false)
try {
    $null = $init.Invoke()
    if ($init.HadErrors) { throw ($init.Streams.Error | Out-String) }
} finally { $init.Dispose() }

# Cache capture runs independently from registry repair and the UI dispatcher.
$script:CacheJob = $null
$script:ObservationStopped = $false
$script:RefreshPending = $false

function Start-InstallationCacheCapture {
    param([string]$AppId)
    if ($null -ne $script:CacheJob) { throw 'A cache capture is still finishing. Please wait.' }
    if ($AppId -notmatch $script:GuidPattern) { throw 'Invalid app ID for cache capture.' }
    $base = 'C:\Temp\IntuneAppRetry-Cache'
    foreach ($parent in @('C:\Temp', $base)) {
        if (-not (Test-Path -LiteralPath $parent)) { New-Item -ItemType Directory -Path $parent -ErrorAction Stop | Out-Null }
        if ((Get-Item -LiteralPath $parent -ErrorAction Stop).Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw ('Cache destination must not be a junction or symbolic link: ' + $parent)
        }
    }
    $destination = Join-Path $base ($AppId + '_' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '_' + [guid]::NewGuid().ToString('N').Substring(0,8))
    New-Item -ItemType Directory -Path $destination -ErrorAction Stop | Out-Null
    $captureAcl = New-Object Security.AccessControl.DirectorySecurity
    $captureAcl.SetAccessRuleProtection($true, $false)
    foreach ($sid in @('S-1-5-18','S-1-5-32-544')) {
        $identity = New-Object Security.Principal.SecurityIdentifier($sid)
        $rule = New-Object Security.AccessControl.FileSystemAccessRule($identity, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        $captureAcl.AddAccessRule($rule)
    }
    Set-Acl -LiteralPath $destination -AclObject $captureAcl -ErrorAction Stop
    $captureState = [hashtable]::Synchronized(@{ Stop=$false; Finish=$false; Path=$destination; Result='Watching'; Error=''; AppId=$AppId })
    $captureRunspace = [RunspaceFactory]::CreateRunspace()
    $captureRunspace.Open()
    $captureRunspace.SessionStateProxy.SetVariable('Capture', $captureState)
    $captureRunspace.SessionStateProxy.SetVariable('Shared', $Shared)
    $capturePS = [PowerShell]::Create()
    $capturePS.Runspace = $captureRunspace
    $null = $capturePS.AddScript({
        Set-StrictMode -Version 2.0
        $ErrorActionPreference = 'Stop'
        $source = Join-Path $env:windir 'IMECache'
        $pattern = '^' + [regex]::Escape($Capture.AppId) + '(?:_\d+)?$'
        $deadline = (Get-Date).AddMinutes(30)
        $seen = @{}
        $copyProcess = $null
        $started = (Get-Date).ToString('o')
        $interrupted = $false
        try {
            while (-not $Capture.Stop -and (Get-Date) -lt $deadline) {
                # A finish request always gets one final pass; Stop cancels promptly.
                $lastPass = [bool]$Capture.Finish
                if (Test-Path -LiteralPath $source) {
                    if ((Get-Item -LiteralPath $source).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'IMECache root is a reparse point.' }
                    $folders = @(Get-ChildItem -LiteralPath $source -Directory -ErrorAction Stop | Where-Object {
                        $_.Name -match $pattern -and -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint)
                    })
                    foreach ($folder in $folders) {
                        if ($Capture.Stop) { break }
                        $target = Join-Path $Capture.Path $folder.Name
                        $copyLog = Join-Path $Capture.Path 'copy.log'
                        # Additive copy only: never mirror, purge, move or delete source files.
                        $info = New-Object Diagnostics.ProcessStartInfo
                        $info.FileName = Join-Path $env:windir 'System32\robocopy.exe'
                        $info.Arguments = '"' + $folder.FullName + '" "' + $target + '" /E /COPY:DAT /DCOPY:DAT /XJ /XJD /XJF /R:0 /W:0 /NP /NFL /NDL /NJH /NJS /LOG+:"' + $copyLog + '"'
                        $info.UseShellExecute = $false
                        $info.CreateNoWindow = $true
                        $copyProcess = New-Object Diagnostics.Process
                        $copyProcess.StartInfo = $info
                        [void]$copyProcess.Start()
                        $copyDeadline = (Get-Date).AddMinutes(2)
                        $cutShort = $false
                        while (-not $copyProcess.WaitForExit(200)) {
                            if ($Capture.Stop -or (Get-Date) -ge $copyDeadline) {
                                $cutShort = $true
                                $interrupted = $true
                                $copyProcess.Kill()
                                $copyProcess.WaitForExit()
                                break
                            }
                        }
                        $exitCode = $copyProcess.ExitCode
                        $copyProcess.Dispose()
                        $copyProcess = $null
                        $previous = if ($seen.ContainsKey($folder.Name)) { $seen[$folder.Name] } else { $null }
                        $seen[$folder.Name] = [pscustomobject]@{ Folder=$folder.Name; Source=$folder.FullName; Destination=$target; LastExitCode=$exitCode; Interrupted=$cutShort; LastAttempt=(Get-Date).ToString('o') }
                        if ($null -eq $previous -or $previous.LastExitCode -ne $exitCode -or $previous.Interrupted -ne $cutShort) {
                            $Shared.Messages.Enqueue(('{0:HH:mm:ss}  Cache: {1}, robocopy exit {2}{3}' -f (Get-Date), $folder.Name, $exitCode, $(if ($cutShort) { ' (interrupted)' } else { '' })))
                        }
                    }
                }
                if ($lastPass) { break }
                for ($pause=0; $pause -lt 5 -and -not $Capture.Stop -and -not $Capture.Finish; $pause++) { Start-Sleep -Milliseconds 200 }
            }
            $files = @(Get-ChildItem -LiteralPath $Capture.Path -Directory | ForEach-Object {
                Get-ChildItem -LiteralPath $_.FullName -File -Recurse -ErrorAction SilentlyContinue
            })
            $failures = @($seen.Values | Where-Object { $_.LastExitCode -ge 8 -or $_.Interrupted })
            $Capture.Result = if ($Capture.Stop) { 'Stopped; retained files may be partial' }
                elseif ((Get-Date) -ge $deadline) { 'Timed out; retained files may be partial' }
                elseif ($files.Count -eq 0) { 'No package files captured' }
                elseif ($failures.Count -gt 0 -or $interrupted) { 'Captured with copy errors or interruptions; inspect copy.log' }
                else { 'Captured available files; live copy, completeness not guaranteed' }
        } catch {
            $Capture.Result = 'Capture failed; retained files may be partial'
            $Capture.Error = $_.Exception.Message
        } finally {
            if ($null -ne $copyProcess) {
                try { if (-not $copyProcess.HasExited) { $copyProcess.Kill(); $copyProcess.WaitForExit() } } catch { }
                $copyProcess.Dispose()
            }
            try {
                [pscustomobject]@{
                    AppId=$Capture.AppId; Started=$started; Finished=(Get-Date).ToString('o')
                    Source=$source; Destination=$Capture.Path; Result=$Capture.Result; Error=$Capture.Error
                    Folders=@($seen.Values)
                    Note='Best-effort live copy for this app ID only. Files can change or disappear during extraction/cleanup. Never execute automatically.'
                } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $Capture.Path 'capture.json') -Encoding UTF8
            } catch { $Capture.Error += ' Could not write capture.json: ' + $_.Exception.Message }
        }
    })
    try {
        $script:CacheJob = @{ PS=$capturePS; Runspace=$captureRunspace; State=$captureState; Handle=$capturePS.BeginInvoke() }
    } catch {
        $capturePS.Dispose(); $captureRunspace.Dispose(); throw
    }
    Write-Activity ('Cache capture started for ' + $AppId + ' -> ' + $destination)
}

function Complete-CacheCapture {
    $captureJob = $script:CacheJob
    if ($null -eq $captureJob -or -not $captureJob.Handle.IsCompleted) { return }
    try {
        $null = $captureJob.PS.EndInvoke($captureJob.Handle)
        if ($captureJob.PS.HadErrors) { Write-Activity ('Cache worker error: ' + ($captureJob.PS.Streams.Error | Out-String)) }
        Write-Activity ('Cache: ' + $captureJob.State.Result + '. Destination: ' + $captureJob.State.Path)
        if ($captureJob.State.Error) { Write-Activity ('Cache error: ' + $captureJob.State.Error) }
    } catch { Write-Activity ('Cache worker error: ' + $_.Exception.Message) }
    finally {
        $captureJob.PS.Dispose(); $captureJob.Runspace.Dispose(); $script:CacheJob = $null
    }
    Update-Buttons
}

function Stop-Observation {
    $script:ObservationStopped = $true
    if ($null -ne $script:CacheJob) { $script:CacheJob.State.Stop = $true }
    if ($null -ne $script:Monitor) { Stop-LiveMonitor 'Monitoring stopped by user' }
    $script:Status.Text = 'Monitoring/cache stop requested. The Intune installation and any registry/service recovery continue.'
    Write-Activity $script:Status.Text
    Update-Buttons
}

function Update-Buttons {
    $BusyBar.Visibility = if ($script:Busy) { 'Visible' } else { 'Collapsed' }
    $BusyText.Visibility = $BusyBar.Visibility
    if (-not $script:Busy) { $BusyText.Text = '' }
    $Connect.IsEnabled = -not $script:Busy
    $Refresh.IsEnabled = -not $script:Busy
    $SyncHelp.IsEnabled = -not $script:Busy
    $History.IsEnabled = -not $script:Busy
    $selected = $Apps.SelectedItem
    $goodGrs = ($null -ne $selected -and ([string]$selected.GRSRegistry).StartsWith('HKEY_LOCAL_MACHINE\', [StringComparison]::OrdinalIgnoreCase))
    $observing = $null -ne $script:Monitor -or $null -ne $script:CacheJob
    $Retry.IsEnabled = (-not $script:Busy -and -not $observing -and $goodGrs)
    $CaptureCache.IsEnabled = -not $script:Busy -and -not $observing
    $StopObservation.IsEnabled = $observing -or ($script:Busy -and $null -ne $script:Job -and $script:Job.Action -eq 'Retry')
    $GraphState.Text = if ($Shared.GraphConnected) { 'Graph: connected' } else { 'Graph: not connected (optional)' }
}

function Start-Work {
    # Only one worker job may run at a time; the GUI timer collects its result.
    param([ValidateSet('Refresh','Connect','Retry','Sync')][string]$Action, $Selected)
    if ($script:Busy) { return }
    $script:Busy = $true
    Update-Buttons
    if ($Action -eq 'Retry') {
        $Shared.LastBackupDir = $null
        $script:RetryLogOffsets = Get-ImeLogOffsets
        $script:ObservationStopped = $false
    }
    if ($Action -eq 'Connect') { $AuthPanel.Visibility = 'Collapsed'; $DeviceCode.Text = ''; $script:DisplayedCode = '' }
    $script:Status.Text = switch ($Action) {
        'Connect' { 'Microsoft Graph: sign-in in progress. The device code will appear in this window shortly.' }
        'Retry' { 'Registry backup / retry / IME check-in in progress ...' }
        'Sync' { 'Triggering MDM sync ...' }
        default { 'Reading local Intune Win32 app entries ...' }
    }
    $Runspace.SessionStateProxy.SetVariable('Action', $Action)
    $Runspace.SessionStateProxy.SetVariable('SelectedApp', $Selected)
    $ps = [PowerShell]::Create()
    $ps.Runspace = $Runspace
    $null = $ps.AddScript(@'
try {
    switch ($Action) {
        'Connect' { Connect-Intune; Get-FailedApps }
        'Refresh' { Get-FailedApps }
        'Retry' { Invoke-AppRetry $SelectedApp; Get-FailedApps }
        'Sync' { Invoke-MdmSync; Get-FailedApps }
    }
} catch {
    Write-Activity ('ERROR: ' + $_.Exception.Message)
    throw
}
'@, $false)
    $script:Job = @{ PS=$ps; Handle=$ps.BeginInvoke(); Action=$Action; Selected=$Selected }
    Update-Buttons
}
$timer = New-Object Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(250)
$timer.Add_Tick({
    Complete-CacheCapture
    if ($script:RefreshPending -and -not $script:Busy) {
        $script:RefreshPending = $false
        Start-Work 'Refresh' $null
    }
    if ($script:Job -and $script:Job.Action -eq 'Connect') {
        $deviceInfo = $Shared.DeviceCodeDisplay
        if ($null -ne $deviceInfo -and -not [string]::IsNullOrWhiteSpace([string]$deviceInfo.Code)) {
            if ($script:DisplayedCode -ne [string]$deviceInfo.Code) {
                $script:DisplayedCode = [string]$deviceInfo.Code
                $DeviceCode.Text = $script:DisplayedCode
                $AuthPanel.Visibility = 'Visible'
                $script:Status.Text = 'Graph sign-in: copy the code and enter it at Microsoft in the browser.'
            }
        }
    }
    $message = ''
    while ($Shared.Messages.TryDequeue([ref]$message)) {
        $Activity.AppendText($message + "`r`n")
        $Activity.ScrollToEnd()
        if ($script:Busy) { $BusyText.Text = ($message -replace '^\d{2}:\d{2}:\d{2}\s+', '') }
    }
    if ($script:Job -and $script:Job.Handle.IsCompleted) {
        $finished = $script:Job
        try {
            $results = @($finished.PS.EndInvoke($finished.Handle))
            if ($finished.PS.HadErrors) { throw ($finished.PS.Streams.Error | Out-String) }
            $script:Rows = @($results | Where-Object { $null -ne $_ -and $null -ne $_.PSObject.Properties['AppId'] })
            $Apps.ItemsSource = $script:Rows
            $count = $script:Rows.Count
            if ($finished.Action -eq 'Retry' -and $null -ne $finished.Selected -and -not $script:ObservationStopped) {
                Start-LiveMonitor $finished.Selected ([string]$Shared.LastBackupDir)
            }
            $script:Status.Text = if ($finished.Action -eq 'Retry') {
                if ($script:ObservationStopped) {
                    Save-RetryOutcome ([string]$Shared.LastBackupDir) 'Monitoring stopped by user before monitoring began'
                    'Local reset completed. Monitoring/cache was stopped by user.'
                } elseif ($null -eq $script:Monitor) {
                    'Local reset completed. Monitoring has finished; see the result below.'
                } else { 'Local reset completed. Monitoring app status; see the log for the sync request result.' }
            } elseif ($finished.Action -eq 'Sync') {
                'MDM sync request processed; see the log for details.'
            } else {
                "Scan complete: $count failed app contexts. Graph connected: $([bool]$Shared.GraphConnected)"
            }
        } catch {
            $script:Status.Text = 'Action failed or completed only partially; see the log for details.'
            if ($finished.Action -eq 'Retry') {
                $Apps.ItemsSource = $null; $script:Rows = @()
                if ($null -ne $script:CacheJob) { $script:CacheJob.State.Stop = $true }
            }
            [Windows.MessageBox]::Show([string]$_.Exception.Message, 'Intune App Retry', 'OK', 'Warning') | Out-Null
        } finally {
            if ($finished.Action -eq 'Connect') {
                $AuthPanel.Visibility = 'Collapsed'
                $DeviceCode.Text = ''
                $script:DisplayedCode = ''
            }
            $finished.PS.Dispose()
            $script:Job = $null
            $script:Busy = $false
            Update-Buttons
        }
    }
})
$Apps.Add_SelectionChanged({
    $row = $Apps.SelectedItem
    $SelectedDetails.Text = if ($null -eq $row) {
        'Please select an app.'
    } else {
        "App: $($row.AppName)   |   App ID: $($row.AppId)   |   Context: $($row.Context)`n" +
        "App registry: $($row.AppRegistry)`nGRS registry: $($row.GRSRegistry)"
    }
    Update-Buttons
})
function Copy-ToClipboard {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return }
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            [Windows.Clipboard]::SetText($Text)
            $script:Status.Text = 'Copied to clipboard.'
            return
        } catch { Start-Sleep -Milliseconds 50 }
    }
    $script:Status.Text = 'Clipboard is busy. Please try again.'
}

function Get-RowValue {
    param($Row, $Column)
    if ($null -eq $Row -or $null -eq $Column -or $null -eq $Column.Binding) { return '' }
    $path = [string]$Column.Binding.Path.Path
    if ($null -eq $Row.PSObject.Properties[$path]) { return '' }
    return [string]$Row.$path
}

function Get-RowText {
    param($Row)
    $lines = foreach ($column in $Apps.Columns) { ([string]$column.Header) + ': ' + (Get-RowValue $Row $column) }
    $lines += 'App registry: ' + [string]$Row.AppRegistry
    return ($lines -join "`r`n")
}
$script:ContextColumn = $null
$Apps.Add_PreviewMouseRightButtonDown({
    param($sender, $e)
    $node = $e.OriginalSource
    while ($null -ne $node -and -not ($node -is [Windows.Controls.DataGridCell])) {
        if ($node -is [Windows.Media.Visual] -or $node -is [Windows.Media.Media3D.Visual3D]) {
            $node = [Windows.Media.VisualTreeHelper]::GetParent($node)
        } else {
            $node = [Windows.LogicalTreeHelper]::GetParent($node)
        }
    }
    if ($null -ne $node) {
        $script:ContextColumn = $node.Column
        $rowContainer = [Windows.Controls.DataGridRow]::GetRowContainingElement($node)
        if ($null -ne $rowContainer) { $Apps.SelectedItem = $rowContainer.Item }
    } else {
        $script:ContextColumn = $null
    }
})
$contextMenu = New-Object Windows.Controls.ContextMenu
$menuCell = New-Object Windows.Controls.MenuItem -Property @{ Header = 'Copy cell' }
$menuName = New-Object Windows.Controls.MenuItem -Property @{ Header = 'Copy app name' }
$menuId = New-Object Windows.Controls.MenuItem -Property @{ Header = 'Copy app ID' }
$menuError = New-Object Windows.Controls.MenuItem -Property @{ Header = 'Copy error code' }
$menuRow = New-Object Windows.Controls.MenuItem -Property @{ Header = 'Copy row details' }
$menuAll = New-Object Windows.Controls.MenuItem -Property @{ Header = 'Copy all rows (table)' }
$menuCell.Add_Click({ Copy-ToClipboard (Get-RowValue $Apps.SelectedItem $script:ContextColumn) })
$menuName.Add_Click({ if ($Apps.SelectedItem) { Copy-ToClipboard ([string]$Apps.SelectedItem.AppName) } })
$menuId.Add_Click({ if ($Apps.SelectedItem) { Copy-ToClipboard ([string]$Apps.SelectedItem.AppId) } })
$menuError.Add_Click({ if ($Apps.SelectedItem) { Copy-ToClipboard ([string]$Apps.SelectedItem.ErrorCode) } })
$menuRow.Add_Click({ if ($Apps.SelectedItem) { Copy-ToClipboard (Get-RowText $Apps.SelectedItem) } })
$menuAll.Add_Click({
    $columns = @($Apps.Columns)
    $lines = @((($columns | ForEach-Object { [string]$_.Header }) -join "`t"))
    foreach ($row in @($script:Rows)) {
        $lines += (($columns | ForEach-Object { Get-RowValue $row $_ }) -join "`t")
    }
    Copy-ToClipboard ($lines -join "`r`n")
})
foreach ($item in @($menuCell, $menuName, $menuId, $menuError)) { $null = $contextMenu.Items.Add($item) }
$null = $contextMenu.Items.Add((New-Object Windows.Controls.Separator))
$null = $contextMenu.Items.Add($menuRow)
$null = $contextMenu.Items.Add($menuAll)
$contextMenu.Add_Opened({
    $hasRow = $null -ne $Apps.SelectedItem
    $menuCell.IsEnabled = $hasRow -and $null -ne $script:ContextColumn
    $menuName.IsEnabled = $hasRow
    $menuId.IsEnabled = $hasRow
    $menuError.IsEnabled = $hasRow
    $menuRow.IsEnabled = $hasRow
    $menuAll.IsEnabled = @($script:Rows).Count -gt 0
})
$Apps.ContextMenu = $contextMenu
$script:RetryLogOffsets = @{}

function Get-ImeLogOffsets {
    # Capture offsets before retry so earlier log entries are excluded where possible.
    $offsets = @{}
    foreach ($name in @('AppWorkload.log', 'IntuneManagementExtension.log')) {
        $path = Join-Path $script:LogsRoot $name
        $offsets[$name] = if (Test-Path -LiteralPath $path) { (Get-Item -LiteralPath $path).Length } else { [long]0 }
    }
    return $offsets
}

function Get-ImeLogStatus {
    # Log evidence is app-specific, not context-specific; treat it as a local indication.
    param([string]$AppId, $Offsets)
    $best = $null
    foreach ($name in @($Offsets.Keys)) {
        $path = Join-Path $script:LogsRoot $name
        if (-not (Test-Path -LiteralPath $path)) { continue }
        $start = [long]$Offsets[$name]
        $len = (Get-Item -LiteralPath $path).Length
        if ($len -lt $start) { $start = 0 }
        if ($len -le $start) { continue }
        $fs = $null; $reader = $null
        try {
            $fs = New-Object IO.FileStream($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
            [void]$fs.Seek($start, [IO.SeekOrigin]::Begin)
            $reader = New-Object IO.StreamReader($fs)
            $text = $reader.ReadToEnd()
        } catch { continue } finally {
            if ($null -ne $reader) { $reader.Dispose() }
            elseif ($null -ne $fs) { $fs.Dispose() }
        }
        foreach ($line in ($text -split "`n")) {
            if ($line.IndexOf($AppId, [StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
            if ($line -match '"ResultantAppState"\s*:\s*(\d+)') {
                $state = [int]$Matches[1]
                $err = if ($line -match '"ErrorCode"\s*:\s*(-?\d+)') { $Matches[1] } else { '-' }
                $best = switch ($state) {
                    1 { [pscustomobject]@{ Status='Installed'; ErrorCode='-'; Enforcement=$null; IsFinal=$true; Detail=('IME log ' + $name + ': ResultantAppState 1') } }
                    2 { [pscustomobject]@{ Status='Failed'; ErrorCode=$err; Enforcement=$null; IsFinal=$true; Detail=('IME log ' + $name + ': ResultantAppState 2, error ' + $err) } }
                    default { [pscustomobject]@{ Status='Installing'; ErrorCode='-'; Enforcement=$null; IsFinal=$false; Detail=('IME log ' + $name + ': ResultantAppState ' + $state) } }
                }
            } elseif ($null -eq $best -or -not $best.IsFinal) {
                $best = [pscustomobject]@{ Status='Installing'; ErrorCode='-'; Enforcement=$null; IsFinal=$false; Detail=('IME log ' + $name + ': app is being processed') }
            }
        }
    }
    return $best
}

function Get-MonitorStatus {
    # Prefer the original context. Other contexts and log entries are fallback evidence.
    param($Monitor)
    $status = $null
    $contexts = @($Monitor.Context)
    if (Test-Path -LiteralPath $script:RegistryRoot) {
        $contexts += @(Get-ChildItem -LiteralPath $script:RegistryRoot -ErrorAction SilentlyContinue | ForEach-Object { [string]$_.PSChildName } | Where-Object { $_ -ne $Monitor.Context -and $_ -ne 'Reporting' -and $_ -ne 'GRS' })
    }
    foreach ($ctx in $contexts) {
        $candidate = Get-LiveStatus $Monitor.AppId $ctx
        if ($candidate.Detail -eq 'no app key') { continue }
        if ($ctx -ne $Monitor.Context) { $candidate.Detail = 'context ' + $ctx + ', ' + $candidate.Detail }
        $status = $candidate
        break
    }
    $log = Get-ImeLogStatus $Monitor.AppId $Monitor.LogOffsets
    if ($null -eq $status) {
        if ($null -ne $log) { return $log }
        return [pscustomobject]@{ Status='Waiting'; ErrorCode='-'; Enforcement=$null; IsFinal=$false; Detail='no app key in any context, no new IME log entry' }
    }
    if (-not $status.IsFinal -and $null -ne $log -and $log.IsFinal) { return $log }
    return $status
}
$script:MonitorInterval = [TimeSpan]::FromSeconds(15)
$script:MonitorTimeout = [TimeSpan]::FromMinutes(30)
$script:Monitor = $null
function Get-LiveStatus {
    # Keep monitor output stable even when the app key has not been recreated yet.
    param([string]$AppId, [string]$Context)
    $result = [pscustomobject]@{
        Status = 'Waiting'; ErrorCode = '-'; Enforcement = $null
        IsFinal = $false; Detail = 'no app key'
    }
    if (-not (Test-AppContext $Context)) { return $result }
    $key = @(Get-SpecificAppKeys -Context $Context -AppId $AppId |
        Sort-Object @{ Expression = {
            if ($_.PSChildName -match '_(\d+)$') { [long]$Matches[1] } else { [long]0 }
        } } -Descending) | Select-Object -First 1
    if ($null -eq $key) { return $result }

    $state = Get-AppRegistryState -Path $key.PSPath
    $result.Status = if ($state.Status -eq 'Unknown') { 'Waiting' } else { $state.Status }
    $result.ErrorCode = $state.ErrorCode
    $result.Enforcement = $state.Enforcement
    $result.IsFinal = $state.IsFinal
    $result.Detail = 'key ' + $key.PSChildName + ', compliance ' + [string]$state.Compliance +
        ', enforcement ' + [string]$state.Enforcement + ', codes ' + ($state.Codes -join '/')
    return $result
}

function Save-RetryOutcome {
    # Monitoring outcomes belong to the original backup, not the currently selected grid row.
    param([string]$BackupDir, [string]$Outcome)
    if ([string]::IsNullOrWhiteSpace($BackupDir)) { return }
    $manifestPath = Join-Path $BackupDir 'manifest.json'
    if (-not (Test-Path -LiteralPath $manifestPath)) { return }
    try {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $manifest | Add-Member -NotePropertyName Outcome -NotePropertyValue $Outcome -Force
        $manifest | Add-Member -NotePropertyName OutcomeTime -NotePropertyValue ((Get-Date).ToString('o')) -Force
        $manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    } catch { Write-Activity ('Retry history: outcome could not be saved: ' + $_.Exception.Message) }
}

function Stop-LiveMonitor {
    # Stopping observation must not interrupt the installation itself.
    param([string]$Outcome)
    $m = $script:Monitor
    if ($null -eq $m) { return }
    $m.Timer.Stop()
    $script:Monitor = $null
    if ([string]::IsNullOrWhiteSpace($Outcome)) { $Outcome = 'Monitoring stopped (last: ' + $(if ($m.LastStatus) { $m.LastStatus } else { '-' }) + ')' }
    Save-RetryOutcome $m.BackupDir $Outcome
    Write-Activity ('Live monitor finished: ' + $m.AppName + ' -> ' + $Outcome)
    $MonitorTitle.Text = 'Live monitor: ' + $m.AppName + ' - ' + $Outcome
    $StopMonitor.Content = 'Hide'
    if ($null -ne $script:CacheJob) { $script:CacheJob.State.Finish = $true }
    $script:RefreshPending = $true
    Update-Buttons
}

function Invoke-MonitorTick {
    try { Invoke-MonitorTickCore } catch {
        $monitorError = $_.Exception.Message
        Write-Activity ('Live monitor error: ' + $monitorError + ' (line ' + $_.InvocationInfo.ScriptLineNumber + ')')
        # A broken observer must not leave a repeating timer and disabled controls behind.
        if ($null -ne $script:Monitor) {
            $script:Monitor.Timer.Stop()
            Stop-LiveMonitor ('Monitoring error: ' + $monitorError)
        }
        Update-Buttons
    }
}

function Invoke-MonitorTickCore {
    $m = $script:Monitor
    if ($null -eq $m) { return }
    try { $liveState = Get-MonitorStatus $m }
    catch {
        Write-Activity ('Live monitor: status could not be read: ' + $_.Exception.Message)
        $liveState = [pscustomobject]@{ Status='Unknown'; ErrorCode='-'; Enforcement=$null; IsFinal=$false; Detail='read error' }
    }
    $elapsed = (Get-Date) - $m.Started
    if ($liveState.Detail -ne $m.LastDetail) { $m.LastDetail = $liveState.Detail; Write-Activity ('Live monitor detail: ' + $liveState.Detail) }
    if ($liveState.Status -ne $m.LastStatus) {
        $m.LastStatus = $liveState.Status
        $suffix = if ($liveState.Status -eq 'Failed') { ' (' + $liveState.ErrorCode + ')' } else { '' }
        $m.Steps.Add((Get-Date -Format 'HH:mm:ss') + ' ' + $liveState.Status + $suffix)
        $detail = if ($null -ne $liveState.Enforcement) { ' (enforcement ' + $liveState.Enforcement + ')' } else { '' }
        Write-Activity ('Live monitor: ' + $m.AppName + ' -> ' + $liveState.Status + $detail + $(if ($liveState.Status -eq 'Failed') { ', error ' + $liveState.ErrorCode } else { '' }))
    }
    $MonitorSteps.Text = ($m.Steps -join '  >  ') + '   |   elapsed ' + $elapsed.ToString('mm\:ss')
    if ($liveState.IsFinal) {
        Stop-LiveMonitor $(if ($liveState.Status -eq 'Installed') { 'Installed' } else { 'Failed again (' + $liveState.ErrorCode + ')' })
    } elseif ($elapsed -ge $script:MonitorTimeout) {
        Stop-LiveMonitor ('Timeout after ' + $script:MonitorTimeout.TotalMinutes + ' min (last: ' + $liveState.Status + ')')
    }
}

function Start-LiveMonitor {
    param($Row, [string]$BackupDir)
    if ($null -ne $script:Monitor) { Stop-LiveMonitor $null }
    $timer = New-Object Windows.Threading.DispatcherTimer
    $timer.Interval = $script:MonitorInterval
    $script:Monitor = @{
        Timer = $timer; AppId = [string]$Row.AppId; Context = [string]$Row.Context; AppName = [string]$Row.AppName
        BackupDir = $BackupDir; Started = Get-Date; LastStatus = $null; LastDetail = $null; LogOffsets = $script:RetryLogOffsets; Steps = (New-Object 'System.Collections.Generic.List[string]')
    }
    $MonitorTitle.Text = 'Live monitor: ' + $Row.AppName
    $MonitorSteps.Text = 'Waiting for the IME to pick up the app ...'
    $StopMonitor.Content = 'Stop monitoring'
    $MonitorPanel.Visibility = 'Visible'
    Write-Activity ('Live monitor started for ' + $Row.AppName + ' (checks every ' + $script:MonitorInterval.TotalSeconds + ' s, max ' + $script:MonitorTimeout.TotalMinutes + ' min).')
    $timer.Add_Tick({ Invoke-MonitorTick })
    $timer.Start()
    Invoke-MonitorTick
}
$StopMonitor.Add_Click({
    if ($null -ne $script:Monitor) { Stop-Observation; return }
    $MonitorPanel.Visibility = 'Collapsed'
    $StopMonitor.Content = 'Stop monitoring'
})
function Get-RetryHistory {
    # A corrupt manifest must not prevent the remaining history from being displayed.
    $backups = Join-Path $DataRoot 'Backups'
    if (-not (Test-Path -LiteralPath $backups)) { return @() }
    $entries = foreach ($folder in @(Get-ChildItem -LiteralPath $backups -Directory -ErrorAction SilentlyContinue)) {
        $manifestPath = Join-Path $folder.FullName 'manifest.json'
        if (-not (Test-Path -LiteralPath $manifestPath)) { continue }
        try {
            $m = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
            $time = [datetime]::MinValue
            [void][datetime]::TryParse([string]$m.Time, [ref]$time)
            $outcome = if ($null -ne $m.PSObject.Properties['Outcome']) { [string]$m.Outcome } else { '(not monitored)' }
            [pscustomobject]@{
                SortTime = $time; Time = $(if ($time -eq [datetime]::MinValue) { '-' } else { $time.ToString('yyyy-MM-dd HH:mm:ss') })
                Operator = [string]$m.Operator; Computer = [string]$m.Computer; AppName = [string]$m.AppName
                Outcome = $outcome; AppId = [string]$m.AppId; Context = [string]$m.Context; Folder = $folder.FullName
            }
        } catch {
            [pscustomobject]@{ SortTime = [datetime]::MinValue; Time = '-'; Operator = '-'; Computer = '-'; AppName = '-'; Outcome = '(manifest unreadable)'; AppId = '-'; Context = '-'; Folder = $folder.FullName }
        }
    }
    return @($entries | Sort-Object SortTime -Descending)
}
$History.Add_Click({
    try {
        $entries = Get-RetryHistory
        [xml]$hx = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Title="Intune App Retry - Retry history" Width="1150" Height="560" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI">
 <DockPanel Margin="16">
  <TextBlock x:Name="HistInfo" DockPanel.Dock="Top" Margin="0,0,0,8" TextWrapping="Wrap"/>
  <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
   <Button x:Name="HistOpen" Content="Open backup folder" Padding="14,6" Margin="0,0,8,0"/>
   <Button x:Name="HistClose" Content="Close" Padding="14,6" IsCancel="True"/>
  </StackPanel>
  <DataGrid x:Name="HistGrid" IsReadOnly="True" AutoGenerateColumns="False" SelectionMode="Single" CanUserAddRows="False" HeadersVisibility="Column" ClipboardCopyMode="IncludeHeader" AlternatingRowBackground="#F2F6FB" AlternationCount="2">
   <DataGrid.Columns>
    <DataGridTextColumn Header="Time" Binding="{Binding Time}" Width="140"/>
    <DataGridTextColumn Header="Operator" Binding="{Binding Operator}" Width="110"/>
    <DataGridTextColumn Header="Computer" Binding="{Binding Computer}" Width="120"/>
    <DataGridTextColumn Header="App" Binding="{Binding AppName}" Width="220"/>
    <DataGridTextColumn Header="Outcome" Binding="{Binding Outcome}" Width="220"/>
    <DataGridTextColumn Header="App ID" Binding="{Binding AppId}" Width="260"/>
    <DataGridTextColumn Header="Context" Binding="{Binding Context}" Width="260"/>
   </DataGrid.Columns>
  </DataGrid>
 </DockPanel>
</Window>
"@
        $hw = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader($hx)))
        $hw.Owner = $window
        $histGrid = $hw.FindName('HistGrid')
        $histGrid.ItemsSource = @($entries)
        $hw.FindName('HistInfo').Text = "$(@($entries).Count) retries recorded. Outcome is written by the live monitor after a retry. Double-click a row to open its backup folder."
        $openFolder = {
            $entry = $histGrid.SelectedItem
            if ($null -ne $entry -and -not [string]::IsNullOrWhiteSpace([string]$entry.Folder)) {
                Start-Process explorer.exe -ArgumentList ('"' + [string]$entry.Folder + '"')
            }
        }.GetNewClosure()
        $hw.FindName('HistOpen').Add_Click($openFolder)
        $histGrid.Add_MouseDoubleClick($openFolder)
        $closeWindow = { $hw.Close() }.GetNewClosure()
        $hw.FindName('HistClose').Add_Click($closeWindow)
        $null = $hw.ShowDialog()
    } catch {
        [Windows.MessageBox]::Show([string]$_.Exception.Message, 'Intune App Retry', 'OK', 'Warning') | Out-Null
    }
})
$StopObservation.Add_Click({ Stop-Observation })
$Refresh.Add_Click({ Start-Work 'Refresh' $null })
$Connect.Add_Click({ Start-Work 'Connect' $null })
$CopyCode.Add_Click({
    if (-not [string]::IsNullOrWhiteSpace($DeviceCode.Text)) {
        [Windows.Clipboard]::SetText([string]$DeviceCode.Text)
        $script:Status.Text = 'Code copied. Now enter it at the Microsoft login.'
    }
})
$OpenLogin.Add_Click({
    $info = $Shared.DeviceCodeDisplay
    if ($null -ne $info -and ([string]$info.Url) -match '^https://(?:www\.)?microsoft\.com/devicelogin/?$') {
        Start-Process ([string]$info.Url)
    } else {
        Start-Process 'https://microsoft.com/devicelogin'
    }
})
$Logs.Add_Click({ Start-Process explorer.exe -ArgumentList ('"' + $DataRoot + '"') })
$SyncHelp.Add_Click({ Start-Work 'Sync' $null })
$Retry.Add_Click({
    $row = $Apps.SelectedItem
    if ($null -eq $row) { return }
    if (-not ([string]$row.GRSRegistry).StartsWith('HKEY_LOCAL_MACHINE\', [StringComparison]::OrdinalIgnoreCase)) {
        [Windows.MessageBox]::Show('GRS key is not unique. Please rescan.', 'Retry blocked', 'OK', 'Warning') | Out-Null
        return
    }
    $prompt = @"
Really reset the selected Win32 app context?

App: $($row.AppName)
App ID: $($row.AppId)
Context: $($row.Context)

APP KEY:
$($row.AppRegistry)

GRS KEY:
$($row.GRSRegistry)

Cache capture: $([bool]$CaptureCache.IsChecked)
When enabled, available files for this app are copied to C:\Temp\IntuneAppRetry-Cache.

Both keys are backed up as .reg files before deletion.
The Intune Management Extension is restarted afterwards.
The IME restart also affects other app evaluations on this client.

Unsupported registry workaround. No uninstall, no installation guarantee.
"@
    if ([Windows.MessageBox]::Show($prompt, 'Confirm registry reset', 'YesNo', 'Warning', 'No') -eq 'Yes') {
        try {
            if ([bool]$CaptureCache.IsChecked) { Start-InstallationCacheCapture ([string]$row.AppId) }
            Start-Work 'Retry' $row
        } catch {
            if ($null -ne $script:CacheJob) { $script:CacheJob.State.Stop = $true }
            $script:Busy = $false
            Update-Buttons
            Write-Activity ('Retry could not be started: ' + $_.Exception.Message)
            [Windows.MessageBox]::Show([string]$_.Exception.Message, 'Retry not started', 'OK', 'Warning') | Out-Null
        }
    }
})
$window.Add_Closing({
    param($sender, $args)
    if ($null -ne $script:CacheJob) {
        Stop-Observation
        $args.Cancel = $true
        $script:Status.Text = 'Stopping cache capture. Close the window again after it finishes.'
        return
    }
    if ($script:Busy) {
        $args.Cancel = $true
        $script:Status.Text = 'Operation in progress. Please wait. Device-code sign-in ends after authorization, refusal, or code expiry.'
    }
})
$window.Add_ContentRendered({
    if (-not $script:Started) {
        $script:Started = $true
        Start-Work 'Refresh' $null
    }
})
$timer.Start()
try { $null = $window.ShowDialog() }
finally {
    $timer.Stop()
    if ($null -ne $script:CacheJob) {
        $script:CacheJob.State.Stop = $true
        # The worker checks Stop every 200 ms and terminates only its own copy process.
        $null = $script:CacheJob.Handle.AsyncWaitHandle.WaitOne(5000)
        if ($script:CacheJob.Handle.IsCompleted) { Complete-CacheCapture }
    }
    if ($null -ne $script:Monitor) { Stop-LiveMonitor 'Window closed by user' }
    if ($null -ne $script:Job) { $script:Job.PS.Dispose() }
    $Runspace.Close()
    $Runspace.Dispose()
}
