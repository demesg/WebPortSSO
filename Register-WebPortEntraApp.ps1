#Requires -Version 7.0
<#
.SYNOPSIS
Registers a single-tenant WebPort Web application with app roles and grants delegated Microsoft Graph consent.
.DESCRIPTION
Before any change, the script lists every planned Entra setting and every SQLite value as current → new.
Run with -WhatIf to see that plan without signing in or changing anything.
Uses interactive Microsoft Graph authentication. Creates the app registration, service principal,
tenant-wide delegated consent, Entra Application Proxy publishing through an existing active connector
group and, by default, a client secret with the maximum lifetime of 24 months. Application Proxy publishes https://<SSLCN>/ (internal)
on https://<plant>-<tenant>.msappproxy.net/ (external, or -ApplicationProxyExternalUrl). The add-on's
callback is <external>/access/loginsso and is saved in SQLite as EntraRedirectUri; the external root is
also registered as a redirect URI for Application Proxy pre-authentication, and the internal SSLCN callback
for sign-in directly on the local network. Without any Application
Proxy connector in the tenant, or with -SkipApplicationProxy, the app is registered without Application Proxy
and the callback stays on https://<SSLCN>/. Never prints or saves the
secret as plaintext. ClientSecret in the returned object is a SecureString. Reads WebPort SQLite by
default. UpdateDatabase (the default) stops WebPort, backs up SQLite, writes the Entra settings and stores
the new client secret in SQLite (EntraClientSecretProtected), encrypted with DPAPI for the account WebPort
runs as: LocalSystem for the WebPortServer service (by a short-lived scheduled task), or the current user
when WebPort runs as a program of that user. WebPort is restarted, as service or program, when it was
running and the add-on is already installed; otherwise Install-EntraModule.ps1 starts it. Requires
administrator rights.
Creates a new application unless its name is already used. A name match is reused automatically when
its client ID matches SQLite KIClientId; Force permits updating an otherwise unambiguous name match.
Reads Plant, SSLCN and the settings DeviceGuid from SQLite.
Defines WebPort.Deny/View/Basic/Advanced/System/Admin for user and group assignment.
It does not assign these roles to users or enable role-based login in WebPort.
Use -WhatIf for a preview without authentication or Graph requests.
.EXAMPLE
$registration = .\Register-WebPortEntraApp.ps1
$registration | Select-Object TenantId,ClientId,RedirectUri,Scopes,AdminConsentGranted,SecretStatus
.EXAMPLE
$registration = .\Register-WebPortEntraApp.ps1 -UpdateDatabase
Stop WebPort, save the registration settings and the encrypted client secret to SQLite after successful registration.
.EXAMPLE
.\Register-WebPortEntraApp.ps1 -ReadDatabaseOnly
Read effective settings without Graph authentication, database writes or process shutdown.
.EXAMPLE
$registration = .\Register-WebPortEntraApp.ps1 -Force -UpdateDatabase:$false
Update one compatible app with the same facility name, preserving role IDs and existing secrets.
Adds a new client secret unless SkipClientSecret is specified. Creates the app if none exists.
.NOTES
Requires Microsoft.Graph.Authentication. The operator needs a role permitted to manage applications,
Application Proxy and the requested delegated consent, such as Application Administrator.
The Graph administration scopes are NOT assigned to the new WebPort app.
#>
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
    [guid]$TenantId = [guid]::Empty,
    [Alias('FacilityName','DisplayName')][string]$Plant,
    [string]$WebPortProgPath = 'C:\Program Files\WebPort',
    [string]$WebPortDataPath = 'C:\ProgramData\WebPort',
    [string]$DatabasePath,
    [string]$Scopes,
    [switch]$ReadDatabaseOnly,
    [switch]$UpdateDatabase,
    [switch]$WithoutGraphGroups,
    [switch]$SkipClientSecret,
    [switch]$Force,
    [string]$ApplicationProxyConnectorGroup,
    [string]$ApplicationProxyExternalUrl,
    [switch]$SkipApplicationProxy,
    [ValidateRange(1,24)][int]$SecretLifetimeMonths = 24,
    [switch]$UseDeviceCode,
    [switch]$UseExistingConnection
)

$banner = @'

__        __   _     ____            _     ____ ____   ___
\ \      / /__| |__ |  _ \ ___  _ __| |_  / ___/ ___| / _ \
 \ \ /\ / / _ \ '_ \| |_) / _ \| '__| __| \___ \___ \| | | |
  \ V  V /  __/ |_) |  __/ (_) | |  | |_   ___) |__) | |_| |
   \_/\_/ \___|_.__/|_|   \___/|_|   \__| |____/____/ \___/


'@
Write-Host $banner -ForegroundColor Magenta

function step($t){ Write-Host "==> $t" -ForegroundColor Cyan }
function ok  ($t){ Write-Host "   ✓ $t" -ForegroundColor Green }
function warn($t){ Write-Host "   ⚠ $t" -ForegroundColor Yellow }
function err ($t){ Write-Host "   ✖ $t" -ForegroundColor Red }
function info($t){ Write-Host "   ℹ $t" -ForegroundColor White }
function setting($name,$value){ Write-Host ("   {0,-24} {1}" -f ($name+':'),$value) }
function detail($value){ Write-Host ("   {0,-24} {1}" -f '',$value) }
function Format-SettingValue($value){ if($null -eq $value){'(saknas)'}elseif([string]::IsNullOrEmpty([string]$value)){'(tom)'}else{[string]$value} }

$ErrorActionPreference = 'Stop'
if (-not $PSBoundParameters.ContainsKey('UpdateDatabase')) { $UpdateDatabase=-not $ReadDatabaseOnly }
if ($ReadDatabaseOnly -and $UpdateDatabase) { throw 'ReadDatabaseOnly cannot update the database' }

function New-WebPortIamSqliteConnection {
    param([Parameter(Mandatory=$true)][string]$DatabasePath,
          [Parameter(Mandatory=$true)][string]$WebPortProgPath,
          [ValidateSet('ReadOnly','ReadWrite','Create')][string]$Mode='ReadOnly')
    if ($Mode -ne 'Create' -and -not (Test-Path -LiteralPath $DatabasePath -PathType Leaf)) { throw "SQLite database missing: $DatabasePath" }
    $microsoft=Join-Path $WebPortProgPath 'Microsoft.Data.Sqlite.dll'
    $system=Join-Path $WebPortProgPath 'System.Data.SQLite.dll'
    if ((Test-Path -LiteralPath $microsoft) -and $PSVersionTable.PSVersion.Major -ge 7) {
        if (-not ('Microsoft.Data.Sqlite.SqliteConnection' -as [type])) {
            foreach ($file in @('SQLitePCLRaw.core.dll','SQLitePCLRaw.provider.dynamic_cdecl.dll','SQLitePCLRaw.batteries_v2.dll','Microsoft.Data.Sqlite.dll')) {
                [void][Reflection.Assembly]::LoadFrom((Join-Path $WebPortProgPath $file))
            }
            [SQLitePCL.Batteries_V2]::Init()
        }
        $builder=[Microsoft.Data.Sqlite.SqliteConnectionStringBuilder]::new()
        $builder.DataSource=$DatabasePath
        $builder.Pooling=$false
        $builder['Default Timeout']=10
        $builder.Mode=[Microsoft.Data.Sqlite.SqliteOpenMode]::ReadOnly
        if ($Mode -eq 'ReadWrite') { $builder.Mode=[Microsoft.Data.Sqlite.SqliteOpenMode]::ReadWrite }
        if ($Mode -eq 'Create') { $builder.Mode=[Microsoft.Data.Sqlite.SqliteOpenMode]::ReadWriteCreate }
        return [Microsoft.Data.Sqlite.SqliteConnection]::new($builder.ConnectionString)
    }
    if (Test-Path -LiteralPath $system) {
        if (-not ('System.Data.SQLite.SQLiteConnection' -as [type])) { [void][Reflection.Assembly]::LoadFrom($system) }
        $builder=[Data.Common.DbConnectionStringBuilder]::new()
        $builder['Data Source']=$DatabasePath; $builder['Pooling']=$false; $builder['Default Timeout']=10
        $builder['Read Only']=($Mode -eq 'ReadOnly'); $builder['FailIfMissing']=($Mode -ne 'Create')
        return [System.Data.SQLite.SQLiteConnection]::new($builder.ConnectionString)
    }
    throw "No SQLite provider found in $WebPortProgPath"
}

function Get-WebPortSettingsDeviceGuid {
    param([string]$DatabasePath,[string]$WebPortProgPath)
    $connection=New-WebPortIamSqliteConnection $DatabasePath $WebPortProgPath
    try {
        $connection.Open(); $command=$connection.CreateCommand()
        try {
            $command.CommandText='SELECT DISTINCT "DeviceGuid", "Key" FROM "Settings" WHERE "Key" IN (''SSLCN'',''Plant'') COLLATE NOCASE'
            $reader=$command.ExecuteReader(); $certificateGuids=@(); $plantGuids=@()
            try { while($reader.Read()) {
                $guid=[guid]::Empty
                if(-not [guid]::TryParse($reader.GetString(0),[ref]$guid) -or $guid -eq [guid]::Empty) { throw 'Invalid DeviceGuid for Plant/SSLCN in SQLite' }
                if($reader.GetString(1) -ieq 'SSLCN') {$certificateGuids+=$guid.ToString('D')}
                else {$plantGuids+=$guid.ToString('D')}
            }} finally {$reader.Dispose()}
            $certificateGuids=@($certificateGuids | Select-Object -Unique)
            $plantGuids=@($plantGuids | Select-Object -Unique)
            if($certificateGuids.Count -eq 0) {throw 'SSLCN saknas i WebPorts databas. Kör LetsEncryptWebport först och kör sedan registreringsskriptet igen.'}
            $candidates=$certificateGuids
            if($candidates.Count -gt 1) {$candidates=@($certificateGuids | Where-Object {$plantGuids -contains $_})}
            if($candidates.Count -ne 1) {throw 'Kan inte identifiera ett entydigt SettingsDeviceGuid från Plant/SSLCN i databasen. Ingen registrering utförs.'}
            $selected=[guid]::Empty
            if(-not [guid]::TryParse($candidates[0],[ref]$selected) -or $selected -eq [guid]::Empty){throw 'Invalid Settings DeviceGuid in SQLite'}
            return $selected
        }finally{$command.Dispose()}
    }finally{$connection.Dispose()}
}

function Read-WebPortIamSettings {
    param([string]$DatabasePath,[string]$WebPortProgPath,[guid]$SettingsDeviceGuid)
    $connection=New-WebPortIamSqliteConnection $DatabasePath $WebPortProgPath
    try {
        $connection.Open()
        $command=$connection.CreateCommand()
        try {
            # DeviceGuid is the settings namespace, not a random LIMIT 1 row from another device.
            $command.CommandText='SELECT "Key","Value" FROM "Settings" WHERE "DeviceGuid" = @dg COLLATE NOCASE'
            [void]$command.Parameters.AddWithValue('@dg',$SettingsDeviceGuid.ToString('D'))
            $reader=$command.ExecuteReader()
            $values=@{}
            $allowed=@('Plant','SSLCN','WebPortServer Url','KIClientId','KIExternalTenantId','KITenantId','KIProviderUri','KICustomPolicy','KISignInPolicyName','EntraRedirectUri','EntraRolesEnabled','AuthenticationTypeKI')
            $rows=0
            try { while ($reader.Read()) {
                $rows++
                $key=$reader.GetString(0)
                if ($allowed -notcontains $key) { continue }
                if ($values.ContainsKey($key)) { throw "Duplicate setting '$key' in the selected DeviceGuid namespace" }
                $values[$key]=if($reader.IsDBNull(1)){''}else{$reader.GetString(1)}
            }} finally { $reader.Dispose() }
            if ($rows -eq 0) { throw "No settings found for DeviceGuid $SettingsDeviceGuid; refusing to guess another namespace" }
            return $values
        } finally { $command.Dispose() }
    } finally { $connection.Dispose() }
}

function Resolve-WebPortIamDefaults {
    param([hashtable]$Settings)
    if([string]::IsNullOrWhiteSpace($Settings['SSLCN'])) {
        throw 'SSLCN saknas i WebPorts databas. Kör LetsEncryptWebport först och kör sedan registreringsskriptet igen.'
    }
    $hostName=$Settings['SSLCN'].Trim().ToLowerInvariant()
    if([uri]::CheckHostName($hostName) -ne [UriHostNameType]::Dns -or $hostName.Contains('*')) {
        throw 'SSLCN måste vara ett specifikt DNS-värdnamn. Kör LetsEncryptWebport med rätt servernamn först.'
    }
    $tenantCandidates=@()
    if(-not [string]::IsNullOrWhiteSpace($Settings['KIExternalTenantId'])) {
        $guid=[guid]::Empty
        if(-not [guid]::TryParse($Settings['KIExternalTenantId'],[ref]$guid) -or $guid -eq [guid]::Empty) {throw 'Invalid KIExternalTenantId in SQLite'}
        $tenantCandidates=@($guid.ToString('D'))
    }
    [pscustomobject]@{
        TenantCandidates=$tenantCandidates
        Plant=$Settings['Plant']
        ServerUrl='https://'+$hostName+'/'
        RedirectUri='https://'+$hostName+'/access/loginsso'
    }
}

function Write-WebPortIamSettings {
    param([string]$DatabasePath,[string]$WebPortProgPath,[guid]$SettingsDeviceGuid,
          [Parameter(Mandatory=$true)][Collections.IDictionary]$Values,
          [Parameter(Mandatory=$true)][hashtable]$ExpectedSettings)
    $allowed=@('Plant','SSLCN','WebPortServer Url','KIClientId','KIExternalTenantId','KITenantId','KIProviderUri','KICustomPolicy','KISignInPolicyName','EntraRedirectUri','AuthenticationTypeKI')
    foreach($key in $Values.Keys) {if($allowed -notcontains $key) {throw "Not a supported native WebPort setting: $key"}}
    $source=New-WebPortIamSqliteConnection $DatabasePath $WebPortProgPath -Mode ReadWrite
    $backupPath=$DatabasePath+'.EntraIAM-backup-'+[DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss-fff')+'-'+[guid]::NewGuid().ToString('N')+'.sqlite'
    $destination=$null; $transaction=$null
    try {
        $source.Open()
        $destination=New-WebPortIamSqliteConnection $backupPath $WebPortProgPath -Mode Create
        $destination.Open()
        # SQLite's backup API includes committed WAL data; copying only the .sqlite file would not.
        if ($source.GetType().FullName -eq 'Microsoft.Data.Sqlite.SqliteConnection') { $source.BackupDatabase($destination) }
        else { $source.BackupDatabase($destination,'main','main',-1,$null,0) }
        $destination.Dispose(); $destination=$null
        $transaction=$source.BeginTransaction()
        $check=$source.CreateCommand()
        try {
            $check.Transaction=$transaction
            $check.CommandText='SELECT "Value" FROM "Settings" WHERE "DeviceGuid"=@dg COLLATE NOCASE AND "Key"=''SSLCN'' COLLATE NOCASE'
            [void]$check.Parameters.AddWithValue('@dg',$SettingsDeviceGuid.ToString('D'))
            $reader=$check.ExecuteReader();$certificateNames=@()
            try{while($reader.Read()){$certificateNames+=$reader.GetString(0)}}finally{$reader.Dispose()}
            if($certificateNames.Count -ne 1 -or -not $ExpectedSettings.ContainsKey('SSLCN') -or $certificateNames[0] -cne $ExpectedSettings['SSLCN']) {
                throw 'SSLCN or the selected settings namespace changed during registration; SQLite was not updated'
            }
        }finally{$check.Dispose()}
        foreach($key in $Values.Keys) {
            if ($key -match '(?i)secret(?!keyid)|password|token') { throw "Refusing to store a credential in SQLite: $key" }
            $command=$source.CreateCommand()
            try {
                $command.Transaction=$transaction
                $command.CommandText='SELECT "Value" FROM "Settings" WHERE "DeviceGuid"=@dg COLLATE NOCASE AND "Key"=@k COLLATE NOCASE'
                [void]$command.Parameters.AddWithValue('@dg',$SettingsDeviceGuid.ToString('D'))
                [void]$command.Parameters.AddWithValue('@k',[string]$key)
                $reader=$command.ExecuteReader(); $existing=@()
                try { while($reader.Read()) { $existing += $reader.GetString(0) } } finally {$reader.Dispose()}
                if ($existing.Count -gt 1) { throw "Duplicate SQLite setting: $key" }
                if ($ExpectedSettings.ContainsKey($key)) {
                    if ($existing.Count -ne 1 -or $existing[0] -cne $ExpectedSettings[$key]) { throw "SQLite setting changed since registration started: $key" }
                } elseif ($existing.Count) { throw "SQLite setting appeared since registration started: $key" }
                [void]$command.Parameters.AddWithValue('@v',[string]$Values[$key])
                $command.CommandText=if($existing.Count){'UPDATE "Settings" SET "Value"=@v WHERE "DeviceGuid"=@dg COLLATE NOCASE AND "Key"=@k COLLATE NOCASE'}else{'INSERT INTO "Settings" ("DeviceGuid","Key","Value") VALUES (@dg,@k,@v)'}
                if ($command.ExecuteNonQuery() -ne 1) { throw "Unexpected affected-row count for $key" }
            } finally { $command.Dispose() }
        }
        $transaction.Commit()
    } catch {
        if ($null -ne $transaction) { try {$transaction.Rollback()} catch {} }
        throw
    } finally {
        if ($null -ne $transaction) {$transaction.Dispose()}
        if ($null -ne $destination) {$destination.Dispose()}
        $source.Dispose()
    }
    [pscustomobject]@{Updated=$true;BackupPath=$backupPath;Keys=@($Values.Keys)}
}

function Stop-WebPortInstallation {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$InstallationPath)

    Add-Type -AssemblyName System.ServiceProcess
    $executable = Join-Path $InstallationPath 'WebPortServer.exe'
    if (-not (Test-Path -LiteralPath $executable -PathType Leaf)) { throw "WebPort executable missing: $executable" }
    $executable = (Resolve-Path -LiteralPath $executable).Path
    $servicePattern = '^\s*(?:"' + [regex]::Escape($executable) + '"|' + [regex]::Escape($executable) + ')(?=\s|$)'
    $processes = @(Get-CimInstance -ClassName Win32_Process -Filter "Name = 'WebPortServer.exe'" -ErrorAction Stop)
    if (@($processes | Where-Object { -not $_.ExecutablePath }).Count) {
        throw 'Cannot determine a WebPortServer process path; refusing to replace potentially locked files'
    }
    $processes = @($processes | Where-Object { [string]::Equals($_.ExecutablePath, $executable, [StringComparison]::OrdinalIgnoreCase) })
    $processIds = @($processes | ForEach-Object { [uint32]$_.ProcessId })
    $services = @(Get-CimInstance -ClassName Win32_Service -ErrorAction Stop | Where-Object {
        ([Environment]::ExpandEnvironmentVariables([string]$_.PathName) -match $servicePattern) -or
        ($_.ProcessId -ne 0 -and $processIds -contains [uint32]$_.ProcessId)
    })
    foreach ($service in $services) {
        $controller = Get-Service -Name $service.Name -ErrorAction Stop
        if ($controller.Status -ne 'Stopped') {
            Write-Host "Stopping WebPort service: $($service.Name)"
            if ($controller.Status -ne 'StopPending') { $controller.Stop() }
            $controller.WaitForStatus([System.ServiceProcess.ServiceControllerStatus]::Stopped, [TimeSpan]::FromSeconds(30))
        }
        $controller.Refresh()
        if ($controller.Status -ne 'Stopped') { throw "WebPort service did not stop: $($service.Name)" }
    }
    # Re-enumerate after service shutdown; its process may already have exited.
    $remaining = @(Get-CimInstance -ClassName Win32_Process -Filter "Name = 'WebPortServer.exe'" -ErrorAction Stop)
    foreach ($item in $remaining) {
        if (-not $item.ExecutablePath) { throw 'Cannot determine remaining WebPortServer process path' }
        if (-not [string]::Equals($item.ExecutablePath, $executable, [StringComparison]::OrdinalIgnoreCase)) { continue }
        $process = Get-Process -Id $item.ProcessId -ErrorAction SilentlyContinue
        if ($null -eq $process) { continue }
        if (-not [string]::Equals($process.Path, $executable, [StringComparison]::OrdinalIgnoreCase)) { throw 'WebPort process identity changed during shutdown' }
        Write-Host "Stopping WebPortServer process: $($process.Id)"
        $closed = $process.CloseMainWindow()
        if (-not ($closed -and $process.WaitForExit(15000))) {
            # Refresh identity immediately before terminating a remaining background process.
            $current = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $($process.Id)" -ErrorAction Stop
            if ($null -ne $current) {
                if (-not [string]::Equals($current.ExecutablePath, $executable, [StringComparison]::OrdinalIgnoreCase)) { throw 'WebPort PID was reused; shutdown aborted' }
                Stop-Process -InputObject $process -Force -ErrorAction Stop
                if (-not $process.WaitForExit(10000)) { throw 'WebPortServer did not exit' }
            }
        }
    }
    $finalProcesses = @(Get-CimInstance -ClassName Win32_Process -Filter "Name = 'WebPortServer.exe'" -ErrorAction Stop)
    if (@($finalProcesses | Where-Object { -not $_.ExecutablePath -or [string]::Equals($_.ExecutablePath, $executable, [StringComparison]::OrdinalIgnoreCase) }).Count) {
        throw 'WebPortServer is still running or restarted; no patch files will be installed'
    }
    foreach ($service in $services) {
        if ((Get-Service -Name $service.Name -ErrorAction Stop).Status -ne 'Stopped') { throw "WebPort service restarted: $($service.Name)" }
    }
    Write-Host 'WebPortServer and its matching services are stopped.'
}

# WebPort runs either as the Windows service WebPortServer or as a program (WebPortServer.exe started by a
# user). The add-on runs in that process, so its account decides who can decrypt the client secret (DPAPI).
function Get-WebPortRuntime {
    param([Parameter(Mandatory=$true)][string]$InstallationPath)
    $executable=Join-Path $InstallationPath 'WebPortServer.exe'
    $service=@(Get-CimInstance -ClassName Win32_Service -Filter "Name='WebPortServer'") | Select-Object -First 1
    $process=$null
    if(-not $service -or $service.State -ne 'Running'){
        $process=@(Get-CimInstance -ClassName Win32_Process -Filter "Name='WebPortServer.exe'" | Where-Object { [string]::Equals([string]$_.ExecutablePath,$executable,[StringComparison]::OrdinalIgnoreCase) }) | Select-Object -First 1
    }
    if($process){
        $owner=Invoke-CimMethod -InputObject $process -MethodName GetOwner
        $ownerSid=Invoke-CimMethod -InputObject $process -MethodName GetOwnerSid
        # The arguments are kept, so that the program can be started again the same way.
        $commandLine=[string]$process.CommandLine
        $arguments=if($commandLine.StartsWith('"')){$commandLine.Substring($commandLine.IndexOf('"',1)+1).Trim()}else{($commandLine -replace '^\S+','').Trim()}
        return [pscustomobject]@{Mode='Program';Name='WebPortServer.exe';Running=$true;Account="$($owner.Domain)\$($owner.User)";Sid=[string]$ownerSid.Sid;Arguments=$arguments;PathName=$executable}
    }
    if($service){
        $sid=if($service.StartName -eq 'LocalSystem'){'S-1-5-18'}else{try{([Security.Principal.NTAccount]$service.StartName).Translate([Security.Principal.SecurityIdentifier]).Value}catch{$null}}
        return [pscustomobject]@{Mode='Service';Name=$service.Name;Running=($service.State -eq 'Running');Account=$service.StartName;Sid=$sid;Arguments=$null;PathName=[string]$service.PathName}
    }
    return [pscustomobject]@{Mode='None';Name=$null;Running=$false;Account=$null;Sid=$null;Arguments=$null;PathName=$null}
}

# Starts WebPort again the way it ran before. A program owned by another account cannot be started from
# here; the user is asked to start it, and the function waits up to five minutes for the process.
function Start-WebPortRuntime {
    param([Parameter(Mandatory=$true)]$Runtime,[Parameter(Mandatory=$true)][string]$InstallationPath)
    if($Runtime.Mode -eq 'Service'){
        Start-Service -Name $Runtime.Name
        (Get-Service -Name $Runtime.Name).WaitForStatus('Running',[TimeSpan]::FromSeconds(60))
        ok "Tjänsten $($Runtime.Name) körs"
        return
    }
    if($Runtime.Mode -ne 'Program'){ return }
    $executable=Join-Path $InstallationPath 'WebPortServer.exe'
    if($Runtime.Sid -eq [Security.Principal.WindowsIdentity]::GetCurrent().User.Value){
        $start=@{FilePath=$executable;WorkingDirectory=$InstallationPath}
        if($Runtime.Arguments){$start.ArgumentList=$Runtime.Arguments}
        Start-Process @start
        ok "WebPortServer.exe startad som program ($($Runtime.Account))"
        return
    }
    warn "WebPort kördes som program av $($Runtime.Account). Starta WebPortServer.exe igen som det kontot."
    for($i=0;$i -lt 150;$i++){
        if(@(Get-CimInstance -ClassName Win32_Process -Filter "Name='WebPortServer.exe'" | Where-Object { [string]::Equals([string]$_.ExecutablePath,$executable,[StringComparison]::OrdinalIgnoreCase) }).Count){ ok 'WebPortServer.exe körs igen'; return }
        Start-Sleep -Seconds 2
    }
    throw 'WebPortServer.exe was not started again within five minutes'
}

# The add-on decrypts its client secret as the account WebPort runs as (DPAPI CurrentUser). For the service
# that is LocalSystem: an administrator cannot encrypt for that account directly, so a short-lived scheduled
# task running as SYSTEM re-encrypts a machine-protected transfer copy and writes it to SQLite. For a program
# run by the current user, the same worker runs directly as that user. The worker contains no secret.
$secretWorker=@'
param([string]$StagePath,[string]$DatabasePath,[string]$WebPortProgPath,[guid]$DeviceGuid,[guid]$TenantId,[guid]$ClientId,[string]$ExpectedSid='S-1-5-18')
$ErrorActionPreference='Stop'
#<sqlite-connection>
$state=[ordered]@{Success=$false}
$plain=$null;$roundtrip=$null;$connection=$null;$transaction=$null
$transfer=Join-Path $StagePath 'secret.bin'
try {
    if([Security.Principal.WindowsIdentity]::GetCurrent().User.Value -ne $ExpectedSid){throw 'The secret worker must run as the WebPort service account'}
    Add-Type -AssemblyName System.Security
    $tenant=$TenantId.ToString('D');$client=$ClientId.ToString('D')
    $plain=[Security.Cryptography.ProtectedData]::Unprotect([IO.File]::ReadAllBytes($transfer),[Text.Encoding]::UTF8.GetBytes('WebPort.Entra.SecretTransfer:'+$tenant+':'+$client),[Security.Cryptography.DataProtectionScope]::LocalMachine)
    if($plain.Length -eq 0 -or $plain.Length -gt 16384){throw 'Invalid client secret length'}
    $entropy=[Text.Encoding]::UTF8.GetBytes('WebPort.Entra.SettingsSecret:'+$tenant+':'+$client)
    $protected='dpapi-user-v1:'+[Convert]::ToBase64String([Security.Cryptography.ProtectedData]::Protect($plain,$entropy,[Security.Cryptography.DataProtectionScope]::CurrentUser))
    $roundtrip=[Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String($protected.Substring(14)),$entropy,[Security.Cryptography.DataProtectionScope]::CurrentUser)
    $same=$roundtrip.Length -eq $plain.Length
    for($i=0;$same -and $i -lt $plain.Length;$i++){$same=$roundtrip[$i] -eq $plain[$i]}
    if(-not $same){throw 'Service-account DPAPI roundtrip failed'}
    $connection=New-WebPortIamSqliteConnection $DatabasePath $WebPortProgPath -Mode ReadWrite
    $connection.Open()
    $transaction=$connection.BeginTransaction()
    $command=$connection.CreateCommand()
    try {
        $command.Transaction=$transaction
        $command.CommandText='DELETE FROM Settings WHERE DeviceGuid=@dg COLLATE NOCASE AND "Key"=''EntraClientSecretProtected'''
        [void]$command.Parameters.AddWithValue('@dg',$DeviceGuid.ToString('D'))
        [void]$command.ExecuteNonQuery()
        $command.CommandText='INSERT INTO Settings(DeviceGuid,"Key","Value") VALUES(@dg,''EntraClientSecretProtected'',@value)'
        [void]$command.Parameters.AddWithValue('@value',$protected)
        if($command.ExecuteNonQuery() -ne 1){throw 'Unexpected affected-row count'}
    } finally {$command.Dispose()}
    $transaction.Commit()
    $state.Success=$true
} catch {
    if($null -ne $transaction){try{$transaction.Rollback()}catch{}}
    $state.Error=$_.Exception.Message
} finally {
    if($plain){[Array]::Clear($plain,0,$plain.Length)}
    if($roundtrip){[Array]::Clear($roundtrip,0,$roundtrip.Length)}
    if($null -ne $transaction){$transaction.Dispose()}
    if($null -ne $connection){$connection.Dispose()}
    if(Test-Path -LiteralPath $transfer){Remove-Item -LiteralPath $transfer -Force}
}
$state | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $StagePath 'state.json') -Encoding UTF8
'@

function Save-WebPortEntraClientSecret {
    param([Parameter(Mandatory=$true)][Security.SecureString]$Secret,[guid]$Tenant,[guid]$Client,[string]$AccountSid='S-1-5-18')
    Add-Type -AssemblyName System.Security
    $stage=Join-Path (Join-Path $WebPortDataPath 'Entra\secret-staging') ([guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    $acl=[Security.AccessControl.DirectorySecurity]::new()
    $acl.SetAccessRuleProtection($true,$false)
    $acl.SetOwner([Security.Principal.SecurityIdentifier]::new('S-1-5-32-544'))
    foreach($sid in @('S-1-5-18','S-1-5-32-544')){
        $acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new([Security.Principal.SecurityIdentifier]::new($sid),[Security.AccessControl.FileSystemRights]::FullControl,[Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit',[Security.AccessControl.PropagationFlags]::None,[Security.AccessControl.AccessControlType]::Allow))
    }
    Set-Acl -LiteralPath $stage -AclObject $acl
    $taskName='WebPortEntraSecret-'+[guid]::NewGuid().ToString('N')
    $registered=$false;$plain=$null;$characters=$null;$pointer=[IntPtr]::Zero
    try {
        $pointer=[Runtime.InteropServices.Marshal]::SecureStringToGlobalAllocUnicode($Secret)
        $characters=New-Object char[] $Secret.Length
        for($i=0;$i -lt $characters.Length;$i++){$characters[$i]=[char][Runtime.InteropServices.Marshal]::ReadInt16($pointer,$i*2)}
        $plain=[Text.Encoding]::UTF8.GetBytes($characters)
        $transferEntropy=[Text.Encoding]::UTF8.GetBytes('WebPort.Entra.SecretTransfer:'+$Tenant.ToString('D')+':'+$Client.ToString('D'))
        [IO.File]::WriteAllBytes((Join-Path $stage 'secret.bin'),[Security.Cryptography.ProtectedData]::Protect($plain,$transferEntropy,[Security.Cryptography.DataProtectionScope]::LocalMachine))
        $connectionFunction="function New-WebPortIamSqliteConnection {`r`n$(${function:New-WebPortIamSqliteConnection})`r`n}"
        $worker=Join-Path $stage 'Save-WebPortEntraSecret.ps1'
        [IO.File]::WriteAllText($worker,$secretWorker.Replace('#<sqlite-connection>',$connectionFunction),[Text.UTF8Encoding]::new($true))
        $arguments='-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "{0}" -StagePath "{1}" -DatabasePath "{2}" -WebPortProgPath "{3}" -DeviceGuid {4} -TenantId {5} -ClientId {6}' -f $worker,$stage,$DatabasePath,$WebPortProgPath,$SettingsDeviceGuid.ToString('D'),$Tenant.ToString('D'),$Client.ToString('D')
        if($AccountSid -ne 'S-1-5-18'){
            # WebPort runs as a program of the current user: the worker encrypts for this user directly.
            if($AccountSid -ne [Security.Principal.WindowsIdentity]::GetCurrent().User.Value){throw 'The client secret can only be stored for LocalSystem or for the current user'}
            Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList ($arguments+" -ExpectedSid $AccountSid") -Wait -NoNewWindow
            $state=$null;$statePath=Join-Path $stage 'state.json'
            if(Test-Path -LiteralPath $statePath){$state=Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json}
            if(-not $state){throw 'The secret worker did not report a result'}
            if(-not $state.Success){throw "The secret worker could not store the client secret: $($state.Error)"}
            return
        }
        $action=New-ScheduledTaskAction -Execute (Get-Process -Id $PID).Path -Argument $arguments
        $principal=New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        $taskSettings=New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::FromMinutes(5)) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $taskSettings | Out-Null
        $registered=$true
        Start-ScheduledTask -TaskName $taskName
        $statePath=Join-Path $stage 'state.json';$deadline=[DateTime]::UtcNow.AddMinutes(3);$state=$null
        while(-not $state -and [DateTime]::UtcNow -lt $deadline){
            if(Test-Path -LiteralPath $statePath){try{$state=Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json}catch{}}
            if(-not $state){Start-Sleep -Milliseconds 500}
        }
        if(-not $state){throw 'The SYSTEM task did not finish within 3 minutes'}
        if(-not $state.Success){throw "The SYSTEM task could not store the client secret: $($state.Error)"}
    } finally {
        if($plain){[Array]::Clear($plain,0,$plain.Length)}
        if($characters){[Array]::Clear($characters,0,$characters.Length)}
        if($pointer -ne [IntPtr]::Zero){[Runtime.InteropServices.Marshal]::ZeroFreeGlobalAllocUnicode($pointer)}
        if($registered){
            for($i=0;$i -lt 20 -and (Get-ScheduledTask -TaskName $taskName).State -eq 'Running';$i++){Start-Sleep -Milliseconds 500}
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        }
        # Never leave the transfer copy behind, even if the task did not run.
        Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Tests replace this check with their own function.
if(-not (Get-Command Test-WebPortElevation -CommandType Function -ErrorAction SilentlyContinue)){
    function Test-WebPortElevation { ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) }
}

step 'Läser WebPorts databas'
if ([string]::IsNullOrWhiteSpace($DatabasePath)) { $DatabasePath=Join-Path $WebPortDataPath 'db\webport.sqlite' }
$DatabasePath=(Resolve-Path -LiteralPath $DatabasePath).Path
$SettingsDeviceGuid=Get-WebPortSettingsDeviceGuid $DatabasePath $WebPortProgPath
$databaseSettings=Read-WebPortIamSettings $DatabasePath $WebPortProgPath $SettingsDeviceGuid
$defaults=Resolve-WebPortIamDefaults $databaseSettings
$InternalRedirectUri=[uri]$defaults.RedirectUri
$ServerUrl=$defaults.ServerUrl
ok "Databas: $DatabasePath"
ok "Inställningar under DeviceGuid $($SettingsDeviceGuid.ToString('D'))"
ok "Servernamn (SSLCN): $($databaseSettings['SSLCN'])"
if(-not $PSBoundParameters.ContainsKey('TenantId') -and $defaults.TenantCandidates.Count -eq 1) { $TenantId=[guid]$defaults.TenantCandidates[0] }
if(-not $PSBoundParameters.ContainsKey('Plant')) { $Plant=$defaults.Plant }
$DisplayName=$Plant
$plantSource=if($PSBoundParameters.ContainsKey('Plant')){'-Plant'}elseif(-not [string]::IsNullOrWhiteSpace($defaults.Plant)){'General settings → Plant i WebPort'}else{$null}
if ([string]::IsNullOrWhiteSpace($DisplayName) -and -not $ReadDatabaseOnly -and -not $WhatIfPreference) {
    info 'Anläggningens namn saknas i WebPort (General settings → Plant).'
    $DisplayName=Read-Host 'Ange anläggningens namn (sparas i General settings → Plant och används som appnamn i Entra)'
    $plantSource='angivet nu'
}
if($plantSource -and -not [string]::IsNullOrWhiteSpace($DisplayName)){ ok "Anläggning: $($DisplayName.Trim()) ($plantSource)" }
if (-not [string]::IsNullOrWhiteSpace($DisplayName)) {
    $DisplayName=$DisplayName.Trim()
    if ($DisplayName.Length -gt 200 -or $DisplayName -match '[\x00-\x1f\x7f]') { throw 'Plant must contain 1-200 characters without control characters' }
} elseif (-not $ReadDatabaseOnly -and -not $WhatIfPreference) {
    throw 'Plant saknas. Ange -Plant för obevakad körning.'
}
if ($null -eq $InternalRedirectUri -or -not $InternalRedirectUri.IsAbsoluteUri -or $InternalRedirectUri.Scheme -ne 'https' -or
    $InternalRedirectUri.UserInfo -or $InternalRedirectUri.Query -or $InternalRedirectUri.Fragment -or
    $InternalRedirectUri.OriginalString -ne $InternalRedirectUri.OriginalString.Trim() -or
    $InternalRedirectUri.Host -eq 'kiona-identity-portal.azurewebsites.net' -or
    -not $InternalRedirectUri.AbsolutePath.EndsWith('/access/loginsso', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'RedirectUri must be your HTTPS WebPort /access/loginsso callback without whitespace, query or fragment'
}
# Users reach WebPort through Application Proxy on an msappproxy.net address by default:
# https://<plant>-<tenant>.msappproxy.net/. The SSLCN address is the proxy's internal URL.
# The add-on's callback lives on the external address, because that is where the browser is.
function ConvertTo-WebPortProxyLabel([string]$Name) {
    $decomposed=([string]$Name).ToLowerInvariant().Normalize([Text.NormalizationForm]::FormD)
    $ascii=-join @($decomposed.ToCharArray() | Where-Object { [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark })
    $label=($ascii -replace '[^a-z0-9]+','-').Trim('-')
    if([string]::IsNullOrEmpty($label)){'webport'}else{$label}
}
$proxyLabel=ConvertTo-WebPortProxyLabel $DisplayName
$ExternalUrl=$null
if($SkipApplicationProxy -and (-not [string]::IsNullOrWhiteSpace($ApplicationProxyExternalUrl) -or -not [string]::IsNullOrWhiteSpace($ApplicationProxyConnectorGroup))){
    throw 'SkipApplicationProxy cannot be combined with ApplicationProxyExternalUrl or ApplicationProxyConnectorGroup'
}
if(-not [string]::IsNullOrWhiteSpace($ApplicationProxyExternalUrl)){
    $candidate=$null
    if(-not [uri]::TryCreate($ApplicationProxyExternalUrl.Trim(),[UriKind]::Absolute,[ref]$candidate) -or $candidate.Scheme -ne 'https' -or
        $candidate.UserInfo -or $candidate.Query -or $candidate.Fragment -or $candidate.AbsolutePath -ne '/'){
        throw 'ApplicationProxyExternalUrl must be an HTTPS root address, for example https://webport-contoso.msappproxy.net/'
    }
    $ExternalUrl=$candidate.GetLeftPart([UriPartial]::Authority).ToLowerInvariant()+'/'
}
if($SkipApplicationProxy){
    $externalLabel=$null
    $RedirectUri=$InternalRedirectUri
    $redirectLabel=$InternalRedirectUri.OriginalString
} else {
    $externalLabel=if($ExternalUrl){$ExternalUrl}else{"https://$proxyLabel-<tenant>.msappproxy.net/"}
    $RedirectUri=if($ExternalUrl){[uri]($ExternalUrl+'access/loginsso')}else{$null}
    $redirectLabel=$externalLabel+'access/loginsso'
}
$requestedScopes = @('openid','profile','email')
if (-not $WithoutGraphGroups) { $requestedScopes += 'User.Read' }
if (-not [string]::IsNullOrWhiteSpace($Scopes)) {
    $requestedScopes=@($Scopes -split '\s+' | Where-Object {$_} | Select-Object -Unique)
    foreach($scope in $requestedScopes) {
        if (@('openid','profile','email','User.Read') -cnotcontains $scope) { throw "Unsupported scope: $scope. Use openid profile email and optionally User.Read." }
    }
    foreach($scope in @('openid','profile','email')) { if($requestedScopes -notcontains $scope) { throw "Required login scope missing: $scope" } }
    if($WithoutGraphGroups) { $requestedScopes=@($requestedScopes | Where-Object {$_ -cne 'User.Read'}) }
}
# App-role values match the role mapping built into the add-on. IDs are unique to each new app.
$appRoles=@(foreach($definition in @(
    @{Name='Deny';Permission='NOACCESS'},
    @{Name='View';Permission='VIEW'},
    @{Name='Basic';Permission='BASIC'},
    @{Name='Advanced';Permission='ADVANCED'},
    @{Name='System';Permission='SYSTEM'},
    @{Name='Admin';Permission='ADMIN'}
)) {
    @{
        allowedMemberTypes=@('User')
        description='WebPort global permission '+$definition.Permission
        displayName='WebPort '+$definition.Name
        id=[guid]::NewGuid().ToString('D')
        isEnabled=$true
        value='WebPort.'+$definition.Name
    }
})
$tenantLabel=if($TenantId -eq [guid]::Empty){'<tenant från inloggningen>'}else{$TenantId.ToString('D')}
$plannedSettings=[ordered]@{
    'Plant'=$DisplayName
    'WebPortServer Url'=$ServerUrl
    'KIClientId'='<klient-ID från Entra>'
    'KIExternalTenantId'=$tenantLabel
    'KIProviderUri'="https://login.microsoftonline.com/$tenantLabel"
    'KICustomPolicy'=''
    'KISignInPolicyName'=''
    'EntraRedirectUri'=$(if($SkipApplicationProxy){''}else{"$redirectLabel (tom om ingen connector finns)"})
    'AuthenticationTypeKI'='1'
}
Write-Host ''
step 'Planerade inställningar i Entra ID'
setting 'Tenant' $tenantLabel
setting 'Åtgärd' $(if($Force){'Skapa ny app, eller uppdatera befintlig app med samma namn'}else{'Skapa ny app, eller uppdatera namnträff vars klient-ID matchar SQLite'})
setting 'Appnamn' (Format-SettingValue $DisplayName)
setting 'Kontotyp' 'Endast denna organisation (single tenant, AzureADMyOrg)'
setting 'Plattform' 'Web'
setting 'Redirect URI' $redirectLabel
if(-not $SkipApplicationProxy){
    detail "och $externalLabel (Application Proxys förautentisering)"
    detail "och $($InternalRedirectUri.OriginalString) (inloggning direkt på det lokala nätet)"
    detail "(utan connector i Entra: bara den interna)"
}
setting 'Startsida' "$(if($SkipApplicationProxy){$ServerUrl}else{$externalLabel})entra/login"
detail '(appens ruta i My Apps loggar in direkt; en manuellt ändrad startsida behålls)'
setting 'Implicit grant' 'Av (varken access- eller ID-token)'
setting 'Public client flows' 'Av'
setting 'API-behörigheter' "Microsoft Graph, delegerade: $($requestedScopes -join ', ')"
setting 'Admin consent' "Hela organisationen (AllPrincipals) för $($requestedScopes -join ', ')"
setting 'Valfria claims' 'email i ID-token'
setting 'Approller' "$($appRoles.Count) st, tilldelningsbara till användare och grupper:"
foreach($role in $appRoles){ Write-Host ("   {0,-24}   {1,-17} → {2}" -f '',$role.value,($role.description -replace '^WebPort global permission ','')) }
setting 'Enterprise app' 'Tilldelning krävs = Ja (appRoleAssignmentRequired)'
if($SkipApplicationProxy){
    setting 'Entra Application Proxy' 'Hoppas över (-SkipApplicationProxy); befintlig publicering ändras inte'
} else {
    setting 'Entra Application Proxy' "Aktiveras via $(if([string]::IsNullOrWhiteSpace($ApplicationProxyConnectorGroup)){'befintlig aktiv connector group (automatiskt val); hoppas över om ingen connector finns'}else{"connector group '$ApplicationProxyConnectorGroup'"})"
    setting 'Proxy extern URL' "$externalLabel$(if(-not $ExternalUrl){' (tenantnamnet hämtas vid inloggningen)'})"
    setting 'Proxy intern URL' $ServerUrl
    setting 'Proxy Host-huvud' 'Översätts inte; WebPort ser den externa adressen, som tillägget kontrollerar'
    setting 'Application ID URI' 'api://<klient-ID>, om appen saknar ett (Entra kräver det för Application Proxy)'
}
if($SkipClientSecret){ setting 'Klienthemlighet' 'Skapas inte (-SkipClientSecret)' }
else {
    setting 'Klienthemlighet' "Ny hemlighet 'WebPort Entra integration', giltig $SecretLifetimeMonths månader (till $([DateTime]::Now.AddMonths($SecretLifetimeMonths).ToString('yyyy-MM-dd')))"
    warn "Hemligheten måste bytas före $([DateTime]::Now.AddMonths($SecretLifetimeMonths).ToString('yyyy-MM-dd')), annars slutar Entra-inloggningen att fungera"
}
if($Force){
    warn 'Om appen redan finns: approller, API-behörigheter, claims och den externa redirect URI:n kompletteras.'
    warn 'Befintliga roll-ID, redirect URI:er, startsida och gamla hemligheter behålls. Andra roller och behörigheter lämnas orörda.'
}
Write-Host ''
step 'Planerade ändringar på servern'
if($UpdateDatabase){
    setting 'WebPort' 'Stoppas efter lyckad registrering (tjänst eller program). Startas igen om det körde och tillägget redan är installerat'
    setting 'Databasbackup' "$DatabasePath.EntraIAM-backup-<tidpunkt>.sqlite"
    if(-not $SkipClientSecret){
        setting 'Klienthemlighet' 'Krypteras med DPAPI för kontot som kör WebPort och sparas i SQLite som EntraClientSecretProtected:'
        Write-Host ("   {0,-24}   {1}" -f '','tjänst: LocalSystem, via ett tillfälligt schemalagt jobb som SYSTEM')
        Write-Host ("   {0,-24}   {1}" -f '','program: ditt konto (scriptet måste köras av samma konto som WebPort)')
    }
    setting 'SQLite-inställningar' "DeviceGuid $($SettingsDeviceGuid.ToString('D')):"
    foreach($key in $plannedSettings.Keys){
        $current=if($databaseSettings.ContainsKey($key)){$databaseSettings[$key]}else{$null}
        $change=if($null -ne $current -and [string]$current -ceq [string]$plannedSettings[$key]){"$(Format-SettingValue $current) (oförändrat)"}else{"$(Format-SettingValue $current) → $(Format-SettingValue $plannedSettings[$key])"}
        Write-Host ("   {0,-24}   {1,-19} {2}" -f '',$key,$change)
    }
} else {
    setting 'WebPort och SQLite' 'Ändras inte (-UpdateDatabase:$false eller -ReadDatabaseOnly)'
}
Write-Host ''
if($ReadDatabaseOnly) {
    ok 'Endast läsning: inga ändringar har gjorts.'
    [pscustomobject]@{DatabasePath=$DatabasePath;SettingsDeviceGuid=$SettingsDeviceGuid.ToString('D');TenantId=$(if($TenantId -ne [guid]::Empty){$TenantId.ToString('D')}else{$null});TenantFromSignIn=($TenantId -eq [guid]::Empty);
        Plant=$DisplayName;DisplayName=$DisplayName;PlantRequired=[string]::IsNullOrWhiteSpace($DisplayName);ServerUrl=$ServerUrl;ApplicationProxyExternalUrl=$externalLabel;RedirectUri=$redirectLabel;Operation=$(if($Force){'CreateOrUpdateByDisplayName'}else{'CreateOrUpdateMatchingDatabaseClientId'});
        Authority=$(if($TenantId -ne [guid]::Empty){"https://login.microsoftonline.com/$($TenantId.ToString('D'))/v2.0"}else{$null});Scopes=($requestedScopes -join ' ');AppRoles=$appRoles;ReadOnly=$true}
    return
}
if (-not $PSCmdlet.ShouldProcess("$tenantLabel / $(Format-SettingValue $DisplayName)", 'Utför de planerade inställningarna ovan')) { return }

# Local prerequisites are checked before anything changes in Entra.
$webPortRuntime=$null
if($UpdateDatabase){
    step 'Kontrollerar förutsättningar på servern'
    if(-not (Test-WebPortElevation)){
        err 'Kör scriptet som administratör. Det stoppar WebPort och sparar klienthemligheten för kontot som kör WebPort.'
        throw 'Run as administrator: UpdateDatabase stops WebPort and stores the client secret for the account WebPort runs as'
    }
    $webPortRuntime=Get-WebPortRuntime $WebPortProgPath
    $currentSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    if(-not $SkipClientSecret){
        # The secret can be encrypted for LocalSystem (service) or for the current user (program run by this user).
        if($webPortRuntime.Mode -eq 'Service' -and $webPortRuntime.Sid -ne 'S-1-5-18'){
            err "Tjänsten WebPortServer kör som $($webPortRuntime.Account). Klienthemligheten kan bara sparas för LocalSystem eller för programmets eget konto."
            throw 'The WebPortServer service must run as LocalSystem; the client secret is encrypted for that account'
        }
        if($webPortRuntime.Mode -eq 'Program' -and $webPortRuntime.Sid -ne $currentSid){
            err "WebPort körs som program av $($webPortRuntime.Account). Kör scriptet som det kontot, eftersom klienthemligheten krypteras för kontot som kör WebPort."
            throw 'WebPort runs as a program of another account; run the script as that account'
        }
        if($webPortRuntime.Mode -eq 'None'){
            err 'WebPort körs inte, och tjänsten WebPortServer finns inte. Starta WebPort (som tjänst eller program), så att scriptet vet vilket konto hemligheten ska krypteras för.'
            throw 'WebPort is neither installed as a service nor running as a program'
        }
    }
    ok "Administratör; WebPort $(switch($webPortRuntime.Mode){'Service'{"körs som tjänst ($($webPortRuntime.Account), $(if($webPortRuntime.Running){'körs'}else{'stoppad'}))"}'Program'{"körs som program av $($webPortRuntime.Account)"}default{'körs inte'}})"
}

step 'Ansluter till Microsoft Graph'
# Directory.ReadWrite.All is needed only for Application Proxy and the tenant's msappproxy.net name.
$graphScopes=@('Application.ReadWrite.All','DelegatedPermissionGrant.ReadWrite.All')
if(-not $SkipApplicationProxy){$graphScopes+='Directory.ReadWrite.All'}
if (-not $UseExistingConnection) {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'Install the prerequisite: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    $connect = @{
        Scopes = $graphScopes
        ContextScope = 'Process'
        NoWelcome = $true
    }
    if ($TenantId -ne [guid]::Empty) { $connect.TenantId=$TenantId.ToString('D') }
    if ($UseDeviceCode) { $connect.UseDeviceCode = $true }
    Connect-MgGraph @connect | Out-Null
}
$context = Get-MgContext
if ($null -eq $context -or $context.Environment -ne 'Global') {
    throw 'Microsoft Graph context must target the requested tenant in the Global environment'
}
$signedInTenant=[guid]::Empty
if (-not [guid]::TryParse([string]$context.TenantId,[ref]$signedInTenant) -or $signedInTenant -eq [guid]::Empty) { throw 'Sign-in did not resolve a specific Entra tenant' }
if ($TenantId -eq [guid]::Empty) {
    $TenantId=$signedInTenant
} elseif ($signedInTenant -ne $TenantId) { throw 'Microsoft Graph context targets a different tenant than requested' }
if ($context.AuthType -ne 'Delegated') { throw 'Use an interactive delegated Microsoft Graph connection for this script' }
foreach ($permission in $graphScopes) {
    if ($context.Scopes -notcontains $permission) { throw "Microsoft Graph connection is missing $permission" }
}
ok "Inloggad$(if($context.Account){" som $($context.Account)"}) i tenant $($TenantId.ToString('D'))"

function Invoke-WebPortGraph {
    param([string]$Method='GET',[Parameter(Mandatory=$true)][string]$Path,$Body,[ValidateSet('v1.0','beta')][string]$ApiVersion='v1.0')
    $uri = if ($Path.StartsWith('https://graph.microsoft.com/',[StringComparison]::Ordinal)) { $Path } else { "https://graph.microsoft.com/$ApiVersion/" + $Path.TrimStart('/') }
    if (-not ($uri.StartsWith('https://graph.microsoft.com/v1.0/',[StringComparison]::Ordinal) -or $uri.StartsWith('https://graph.microsoft.com/beta/',[StringComparison]::Ordinal))) { throw 'Unexpected Microsoft Graph URL' }
    $arguments = @{ Method=$Method; Uri=$uri; OutputType='PSObject'; ErrorAction='Stop' }
    if ($null -ne $Body) { $arguments.Body = ($Body | ConvertTo-Json -Depth 20 -Compress); $arguments.ContentType='application/json' }
    for ($attempt=0; $attempt -lt 5; $attempt++) {
        try { return Invoke-MgGraphRequest @arguments }
        catch {
            $code = 0
            if ($_.Exception.ResponseStatusCode) { $code = [int]$_.Exception.ResponseStatusCode }
            elseif ($_.Exception.Response -and $_.Exception.Response.StatusCode) { $code = [int]$_.Exception.Response.StatusCode }
            # Directory replication may temporarily hide a newly created application or service principal.
            # Never automatically repeat a non-idempotent POST: its first result may have been committed.
            if ($Method -notin @('GET','PATCH','PUT') -or $code -notin @(404,429,503) -or $attempt -eq 4) { throw }
            Start-Sleep -Seconds ([Math]::Min(8,[Math]::Pow(2,$attempt)))
        }
    }
}
function Get-WebPortGraphCollection {
    param([string]$Path,[ValidateSet('v1.0','beta')][string]$ApiVersion='v1.0')
    do {
        $response = Invoke-WebPortGraph -Path $Path -ApiVersion $ApiVersion
        foreach ($item in @($response.value)) { if ($null -ne $item) { $item } }
        $Path = $response.'@odata.nextLink'
        if ($Path -and -not ($Path.StartsWith('https://graph.microsoft.com/v1.0/',[StringComparison]::Ordinal) -or $Path.StartsWith('https://graph.microsoft.com/beta/',[StringComparison]::Ordinal))) { throw 'Unexpected Graph pagination URL' }
    } while ($Path)
}
function Stop-WebPortRegistrationValidation {
    param([string]$Reason)
    $script:WebPortRegistrationFailureReason=$Reason
    throw $Reason
}
function Assert-WebPortRegistration {
    param($App,$Grant,[string]$PrincipalId,[string]$GraphPrincipalId)
    if (-not [string]::Equals($App.displayName,$DisplayName,[StringComparison]::OrdinalIgnoreCase) -or $App.signInAudience -ne 'AzureADMyOrg' -or $App.isFallbackPublicClient -eq $true -or @($App.web.redirectUris) -notcontains $RedirectUri.OriginalString -or
        ($proxyGroup -and @($App.web.redirectUris) -notcontains $ExternalUrl)) {
        throw 'Application verification failed: tenant restriction, client type or redirect URI mismatch'
    }
    if ($App.web.implicitGrantSettings.enableAccessTokenIssuance -or $App.web.implicitGrantSettings.enableIdTokenIssuance) { throw 'Implicit flow unexpectedly enabled' }
    $configuredGraph = @($App.requiredResourceAccess | Where-Object { $_.resourceAppId -eq '00000003-0000-0000-c000-000000000000' })
    foreach ($access in $resourceAccess) {
        if (-not @($configuredGraph.resourceAccess | Where-Object { $_.id -eq $access.id -and $_.type -eq $access.type }).Count) { throw 'Application API permission readback mismatch' }
    }
    if (-not @($App.optionalClaims.idToken | Where-Object { $_.name -eq 'email' }).Count) { throw 'Optional email claim not configured' }
    if (@($App.appRoles).Count -ne $manifestRoles.Count) { throw 'Application app role readback mismatch' }
    foreach($expectedRole in $manifestRoles) {
        $matches=@($App.appRoles | Where-Object { $_.id -eq $expectedRole.id })
        if($matches.Count -ne 1) { throw 'Application app role ID readback mismatch' }
        $role=$matches[0]
        if($role.value -cne $expectedRole.value -or $role.displayName -cne $expectedRole.displayName -or [string]$role.description -cne [string]$expectedRole.description -or
            $role.isEnabled -ne $expectedRole.isEnabled -or ((@($role.allowedMemberTypes | Sort-Object) -join ',') -cne (@($expectedRole.allowedMemberTypes | Sort-Object) -join ','))) {
            throw 'Application app role definition readback mismatch'
        }
    }
    if ($Grant.clientId -ne $PrincipalId -or $Grant.resourceId -ne $GraphPrincipalId -or $Grant.consentType -ne 'AllPrincipals') {
        throw 'Tenant-wide delegated consent verification failed'
    }
    $consented = @($Grant.scope -split ' ' | Where-Object { $_ })
    foreach ($scope in $requestedScopes) { if ($consented -notcontains $scope) { throw "Consent verification failed for $scope" } }
}

function Select-WebPortApplicationProxyConnectorGroup {
    $groups=@(Get-WebPortGraphCollection -ApiVersion beta -Path 'onPremisesPublishingProfiles/applicationProxy/connectorGroups?$select=id,name,isDefault,connectorGroupType')
    $candidates=@();$connectorCount=0
    foreach($group in $groups){
        if(-not $group.id -or ($group.connectorGroupType -and $group.connectorGroupType -ne 'applicationProxy')){continue}
        $members=@(Get-WebPortGraphCollection -ApiVersion beta -Path "onPremisesPublishingProfiles/applicationProxy/connectorGroups/$($group.id)/members")
        $connectorCount+=$members.Count
        $active=@($members | Where-Object {$_.status -eq 'active'})
        if($active.Count){$candidates+=[pscustomobject]@{Group=$group;ActiveConnectors=$active}}
    }
    # No connector at all means Application Proxy is not set up in the tenant; the caller then registers without it.
    # Connectors that exist but are all inactive are an error, so a temporary outage never moves the callback.
    if($connectorCount -eq 0 -and [string]::IsNullOrWhiteSpace($ApplicationProxyConnectorGroup)){return $null}
    if(-not [string]::IsNullOrWhiteSpace($ApplicationProxyConnectorGroup)){
        $selected=@($candidates | Where-Object {$_.Group.id -eq $ApplicationProxyConnectorGroup -or [string]::Equals($_.Group.name,$ApplicationProxyConnectorGroup,[StringComparison]::OrdinalIgnoreCase)})
        if($selected.Count -ne 1){Stop-WebPortRegistrationValidation "Application Proxy connector group '$ApplicationProxyConnectorGroup' was not found uniquely with an active connector"}
        return $selected[0]
    }
    $local=@($candidates | Where-Object {@($_.ActiveConnectors | Where-Object {[string]::Equals($_.machineName,$env:COMPUTERNAME,[StringComparison]::OrdinalIgnoreCase)}).Count})
    if($local.Count -eq 1){return $local[0]}
    $default=@($candidates | Where-Object {$_.Group.isDefault -eq $true})
    if($default.Count -eq 1){return $default[0]}
    if($candidates.Count -eq 1){return $candidates[0]}
    if($candidates.Count -eq 0){Stop-WebPortRegistrationValidation 'Entra Application Proxy connectors exist, but none is active. Start a connector, or use -SkipApplicationProxy'}
    Stop-WebPortRegistrationValidation 'Multiple active Entra Application Proxy connector groups exist. Specify -ApplicationProxyConnectorGroup with the group name or ID'
}

function Merge-WebPortAppRoles {
    param($ExistingRoles)
    $merged=[Collections.Generic.List[object]]::new()
    $ids=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $found=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach($role in @($ExistingRoles)) {
        if($null -eq $role){continue}
        $id=[guid]::Empty
        if(-not [guid]::TryParse([string]$role.id,[ref]$id) -or $id -eq [guid]::Empty -or -not $ids.Add($id.ToString('D'))){Stop-WebPortRegistrationValidation 'Existing app has invalid or duplicate role IDs'}
        $matches=@($appRoles | Where-Object {$_.value -ceq $role.value})
        if($matches.Count -eq 1) {
            if(-not $found.Add([string]$role.value)){Stop-WebPortRegistrationValidation 'Existing app has duplicate WebPort role values'}
            if(@($role.allowedMemberTypes).Count -ne 1 -or $role.allowedMemberTypes[0] -cne 'User'){Stop-WebPortRegistrationValidation 'Existing WebPort role has conflicting member types; review its application permission assignments first'}
            $matches[0].id=$id.ToString('D')
            $merged.Add($matches[0])
        }else{
            if(@($appRoles | Where-Object {$_.value -ieq $role.value}).Count){Stop-WebPortRegistrationValidation 'Existing WebPort role differs only in letter case; review it before updating'}
            # origin is read-only and must never be sent back in PATCH.
            $merged.Add(@{id=$id.ToString('D');value=$role.value;displayName=$role.displayName;description=$role.description;isEnabled=$role.isEnabled;allowedMemberTypes=@($role.allowedMemberTypes)})
        }
    }
    foreach($role in $appRoles){if(-not $found.Contains([string]$role.value)){$merged.Add($role)}}
    return $merged.ToArray()
}

function Merge-WebPortPermissions {
    param($Existing)
    $merged=[Collections.Generic.List[object]]::new();$graphFound=$false
    foreach($resource in @($Existing)) {
        if($null -eq $resource){continue}
        $access=@($resource.resourceAccess)
        if($resource.resourceAppId -eq $graphAppId){
            if($graphFound){throw 'Existing app has duplicate Graph permission entries'}
            $graphFound=$true
            foreach($required in $resourceAccess){if(-not @($access | Where-Object {$_.id -eq $required.id -and $_.type -eq $required.type}).Count){$access+=$required}}
        }
        $merged.Add(@{resourceAppId=$resource.resourceAppId;resourceAccess=$access})
    }
    if(-not $graphFound){$merged.Add(@{resourceAppId=$graphAppId;resourceAccess=$resourceAccess})}
    return $merged.ToArray()
}

$application = $null
$newApplication = $false
$existingApplication=$null;$principal=$null;$grant=$null
$script:WebPortRegistrationFailureReason=$null
$stage = 'read Microsoft Graph permissions'
step 'Läser Microsoft Graph-behörigheter'
try {
    $graphAppId = '00000003-0000-0000-c000-000000000000'
    $graphFilter = [uri]::EscapeDataString("appId eq '$graphAppId'")
    $graphPrincipals = @(Get-WebPortGraphCollection -Path ('servicePrincipals?$filter=' + $graphFilter + '&$select=id,appId,oauth2PermissionScopes'))
    if ($graphPrincipals.Count -ne 1) { throw 'Cannot uniquely resolve the Microsoft Graph service principal' }
    $graph = $graphPrincipals[0]
    $resourceAccess = @()
    foreach ($scope in $requestedScopes) {
        $matches = @($graph.oauth2PermissionScopes | Where-Object { $_.value -ceq $scope -and $_.isEnabled })
        if ($matches.Count -ne 1) { throw "Graph does not expose the enabled delegated permission $scope" }
        $resourceAccess += @{ id=[string]$matches[0].id; type='Scope' }
    }
    ok "Microsoft Graph erbjuder $($requestedScopes -join ', ')"
    $proxyGroup=$null
    if(-not $SkipApplicationProxy){
        $stage='select existing Application Proxy connector group'
        step 'Söker befintlig Entra Application Proxy-connector'
        $proxySelection=Select-WebPortApplicationProxyConnectorGroup
        if($proxySelection){
            $proxyGroup=$proxySelection.Group
            ok "Connector group: $($proxyGroup.name) ($($proxyGroup.id)); aktiva connectors: $(@($proxySelection.ActiveConnectors).Count)"
        } else {
            if($ExternalUrl){Stop-WebPortRegistrationValidation 'ApplicationProxyExternalUrl was given, but no Entra Application Proxy connector is configured'}
            info 'Ingen Application Proxy-connector är konfigurerad i Entra. Appen publiceras inte via Application Proxy.'
            info "Inloggning sker på den interna adressen $ServerUrl"
            $RedirectUri=$InternalRedirectUri
        }
    }
    if($proxyGroup -and -not $ExternalUrl){
        $stage='determine Application Proxy external URL'
        step 'Bestämmer extern adress för Application Proxy'
        # msappproxy.net addresses end with the tenant's initial onmicrosoft.com name.
        $initialDomains=@(Get-WebPortGraphCollection -Path 'organization?$select=verifiedDomains' | ForEach-Object { @($_.verifiedDomains) } | Where-Object { $_.isInitial -and ([string]$_.name).EndsWith('.onmicrosoft.com',[StringComparison]::OrdinalIgnoreCase) })
        if($initialDomains.Count -ne 1){Stop-WebPortRegistrationValidation 'Cannot determine the tenant''s initial onmicrosoft.com domain. Specify -ApplicationProxyExternalUrl'}
        $tenantName=([string]$initialDomains[0].name).Split('.')[0].ToLowerInvariant()
        $maxLabel=63-1-$tenantName.Length
        $label=if($proxyLabel.Length -gt $maxLabel){$proxyLabel.Substring(0,$maxLabel).TrimEnd('-')}else{$proxyLabel}
        $ExternalUrl="https://$label-$tenantName.msappproxy.net/"
        $RedirectUri=[uri]($ExternalUrl+'access/loginsso')
        ok "Extern adress: $ExternalUrl (tenant $($initialDomains[0].name))"
    }
    step "Söker efter app med namnet '$DisplayName'"
    $stage='find application by facility name'
    $nameFilter=[uri]::EscapeDataString("displayName eq '"+$DisplayName.Replace("'","''")+"'")
    $matchingApps=@(Get-WebPortGraphCollection -Path ('applications?$filter='+$nameFilter+'&$select=id,appId,displayName') | Where-Object {[string]::Equals($_.displayName,$DisplayName,[StringComparison]::OrdinalIgnoreCase)})
    $databaseClientId=[guid]::Empty
    $hasDatabaseClientId=[guid]::TryParse([string]$databaseSettings.KIClientId,[ref]$databaseClientId) -and $databaseClientId -ne [guid]::Empty
    $databaseClientMatches=@()
    if($hasDatabaseClientId){
        $databaseClientMatches=@($matchingApps | Where-Object {
            $candidate=[guid]::Empty
            [guid]::TryParse([string]$_.appId,[ref]$candidate) -and $candidate -eq $databaseClientId
        })
    }
    $matchedByDatabaseClientId=$databaseClientMatches.Count -eq 1
    if($matchedByDatabaseClientId){$matchingApps=@($databaseClientMatches[0])}
    if($matchingApps.Count -gt 1){
        $instanceClientLabel=if($hasDatabaseClientId){$databaseClientId.ToString('D')}else{'not configured (KIClientId missing or invalid)'}
        Stop-WebPortRegistrationValidation "Multiple applications have the same facility name. DisplayName: '$DisplayName'. This WebPort instance's ClientId (SQLite KIClientId): $instanceClientLabel. No application can be selected safely, even with Force"
    }
    if($matchingApps.Count -eq 1){
        if(-not $Force -and -not $matchedByDatabaseClientId){Stop-WebPortRegistrationValidation 'An application with this facility name already exists, but its client ID does not match SQLite KIClientId. Use -Force to update it.'}
        $existingApplication=Invoke-WebPortGraph -Path ("applications/$($matchingApps[0].id)"+'?$select=id,appId,displayName,signInAudience,isFallbackPublicClient,web,passwordCredentials,requiredResourceAccess,optionalClaims,appRoles')
        $application=$existingApplication
        if($existingApplication.id -ne $matchingApps[0].id -or -not $existingApplication.appId){Stop-WebPortRegistrationValidation 'Existing application identity readback mismatch'}
        # A WebPort app has this run's callback, the internal SSLCN callback (registered without Application Proxy)
        # or an msappproxy.net callback (registered with it). Any of them makes the app compatible.
        $existingRedirects=@($existingApplication.web.redirectUris)
        $knownCallback=@($existingRedirects | Where-Object {
            $candidate=$null
            $_ -eq $RedirectUri.OriginalString -or $_ -eq $InternalRedirectUri.OriginalString -or
            ([uri]::TryCreate([string]$_,[UriKind]::Absolute,[ref]$candidate) -and $candidate.Scheme -eq 'https' -and $candidate.Host.EndsWith('.msappproxy.net',[StringComparison]::OrdinalIgnoreCase) -and $candidate.AbsolutePath -eq '/access/loginsso')
        })
        if($existingApplication.signInAudience -ne 'AzureADMyOrg' -or $existingApplication.isFallbackPublicClient -eq $true -or $knownCallback.Count -eq 0 -or
            $existingApplication.web.implicitGrantSettings.enableAccessTokenIssuance -or $existingApplication.web.implicitGrantSettings.enableIdTokenIssuance){Stop-WebPortRegistrationValidation 'Existing app is not a compatible single-tenant WebPort Web app with the expected callback'}
        $spFilter=[uri]::EscapeDataString("appId eq '$($existingApplication.appId)'")
        $principals=@(Get-WebPortGraphCollection -Path ('servicePrincipals?$filter='+$spFilter+'&$select=id,appId,displayName'))
        if($principals.Count -gt 1){Stop-WebPortRegistrationValidation 'Multiple enterprise applications match the client ID'}
        if($principals.Count -eq 1){
            $principal=$principals[0]
            if(-not $principal.id -or $principal.appId -ne $existingApplication.appId){Stop-WebPortRegistrationValidation 'Enterprise application identity mismatch'}
            $grantFilter=[uri]::EscapeDataString("clientId eq '$($principal.id)' and resourceId eq '$($graph.id)' and consentType eq 'AllPrincipals'")
            $grants=@(Get-WebPortGraphCollection -Path ('oauth2PermissionGrants?$filter='+$grantFilter) | Where-Object {$_.clientId -eq $principal.id -and $_.resourceId -eq $graph.id -and $_.consentType -eq 'AllPrincipals'})
            if($grants.Count -gt 1){Stop-WebPortRegistrationValidation 'Multiple tenant-wide Graph consent grants match the enterprise application'}
            if($grants.Count -eq 1){$grant=$grants[0]}
        }
        ok "Befintlig app hittad: klient-ID $($existingApplication.appId)$(if($matchedByDatabaseClientId){' (matchar SQLite KIClientId)'})"
        ok "Enterprise app: $(if($principal){'finns'}else{'saknas, skapas'}); admin consent: $(if($grant){"finns ($($grant.scope))"}else{'saknas, skapas'})"
    } else {
        ok 'Ingen app med det namnet finns'
    }
    $permissions = @()
    # Application Proxy pre-authentication returns to the external root, so it is registered next to the add-on callback.
    # The internal SSLCN callback is registered too, so that sign-in also works directly on the local network.
    $redirects = @($RedirectUri.OriginalString)
    if($proxyGroup){ $redirects += @($ExternalUrl,$InternalRedirectUri.OriginalString) }
    # The My Apps tile opens the home page; /entra/login signs the user straight in through SSO.
    $homePage=$(if($ExternalUrl){$ExternalUrl}else{$ServerUrl})+'entra/login'
    $optionalIdClaims = @(@{ name='email'; source=$null; essential=$false; additionalProperties=@() })
    $optionalClaims = @{ idToken=$optionalIdClaims }
    $permissions += @{ resourceAppId=$graphAppId; resourceAccess=$resourceAccess }
    $manifestRoles=$appRoles
    if($existingApplication){
        $manifestRoles=@(Merge-WebPortAppRoles $existingApplication.appRoles)
        $permissions=@(Merge-WebPortPermissions $existingApplication.requiredResourceAccess)
        $optionalClaims=@{}
        foreach($kind in @('idToken','accessToken','saml2Token')){if($existingApplication.optionalClaims.$kind){$optionalClaims[$kind]=@($existingApplication.optionalClaims.$kind)}}
        if(-not @($optionalClaims.idToken | Where-Object {$_.name -ceq 'email'}).Count){$optionalClaims.idToken=@($optionalClaims.idToken | Where-Object {$null -ne $_})+$optionalIdClaims}
    }
    $body = @{
        displayName=$DisplayName
        signInAudience='AzureADMyOrg'; isFallbackPublicClient=$false
        web=@{ homePageUrl=$homePage; redirectUris=$redirects; implicitGrantSettings=@{ enableAccessTokenIssuance=$false; enableIdTokenIssuance=$false } }
        requiredResourceAccess=$permissions; optionalClaims=$optionalClaims; appRoles=$manifestRoles
    }
    if($existingApplication){
        $stage='update existing application'
        step "Uppdaterar befintlig app i Entra: $DisplayName"
        $update=@{appRoles=$manifestRoles;requiredResourceAccess=$permissions;optionalClaims=$optionalClaims}
        $missingRedirects=@($redirects | Where-Object { $existingRedirects -notcontains $_ })
        # Only a home page this script set earlier (or none) is replaced; a manually chosen one stays.
        $existingHomePage=[string]$existingApplication.web.homePageUrl
        $scriptHomePages=@($ServerUrl,$ExternalUrl) | Where-Object {$_} | ForEach-Object { $_; $_.TrimEnd('/') }
        $replaceHomePage=$existingHomePage -ne $homePage -and ([string]::IsNullOrWhiteSpace($existingHomePage) -or $scriptHomePages -contains $existingHomePage)
        if($missingRedirects.Count -or $replaceHomePage){
            # The whole web object is sent so that the other web settings stay as they are.
            $update.web=@{ homePageUrl=$(if($replaceHomePage){$homePage}else{$existingApplication.web.homePageUrl}); redirectUris=@($existingRedirects+$missingRedirects); implicitGrantSettings=@{ enableAccessTokenIssuance=$false; enableIdTokenIssuance=$false } }
            if($existingApplication.web.logoutUrl){$update.web.logoutUrl=$existingApplication.web.logoutUrl}
        }
        Invoke-WebPortGraph -Method PATCH -Path "applications/$($application.id)" -Body $update | Out-Null
        ok "Approller: $(@($manifestRoles).Count) st (befintliga roll-ID behålls)"
        ok 'API-behörigheter och valfria claims kompletterade'
        foreach($redirect in $redirects){ ok "Redirect URI: $redirect $(if($missingRedirects -contains $redirect){'tillagd; befintliga behålls'}else{'(fanns redan)'})" }
        if($replaceHomePage){ ok "Startsida: $homePage (ersatte $(if($existingHomePage){$existingHomePage}else{'tom'}))" }
        elseif($existingHomePage -eq $homePage){ ok "Startsida: $homePage (fanns redan)" }
        else{ ok "Startsida: $existingHomePage behålls (ändrad manuellt)" }
    }else{
        $stage = 'create new application'
        step "Registrerar en ny app i Entra: $DisplayName"
        $application = Invoke-WebPortGraph -Method POST -Path 'applications' -Body $body
        $newApplication=$true
        ok 'Single tenant, plattform Web, redirect URI, startsida, implicit grant av'
        ok "API-behörigheter: $($requestedScopes -join ', '); valfritt claim: email"
        ok "Approller: $(($manifestRoles | ForEach-Object { $_.value }) -join ', ')"
    }
    if (-not $application.id -or -not $application.appId) { throw 'Graph did not return application object ID and client ID' }
    ok "Klient-ID (Application ID): $($application.appId)"
    ok "Objekt-ID: $($application.id)"
    $stage = 'create enterprise application'
    step 'Konfigurerar enterprise app'
    Invoke-WebPortGraph -Path "applications/$($application.id)" | Out-Null
    if(-not $principal){
        $principal = Invoke-WebPortGraph -Method POST -Path 'servicePrincipals' -Body @{ appId=$application.appId; displayName=$DisplayName; appRoleAssignmentRequired=$true }
    }else{
        $stage='require enterprise application assignment'
        Invoke-WebPortGraph -Method PATCH -Path "servicePrincipals/$($principal.id)" -Body @{appRoleAssignmentRequired=$true} | Out-Null
    }
    if (-not $principal.id) { throw 'Missing enterprise application object ID' }
    $stage='verify enterprise application assignment requirement'
    $verifiedPrincipal=Invoke-WebPortGraph -Path "servicePrincipals/$($principal.id)"
    if($verifiedPrincipal.id -ne $principal.id -or $verifiedPrincipal.appId -ne $application.appId -or $verifiedPrincipal.appRoleAssignmentRequired -ne $true){
        Stop-WebPortRegistrationValidation 'Enterprise application verification failed: Assignment required must be Yes (appRoleAssignmentRequired=true)'
    }
    ok "Enterprise app $($principal.id): Tilldelning krävs = Ja (verifierat)"
    if($proxyGroup){
        $stage='enable Entra Application Proxy'
        step 'Aktiverar Entra Application Proxy'
        # Entra accepts onPremisesPublishing only on an app that has an application ID URI. Microsoft's guide
        # sets api://<appId>; an existing value is kept.
        $identity=Invoke-WebPortGraph -Path ("applications/$($application.id)"+'?$select=id,appId,identifierUris')
        if(@($identity.identifierUris | Where-Object {$_}).Count -eq 0){
            Invoke-WebPortGraph -Method PATCH -Path "applications/$($application.id)" -Body @{identifierUris=@("api://$($application.appId)")} | Out-Null
            ok "Application ID URI: api://$($application.appId) (krävs av Application Proxy)"
        } else {
            ok "Application ID URI: $(@($identity.identifierUris) -join ', ') (fanns redan)"
        }
        $proxySettings=@{
            externalAuthenticationType='aadPreAuthentication'
            internalUrl=$ServerUrl
            externalUrl=$ExternalUrl
            isBackendCertificateValidationEnabled=$true
            isHttpOnlyCookieEnabled=$true
            isOnPremPublishingEnabled=$true
            isPersistentCookieEnabled=$false
            isSecureCookieEnabled=$true
            isStateSessionEnabled=$true
            # WebPort must see the external host: the add-on accepts its callback only on that origin.
            isTranslateHostHeaderEnabled=$false
            isTranslateLinksInBodyEnabled=$false
        }
        Invoke-WebPortGraph -ApiVersion beta -Method PATCH -Path "applications/$($application.id)" -Body @{onPremisesPublishing=$proxySettings} | Out-Null
        Invoke-WebPortGraph -ApiVersion beta -Method PUT -Path "applications/$($application.id)/connectorGroup/`$ref" -Body @{
            '@odata.id'="https://graph.microsoft.com/beta/onPremisesPublishingProfiles/applicationProxy/connectorGroups/$($proxyGroup.id)"
        } | Out-Null
        $verifiedProxyApp=Invoke-WebPortGraph -ApiVersion beta -Path ("applications/$($application.id)"+'?$select=id,onPremisesPublishing')
        $verifiedProxyGroup=Invoke-WebPortGraph -ApiVersion beta -Path "applications/$($application.id)/connectorGroup"
        if($verifiedProxyApp.id -ne $application.id -or $verifiedProxyApp.onPremisesPublishing.isOnPremPublishingEnabled -ne $true -or
            $verifiedProxyApp.onPremisesPublishing.internalUrl -ne $ServerUrl -or $verifiedProxyApp.onPremisesPublishing.externalUrl -ne $ExternalUrl -or
            $verifiedProxyApp.onPremisesPublishing.isTranslateHostHeaderEnabled -ne $false -or
            $verifiedProxyApp.onPremisesPublishing.externalAuthenticationType -ne 'aadPreAuthentication' -or $verifiedProxyGroup.id -ne $proxyGroup.id){
            Stop-WebPortRegistrationValidation 'Entra Application Proxy verification failed: publishing settings or connector group mismatch'
        }
        ok "Application Proxy aktiv: $ExternalUrl → $ServerUrl via $($proxyGroup.name) (verifierat)"
    }
    step 'Ger admin consent för hela organisationen'
    $stage = 'grant tenant-wide delegated consent'
    if($grant){
        $consentScopes=@($grant.scope -split ' ' | Where-Object {$_})
        foreach($scope in $requestedScopes){if($consentScopes -cnotcontains $scope){$consentScopes+=$scope}}
        Invoke-WebPortGraph -Method PATCH -Path "oauth2PermissionGrants/$($grant.id)" -Body @{scope=($consentScopes -join ' ')} | Out-Null
    }else{
        $grant = Invoke-WebPortGraph -Method POST -Path 'oauth2PermissionGrants' -Body @{
            clientId=$principal.id; resourceId=$graph.id; consentType='AllPrincipals'; principalId=$null; scope=($requestedScopes -join ' ')
        }
    }
    if (-not $grant.id) { throw 'Consent grant ID missing' }
    ok "Consent (AllPrincipals): $($requestedScopes -join ', ')"
    step 'Verifierar appregistrering och consent'
    $stage = 'verify application and consent'
    $verifiedApp = Invoke-WebPortGraph -Path ("applications/$($application.id)" + '?$select=id,appId,displayName,signInAudience,isFallbackPublicClient,web,passwordCredentials,requiredResourceAccess,optionalClaims,appRoles')
    $verifiedGrant = Invoke-WebPortGraph -Path "oauth2PermissionGrants/$($grant.id)"
    Assert-WebPortRegistration $verifiedApp $verifiedGrant $principal.id $graph.id
    ok 'Kontotyp, redirect URI, implicit grant, behörigheter, claims, approller och consent stämmer'
    $secret = $null; $secretKeyId=$null; $secretExpiry=$null; $secretStatus='Skipped'
    if (-not $SkipClientSecret) {
        if ($newApplication -or $Force) {
            step 'Skapar klienthemlighet'
            $stage='create client secret'
            $expires=[DateTimeOffset]::UtcNow.AddMonths($SecretLifetimeMonths)
            $credential=Invoke-WebPortGraph -Method POST -Path "applications/$($application.id)/addPassword" -Body @{
                passwordCredential=@{ displayName='WebPort Entra integration'; endDateTime=$expires.ToString('o') }
            }
            if (-not $credential.secretText) { throw 'Graph did not return the one-time secret value' }
            $secret=ConvertTo-SecureString -String $credential.secretText -AsPlainText -Force
            $secretKeyId=$credential.keyId; $secretExpiry=$credential.endDateTime; $secretStatus='Created'
            $credential.secretText=$null
            ok "Hemlighet skapad (Secret ID $secretKeyId), giltig till $(([DateTimeOffset]$secretExpiry).ToString('yyyy-MM-dd'))"
            ok 'Värdet returneras som SecureString i ClientSecret och visas inte'
        } else { $secretStatus='ExistingSecretNotRetrievable' }
    }
    $registrationResult=[pscustomobject]@{
        Plant=$DisplayName; DisplayName=$DisplayName; ServerUrl=$ServerUrl
        TenantId=$TenantId.ToString('D'); ClientId=[string]$application.appId
        ApplicationObjectId=[string]$application.id; ServicePrincipalObjectId=[string]$principal.id
        RedirectUri=$RedirectUri.OriginalString
        Authority="https://login.microsoftonline.com/$($TenantId.ToString('D'))/v2.0"
        Scopes=($requestedScopes -join ' '); ConsentedScopes=[string]$verifiedGrant.scope
        AppRoles=@($verifiedApp.appRoles)
        AdminConsentGranted=$true; ConsentGrantId=[string]$verifiedGrant.id
        ApplicationProxyEnabled=($null -ne $proxyGroup); ApplicationProxyInternalUrl=$ServerUrl; ApplicationProxyExternalUrl=$(if($proxyGroup){$ExternalUrl}else{$null})
        ApplicationProxyConnectorGroupId=$(if($proxyGroup){[string]$proxyGroup.id}else{$null}); ApplicationProxyConnectorGroupName=$(if($proxyGroup){[string]$proxyGroup.name}else{$null})
        ClientSecret=$secret; SecretKeyId=$secretKeyId; SecretExpiresUtc=$secretExpiry; SecretStatus=$secretStatus
        ApplicationCreated=$newApplication
        ApplicationUpdated=($null -ne $existingApplication)
    }
} catch {
    $appReference=if($null -ne $application) { " ApplicationObjectId: $($application.id)." } else { '' }
    # Graph's own error code and message say why a request was refused. Error responses carry no secret.
    if(-not $script:WebPortRegistrationFailureReason){
        $graphError=$null
        if($_.ErrorDetails -and $_.ErrorDetails.Message){
            try { $detail=($_.ErrorDetails.Message | ConvertFrom-Json).error; if($detail.code){ $graphError="$($detail.code): $($detail.message)" } } catch {}
        }
        if(-not $graphError -and $_.Exception.Message -match '"code"\s*:\s*"([^"]+)"[^}]*?"message"\s*:\s*"([^"]+)"'){ $graphError="$($Matches[1]): $($Matches[2])" }
        $script:WebPortRegistrationFailureReason=if($graphError){"Microsoft Graph: $graphError"}else{$_.Exception.Message}
    }
    err "Registreringen misslyckades i steget '$stage'.$(if($script:WebPortRegistrationFailureReason){" $script:WebPortRegistrationFailureReason"})"
    if($null -ne $application){ warn "Delvisa ändringar kan finnas på app $($application.id). Granska den innan du kör igen med -Force." }
    # Preserve the original error for diagnostics without displaying response bodies that may contain credentials.
    throw [InvalidOperationException]::new("Entra registration failed during '$stage'.$appReference $script:WebPortRegistrationFailureReason Partial changes may exist; inspect this application before retrying with -Force. Check Graph permissions, administrator role and tenant credential policies.", $_.Exception)
}

$proposedSettings=[ordered]@{
    'Plant'=$registrationResult.Plant
    'WebPortServer Url'=$registrationResult.ServerUrl
    'KIClientId'=$registrationResult.ClientId
    'KIExternalTenantId'=$registrationResult.TenantId
    'KIProviderUri'="https://login.microsoftonline.com/$($registrationResult.TenantId)"
    'KICustomPolicy'=''
    'KISignInPolicyName'=''
    # Empty means the add-on derives its callback from WebPortServer Url (no Application Proxy).
    'EntraRedirectUri'=$(if($registrationResult.ApplicationProxyEnabled){$registrationResult.RedirectUri}else{''})
    # Turns on WebPort's single sign-on (the SSO button on the sign-in page, which goes to /access/loginsso).
    'AuthenticationTypeKI'='1'
}
$registrationResult | Add-Member NoteProperty DatabasePath $DatabasePath
$registrationResult | Add-Member NoteProperty SettingsDeviceGuid $SettingsDeviceGuid.ToString('D')
$registrationResult | Add-Member NoteProperty ProposedDatabaseSettings $proposedSettings
$registrationResult | Add-Member NoteProperty DatabaseUpdated $false
$registrationResult | Add-Member NoteProperty DatabaseBackupPath $null
$registrationResult | Add-Member NoteProperty DatabaseUpdateError $null
$registrationResult | Add-Member NoteProperty SecretStored $false
$registrationResult | Add-Member NoteProperty SecretStoreError $null
if($UpdateDatabase) {
    try {
        step 'Stoppar WebPort'
        Stop-WebPortInstallation -InstallationPath $WebPortProgPath
        step 'Skriver inställningar till SQLite'
        $written=Write-WebPortIamSettings -DatabasePath $DatabasePath -WebPortProgPath $WebPortProgPath -SettingsDeviceGuid $SettingsDeviceGuid -Values $proposedSettings -ExpectedSettings $databaseSettings
        $registrationResult.DatabaseUpdated=$true
        $registrationResult.DatabaseBackupPath=$written.BackupPath
        ok "Backup: $($written.BackupPath)"
        foreach($key in $proposedSettings.Keys){
            $before=if($databaseSettings.ContainsKey($key)){$databaseSettings[$key]}else{$null}
            if($null -ne $before -and [string]$before -ceq [string]$proposedSettings[$key]){ ok ("{0,-19} {1} (oförändrat)" -f $key,(Format-SettingValue $before)) }
            else { ok ("{0,-19} {1} → {2}" -f $key,(Format-SettingValue $before),(Format-SettingValue $proposedSettings[$key])) }
        }
    } catch {
        # Keep the result and the one-time secret available if a later local write fails.
        $registrationResult.DatabaseUpdateError=$_.Exception.Message
        err "SQLite uppdaterades INTE: $($_.Exception.Message)"
        Write-Warning 'Entra registration succeeded, but SQLite was NOT updated. Inspect DatabaseUpdateError and save the returned ClientSecret. WebPort may remain stopped.'
    }
    if($registrationResult.DatabaseUpdated -and $registrationResult.ClientSecret){
        $secretAccount=if($webPortRuntime.Mode -eq 'Program'){$webPortRuntime.Account}else{'LocalSystem'}
        step "Sparar klienthemligheten krypterad för kontot som kör WebPort ($secretAccount)"
        try {
            Save-WebPortEntraClientSecret -Secret $registrationResult.ClientSecret -Tenant ([guid]$registrationResult.TenantId) -Client ([guid]$registrationResult.ClientId) -AccountSid $(if($webPortRuntime.Mode -eq 'Program'){$webPortRuntime.Sid}else{'S-1-5-18'})
            $registrationResult.SecretStored=$true
            ok "Sparad i SQLite som EntraClientSecretProtected (DPAPI, $secretAccount, bunden till tenant och klient-ID)"
            ok $(if($webPortRuntime.Mode -eq 'Program'){'Överföringsfilen är borttagen'}else{'Det tillfälliga SYSTEM-jobbet och överföringsfilen är borttagna'})
        } catch {
            $registrationResult.SecretStoreError=$_.Exception.Message
            err "Klienthemligheten sparades INTE: $($_.Exception.Message)"
            warn 'Hemligheten finns kvar i resultatet (ClientSecret). Kör scriptet igen med -Force för att skapa och spara en ny.'
        }
    }
    if($registrationResult.DatabaseUpdated){
        $addonInstalled=Test-Path -LiteralPath (Join-Path $WebPortProgPath 'addons\EntraPkce\ExtEntraPkce.dll')
        if($webPortRuntime -and $webPortRuntime.Running -and $addonInstalled){
            step 'Startar WebPort igen'
            Start-WebPortRuntime $webPortRuntime $WebPortProgPath
            ok 'WebPort körs med de nya inställningarna'
        } else {
            warn 'WebPort är stoppad. Den startas när tillägget installeras (Install-EntraModule.ps1).'
        }
    }
}

Write-Host ''
step 'Sammanfattning'
setting 'Appnamn' $registrationResult.DisplayName
setting 'Åtgärd' $(if($registrationResult.ApplicationCreated){'Ny app skapad'}else{'Befintlig app uppdaterad'})
setting 'Tenant-ID' $registrationResult.TenantId
setting 'Klient-ID' $registrationResult.ClientId
setting 'Redirect URI' $registrationResult.RedirectUri
if($registrationResult.ApplicationProxyEnabled){
    setting 'Application Proxy' "$($registrationResult.ApplicationProxyExternalUrl) → $($registrationResult.ApplicationProxyInternalUrl) via $($registrationResult.ApplicationProxyConnectorGroupName)"
    setting 'Inloggning' "$($registrationResult.ApplicationProxyExternalUrl)entra/login"
} else {
    setting 'Application Proxy' $(if($SkipApplicationProxy){'Hoppades över (-SkipApplicationProxy)'}else{'Ingen connector konfigurerad i Entra'})
    setting 'Inloggning' "$($registrationResult.ApplicationProxyInternalUrl)entra/login"
}
setting 'Admin consent' $registrationResult.ConsentedScopes
setting 'Klienthemlighet' $(switch($registrationResult.SecretStatus){'Created'{"Skapad, giltig till $(([DateTimeOffset]$registrationResult.SecretExpiresUtc).ToString('yyyy-MM-dd'))"}'Skipped'{'Ingen skapad (-SkipClientSecret)'}default{'Ingen ny; befintliga hemligheter kan inte läsas från Entra'}})
setting 'SQLite' $(if($registrationResult.DatabaseUpdated){'Uppdaterad'}elseif($UpdateDatabase){'INTE uppdaterad, se DatabaseUpdateError'}else{'Inte ändrad'})
setting 'Hemlighet i SQLite' $(if($registrationResult.SecretStored){'Sparad (EntraClientSecretProtected)'}elseif($registrationResult.SecretStoreError){'INTE sparad, se SecretStoreError'}elseif($registrationResult.ClientSecret){'Inte sparad (-UpdateDatabase:$false)'}else{'Oförändrad'})
Write-Host ''
step 'Nästa steg'
Write-Host '   1. Installera tillägget om det inte redan är installerat:'
Write-Host '        .\Install-EntraModule.ps1' -ForegroundColor White
Write-Host '   2. Tilldela användare eller grupper approller i Entra:'
Write-Host "        Enterprise applications → $($registrationResult.DisplayName) → Users and groups" -ForegroundColor White
if($registrationResult.SecretStatus -eq 'Created'){
    $secretExpiryDate=([DateTimeOffset]$registrationResult.SecretExpiresUtc).ToLocalTime().ToString('yyyy-MM-dd')
    Write-Host ''
    warn "OBS: klienthemligheten går ut $secretExpiryDate. Byt den före dess, annars slutar Entra-inloggningen att fungera."
    warn 'Skapa en ny hemlighet och spara den i SQLite med:'
    Write-Host '        .\Register-WebPortEntraApp.ps1 -Force' -ForegroundColor Yellow
    warn "Ta sedan bort den gamla under App registrations → $($registrationResult.DisplayName) → Certificates & secrets."
    warn 'Lägg gärna in en påminnelse i kalendern en månad innan.'
}
Write-Host "`n✓ KLART" -ForegroundColor Green
$registrationResult
