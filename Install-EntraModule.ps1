#Requires -Version 7.0
<#
.SYNOPSIS
Installs, upgrades, rolls back or uninstalls the WebPort Entra add-on (ExtEntraPkce.dll).
.DESCRIPTION
Run Register-WebPortEntraApp.ps1 first. It writes the Entra settings and the encrypted client secret to
WebPort's SQLite database. This script installs the add-on that uses them. It needs no other script.

Before any change, the script lists exactly what it will do. Run with -WhatIf to see only the plan.

  • Checks that the Entra settings and the encrypted client secret exist in SQLite
  • Backs up WebPortServer.exe.config, the current add-on DLL and the SQLite database to
    C:\ProgramData\WebPort\Entra\install-backups\<timestamp>
  • Stops WebPortServer
  • Copies ExtEntraPkce.dll to <WebPortProgPath>\addons\EntraPkce\
  • Writes no settings. Without an EntraRolesEnabled row the add-on uses Entra app roles; only the row
    EntraRolesEnabled=0 in Settings selects local permissions.
    A redundant EntraRolesEnabled=1 from older installers is removed. The add-on creates no tables: it finds
    the WebPort account whose user name is the user's UPN, and the role mapping is built in
  • Registers the event log source "WebPort Entra" (Windows Application log) for sign-in events
  • Removes the legacy keys WebPort.Entra.ModuleConfig and WebPort.Entra.ClientSecretFile from
    WebPortServer.exe.config if they are present
  • Compares the WebPort libraries the DLL was built against with the installed ones. Where a version
    differs, adds a bindingRedirect to WebPortServer.exe.config, so that WebPort can load the add-on. No
    WebPort binary is changed. Run the installer again after every WebPort upgrade, which replaces the file
  • Starts WebPortServer and waits until https://<server>/entra/status reports the copied version as active
  • Leaves the WebPortEntra* tables of version 1.2 untouched, so that a rollback to 1.2 keeps working;
    1.3 does not use them
  • Restores the DLL and WebPortServer.exe.config automatically if any step fails
.PARAMETER ModulePath
The add-on DLL. Default: ExtEntraPkce.dll next to this script, otherwise Development\dist\EntraModule\ExtEntraPkce.dll.
.PARAMETER Rollback
Restore the add-on DLL and WebPortServer.exe.config from -BackupPath.
.PARAMETER Uninstall
Remove the add-on: deletes addons\EntraPkce, removes only the bindingRedirects this installer added to
WebPortServer.exe.config and removes the event log source. The settings in SQLite are kept, so that a
later installation works at once. KIClientId, KIExternalTenantId, KIProviderUri and WebPortServer Url are
never touched; they belong to WebPort's own Kiona sign-in. Backs up first and restores on failure.
.PARAMETER RemoveSettings
With -Uninstall: also deletes EntraClientSecretProtected, EntraRedirectUri and EntraRolesEnabled from
Settings, and drops the WebPortEntra* tables of version 1.2.
.EXAMPLE
.\Install-EntraModule.ps1 -WhatIf
Show the installation plan without changing anything.
.EXAMPLE
.\Install-EntraModule.ps1
Install or upgrade the add-on.
.EXAMPLE
.\Install-EntraModule.ps1 -Rollback -BackupPath 'C:\ProgramData\WebPort\Entra\install-backups\20261025-150000-000'
.EXAMPLE
.\Install-EntraModule.ps1 -Uninstall -WhatIf
Show what an uninstallation would remove.
.EXAMPLE
.\Install-EntraModule.ps1 -Uninstall -RemoveSettings
Remove the add-on together with its settings in SQLite.
.NOTES
Requires PowerShell 7 and administrative privileges.
#>
[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
    [Alias('WebPortPath')][string]$WebPortProgPath='C:\Program Files\WebPort',
    [string]$WebPortDataPath='C:\ProgramData\WebPort',
    [string]$ModulePath,
    [switch]$Rollback,
    [string]$BackupPath,
    [switch]$Uninstall,
    [switch]$RemoveSettings
)
if($Uninstall -and $Rollback){ throw 'Use either -Uninstall or -Rollback' }
if($RemoveSettings -and -not $Uninstall){ throw '-RemoveSettings is only used with -Uninstall' }

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
function setting($name,$value){ Write-Host ("   {0,-24} {1}" -f ($name+':'),$value) }
function detail($value){ Write-Host ("   {0,-24}   {1}" -f '',$value) }

$ErrorActionPreference='Stop'

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
    if (Test-Path -LiteralPath $microsoft) { throw 'Use PowerShell 7 (pwsh) for the installed Microsoft.Data.Sqlite provider. Windows PowerShell requires System.Data.SQLite.' }
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

function Invoke-WebPortSql {
    param([Parameter(Mandatory=$true)]$Connection,[Parameter(Mandatory=$true)][string]$Sql,[hashtable]$Parameters=@{},$Transaction,[switch]$Scalar)
    $command=$Connection.CreateCommand()
    try {
        if($null -ne $Transaction){$command.Transaction=$Transaction}
        $command.CommandText=$Sql
        foreach($key in $Parameters.Keys){[void]$command.Parameters.AddWithValue($key,$Parameters[$key])}
        if($Scalar){return $command.ExecuteScalar()}
        return $command.ExecuteNonQuery()
    } finally {$command.Dispose()}
}

# .NET Framework binds a strong-named reference only to the exact version it was compiled against. When the
# add-on was built against another WebPort version, WebPort cannot load it at all. A bindingRedirect in
# WebPortServer.exe.config points the add-on's reference to the installed version; WebPort's binaries stay
# unchanged, and the add-on's own startup check then decides whether this WebPort version is compatible.
# WebPortCommon.dll is obfuscated: .NET (PowerShell 7) cannot read its metadata, so Windows PowerShell 5.1
# (.NET Framework) reads the assembly identities.
function Get-WebPortAssemblyReferences {
    param([Parameter(Mandatory=$true)][string]$ModulePath,[Parameter(Mandatory=$true)][string]$ProgramPath)
    $quote={param($s) "'"+($s -replace "'","''")+"'"}
    $script=@"
`$ErrorActionPreference='Stop';`$ProgressPreference='SilentlyContinue'
function token(`$n){ `$t=`$n.GetPublicKeyToken(); if(`$t){ (@(`$t) | ForEach-Object { `$_.ToString('x2') }) -join '' } else { '' } }
`$module=[Reflection.Assembly]::ReflectionOnlyLoadFrom($(& $quote $ModulePath))
`$result=foreach(`$reference in `$module.GetReferencedAssemblies()){
    `$file=Join-Path $(& $quote $ProgramPath) (`$reference.Name+'.dll')
    if(-not (Test-Path -LiteralPath `$file)){ continue }
    `$installed=[Reflection.AssemblyName]::GetAssemblyName(`$file)
    [pscustomobject]@{Name=`$reference.Name;Token=(token `$reference);Referenced=`$reference.Version.ToString();Installed=`$installed.Version.ToString();InstalledToken=(token `$installed)}
}
ConvertTo-Json -InputObject @(`$result) -Compress
"@
    $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($script))
    $json=& (Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe') -NoProfile -NonInteractive -OutputFormat Text -EncodedCommand $encoded
    if($LASTEXITCODE -ne 0 -or -not $json){throw 'Could not read the assembly versions with Windows PowerShell 5.1'}
    @(($json -join '') | ConvertFrom-Json)
}

function Test-VersionInRange([version]$Version,[string]$Range) {
    $bounds=$Range.Split('-')
    return $Version -ge [version]$bounds[0] -and $Version -le [version]$bounds[-1]
}

# Returns one entry per reference that needs a new bindingRedirect. A reference that an existing redirect
# already sends to the installed version needs nothing; one that is redirected elsewhere stops the install.
function Get-WebPortBindingRedirectPlan {
    param([Parameter(Mandatory=$true)][xml]$Config,[Parameter(Mandatory=$true)]$References)
    $ns=[Xml.XmlNamespaceManager]::new($Config.NameTable);$ns.AddNamespace('b','urn:schemas-microsoft-com:asm.v1')
    foreach($reference in $References){
        if(-not $reference.Token -or $reference.Token -ne $reference.InstalledToken -or $reference.Referenced -eq $reference.Installed){ continue }
        $dependent=$Config.SelectNodes('/configuration/runtime/b:assemblyBinding/b:dependentAssembly',$ns) | Where-Object {
            $identity=$_.SelectSingleNode('b:assemblyIdentity',$ns)
            $identity -and $identity.GetAttribute('name') -ieq $reference.Name -and $identity.GetAttribute('publicKeyToken') -ieq $reference.Token
        } | Select-Object -First 1
        $covering=if($dependent){ @($dependent.SelectNodes('b:bindingRedirect',$ns)) | Where-Object { Test-VersionInRange $reference.Referenced $_.GetAttribute('oldVersion') } | Select-Object -First 1 }
        if($covering){
            if($covering.GetAttribute('newVersion') -eq $reference.Installed){ continue }
            throw "WebPortServer.exe.config redirects $($reference.Name) $($reference.Referenced) to $($covering.GetAttribute('newVersion')), but $($reference.Installed) is installed"
        }
        [pscustomobject]@{Name=$reference.Name;Token=$reference.Token;From=$reference.Referenced;To=$reference.Installed;Dependent=$dependent}
    }
}

function Add-WebPortBindingRedirect {
    param([Parameter(Mandatory=$true)][xml]$Config,[Parameter(Mandatory=$true)]$Redirect)
    $asm='urn:schemas-microsoft-com:asm.v1'
    $ns=[Xml.XmlNamespaceManager]::new($Config.NameTable);$ns.AddNamespace('b',$asm)
    $dependent=$Redirect.Dependent
    if(-not $dependent){
        $runtime=$Config.SelectSingleNode('/configuration/runtime')
        if(-not $runtime){ $runtime=$Config.DocumentElement.AppendChild($Config.CreateElement('runtime')) }
        $binding=$runtime.SelectSingleNode('b:assemblyBinding',$ns)
        if(-not $binding){ $binding=$runtime.AppendChild($Config.CreateElement('assemblyBinding',$asm)) }
        $dependent=$binding.AppendChild($Config.CreateElement('dependentAssembly',$asm))
        $identity=$dependent.AppendChild($Config.CreateElement('assemblyIdentity',$asm))
        $identity.SetAttribute('name',$Redirect.Name);$identity.SetAttribute('publicKeyToken',$Redirect.Token);$identity.SetAttribute('culture','neutral')
    }
    [void]$dependent.AppendChild($Config.CreateComment(' Added by Install-EntraModule.ps1: ExtEntraPkce.dll was built against another version '))
    $element=$dependent.AppendChild($Config.CreateElement('bindingRedirect',$asm))
    $element.SetAttribute('oldVersion',$Redirect.From);$element.SetAttribute('newVersion',$Redirect.To)
}

# Finds the redirects Add-WebPortBindingRedirect wrote (the element right after its marker comment) and,
# with -Remove, deletes them. A dependentAssembly left without any redirect is deleted as well, so WebPort's
# own entries stay exactly as they were.
function Get-WebPortOwnBindingRedirects {
    param([Parameter(Mandatory=$true)][xml]$Config,[switch]$Remove)
    $ns=[Xml.XmlNamespaceManager]::new($Config.NameTable);$ns.AddNamespace('b','urn:schemas-microsoft-com:asm.v1')
    foreach($element in @($Config.SelectNodes('/configuration/runtime/b:assemblyBinding/b:dependentAssembly/b:bindingRedirect',$ns))){
        $previous=$element.PreviousSibling
        while($previous -is [Xml.XmlWhitespace]){ $previous=$previous.PreviousSibling }
        if(-not ($previous -is [Xml.XmlComment] -and $previous.Value -match 'Added by Install-EntraModule\.ps1')){ continue }
        $dependent=$element.ParentNode
        [pscustomobject]@{Name=$dependent.SelectSingleNode('b:assemblyIdentity',$ns).GetAttribute('name');From=$element.GetAttribute('oldVersion');To=$element.GetAttribute('newVersion')}
        if($Remove){
            [void]$dependent.RemoveChild($previous);[void]$dependent.RemoveChild($element)
            if(-not $dependent.SelectSingleNode('b:bindingRedirect',$ns)){ [void]$dependent.ParentNode.RemoveChild($dependent) }
        }
    }
}

# The add-on has no tables of its own. These were used by 1.2 only; they are left in place, so that a
# rollback to 1.2 keeps its data.
$legacyTables=@('WebPortEntraUsers','WebPortEntraRolePermissions','WebPortEntraDiagnostics','WebPortEntraPendingUsers','WebPortEntraProvisionedUsers')
$eventSource='WebPort Entra'
$legacyKeys=@('WebPort.Entra.ModuleConfig','WebPort.Entra.ClientSecretFile')
$configPath=Join-Path $WebPortProgPath 'WebPortServer.exe.config'
$moduleDirectory=Join-Path $WebPortProgPath 'addons\EntraPkce'
$targetPath=Join-Path $moduleDirectory 'ExtEntraPkce.dll'
$databasePath=Join-Path $WebPortDataPath 'db\webport.sqlite'
$backupRoot=Join-Path $WebPortDataPath 'Entra\install-backups'

$identity=[Security.Principal.WindowsPrincipal]::new([Security.Principal.WindowsIdentity]::GetCurrent())
if(-not $identity.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){
    err 'Kör scriptet i ett PowerShell 7-fönster som administratör.'
    throw 'Run this installer elevated: it must stop WebPort and change its program folder.'
}
step 'Kontrollerar WebPort-installationen'
if(-not (Test-Path -LiteralPath (Join-Path $WebPortProgPath 'WebPortServer.exe') -PathType Leaf)){
    err "WebPortServer.exe finns inte i $WebPortProgPath. Ange rätt sökväg med -WebPortProgPath."
    throw 'WebPortServer.exe is missing in the selected installation'
}
$runtime=Get-WebPortRuntime $WebPortProgPath
if($runtime.Mode -eq 'Service' -and $runtime.PathName -notlike ('*'+(Join-Path $WebPortProgPath 'WebPortServer.exe')+'*')){
    err "Tjänsten WebPortServer pekar inte på $WebPortProgPath. Ange rätt sökväg med -WebPortProgPath."
    throw 'Service does not match selected installation'
}
$wasRunning=$runtime.Running
switch($runtime.Mode){
    'Service' { ok "Tjänst: $($runtime.Name) ($(if($runtime.Running){'körs'}else{'stoppad'}), kör som $($runtime.Account))" }
    'Program' { ok "WebPort körs som program av $($runtime.Account)$(if($runtime.Arguments){" med argumenten: $($runtime.Arguments)"})" }
    default   { warn 'WebPort körs inte, och tjänsten WebPortServer finns inte. Starta WebPort efteråt och kontrollera /entra/status.' }
}
ok "Program: $WebPortProgPath"
ok "Databas: $databasePath"

function Restore-InstallBackup([string]$source) {
    $manifest=Get-Content -LiteralPath (Join-Path $source 'backup.json') -Raw | ConvertFrom-Json
    if($manifest.InstallPath -ne $WebPortProgPath){throw 'Backup belongs to another installation'}
    step "Återställer från $source"
    Stop-WebPortInstallation -InstallationPath $WebPortProgPath
    Copy-Item -LiteralPath (Join-Path $source 'WebPortServer.exe.config') -Destination $configPath -Force
    ok "Återställd: $configPath"
    if($manifest.HadModule){New-Item -ItemType Directory -Path $moduleDirectory -Force | Out-Null; Copy-Item -LiteralPath (Join-Path $source 'ExtEntraPkce.dll') -Destination $targetPath -Force; ok "Återställd: $targetPath"}
    elseif(Test-Path -LiteralPath $targetPath){Remove-Item -LiteralPath $targetPath -Force; ok "Borttagen: $targetPath"}
    # Backups from before program support only knew the service.
    $saved=if($manifest.Runtime){$manifest.Runtime}else{[pscustomobject]@{Mode='Service';Name='WebPortServer';Account='LocalSystem';Sid='S-1-5-18';Arguments=$null}}
    if($manifest.WasRunning){Start-WebPortRuntime $saved $WebPortProgPath}
    else { warn 'WebPort var stoppad före installationen och lämnas stoppad' }
}

$backup=$null
try {
    if($Rollback){
        if(-not $BackupPath -or -not (Test-Path -LiteralPath (Join-Path $BackupPath 'backup.json'))){
            err "Ange en installationsbackup med -BackupPath (en mapp under $backupRoot som innehåller backup.json)."
            throw 'An installation backup is required'
        }
        $manifest=Get-Content -LiteralPath (Join-Path $BackupPath 'backup.json') -Raw | ConvertFrom-Json
        Write-Host ''
        step 'Planerad återställning'
        setting 'Backup' $BackupPath
        setting 'WebPortServer.exe.config' 'Återställs från backupen'
        setting 'Tilläggets DLL' $(if($manifest.HadModule){"Återställs från backupen (version $($manifest.PreviousVersion))"}else{"Tas bort ($targetPath)"})
        setting 'SQLite' 'Ändras inte. Tilläggets tabeller och inställningar ligger kvar och påverkar inte WebPort utan tillägget'
        setting 'WebPort' $(if($manifest.WasRunning){'Stoppas och startas igen som före installationen'}else{'Stoppas och lämnas stoppad'})
        Write-Host ''
        if(-not $PSCmdlet.ShouldProcess($WebPortProgPath,'Återställ Entra-tillägget från backup')){return}
        Restore-InstallBackup $BackupPath
        Write-Host "`n✓ KLART" -ForegroundColor Green
        return
    }

    if($Uninstall){
        step 'Läser den nuvarande installationen'
        $hadModule=Test-Path -LiteralPath $targetPath
        $installedVersion=if($hadModule){[Reflection.AssemblyName]::GetAssemblyName($targetPath).Version}else{$null}
        $moduleFiles=if(Test-Path -LiteralPath $moduleDirectory){@(Get-ChildItem -LiteralPath $moduleDirectory -Recurse -File)}else{@()}
        [xml]$config=Get-Content -LiteralPath $configPath
        $ownRedirects=@(Get-WebPortOwnBindingRedirects $config)
        $hasEventSource=[Diagnostics.EventLog]::SourceExists($eventSource)
        $deviceGuid=Get-WebPortSettingsDeviceGuid $databasePath $WebPortProgPath
        $native=Read-WebPortIamSettings $databasePath $WebPortProgPath $deviceGuid
        $addonKeys=@('EntraClientSecretProtected','EntraRedirectUri','EntraRolesEnabled')
        $connection=New-WebPortIamSqliteConnection $databasePath $WebPortProgPath
        try {
            $connection.Open()
            $presentKeys=@($addonKeys | Where-Object { [int](Invoke-WebPortSql $connection 'SELECT COUNT(*) FROM Settings WHERE DeviceGuid=@dg COLLATE NOCASE AND "Key"=@key' @{'@dg'=$deviceGuid.ToString('D');'@key'=$_} -Scalar) -gt 0 })
            $oldTables=@($legacyTables | Where-Object { [int](Invoke-WebPortSql $connection "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@name" @{'@name'=$_} -Scalar) -gt 0 })
        } finally {$connection.Dispose()}
        $statusUrl=if($native['WebPortServer Url']){([uri]$native['WebPortServer Url']).GetLeftPart([UriPartial]::Authority)+'/entra/status'}else{$null}
        $backup=Join-Path $backupRoot ([DateTime]::Now.ToString('yyyyMMdd-HHmmss-fff'))

        Write-Host ''
        step 'Planerad avinstallation'
        setting 'Tilläggets DLL' $(if($hadModule){"Tas bort: ExtEntraPkce $installedVersion"}else{'Är inte installerad'})
        if($moduleFiles.Count){ detail "Mappen $moduleDirectory tas bort ($($moduleFiles.Count) filer)" }
        setting 'WebPortServer.exe.config' $(if($ownRedirects.Count){'Installationens egna bindingRedirect tas bort:'}else{'Ändras inte (inga egna bindingRedirect)'})
        foreach($redirect in $ownRedirects){ detail "$($redirect.Name) $($redirect.From) → $($redirect.To)" }
        detail 'WebPorts egna poster lämnas orörda'
        setting 'Händelselogg' $(if($hasEventSource){"Källan '$eventSource' tas bort (loggade händelser finns kvar)"}else{"Källan '$eventSource' finns inte"})
        if($RemoveSettings){
            setting 'Settings' $(if($presentKeys.Count){"Tas bort: $($presentKeys -join ', ')"}else{'Inga av tilläggets poster finns'})
            setting 'Tabeller från 1.2' $(if($oldTables.Count){"Tas bort: $($oldTables -join ', ')"}else{'Finns inte'})
        } else {
            setting 'Settings' "Lämnas kvar$(if($presentKeys.Count){": $($presentKeys -join ', ')"}) (ta bort med -RemoveSettings)"
            if($oldTables.Count){ setting 'Tabeller från 1.2' "Lämnas kvar: $($oldTables -join ', ') (ta bort med -RemoveSettings)" }
        }
        setting 'Rörs aldrig' 'KIClientId, KIExternalTenantId, KIProviderUri, WebPortServer Url, AuthenticationTypeKI (WebPorts egna inställningar)'
        if($native['AuthenticationTypeKI'] -eq '1'){ detail 'AuthenticationTypeKI=1: sätt 0 i WebPort om SSO-valet på inloggningssidan ska stängas av' }
        detail 'Konton i Users som tillägget har skapat (lås eller ta bort dem i WebPort vid behov)'
        detail 'Appregistreringen i Entra (ta bort den under App registrations om den inte behövs)'
        setting 'Backup' $backup
        detail "WebPortServer.exe.config$(if($hadModule){', ExtEntraPkce.dll'}), webport.sqlite"
        setting 'WebPort' $(if($wasRunning){"Stoppas och startas igen ($(if($runtime.Mode -eq 'Program'){"program, $($runtime.Account)"}else{"tjänst, $($runtime.Account)"}))"}else{'Körs inte och lämnas stoppad'})
        if($wasRunning -and $statusUrl){ setting 'Kontroll' "$statusUrl ska inte längre svara som tillägget" }
        setting 'Vid fel' "DLL och WebPortServer.exe.config återställs automatiskt$(if($RemoveSettings){'; borttagna poster i SQLite finns i backupens webport.sqlite'})"
        Write-Host ''
        if(-not $PSCmdlet.ShouldProcess($WebPortProgPath,'Avinstallera Entra-tillägget enligt planen ovan')){$backup=$null;return}

        step 'Säkerhetskopierar'
        New-Item -ItemType Directory -Path $backup -Force | Out-Null
        Copy-Item -LiteralPath $configPath -Destination $backup
        if($hadModule){Copy-Item -LiteralPath $targetPath -Destination $backup}
        $source=New-WebPortIamSqliteConnection $databasePath $WebPortProgPath
        $destination=New-WebPortIamSqliteConnection (Join-Path $backup 'webport.sqlite') $WebPortProgPath -Mode Create
        try {
            $source.Open();$destination.Open()
            if($source.GetType().FullName -eq 'Microsoft.Data.Sqlite.SqliteConnection'){$source.BackupDatabase($destination)}
            else{$source.BackupDatabase($destination,'main','main',-1,$null,0)}
        } finally {$destination.Dispose();$source.Dispose()}
        [ordered]@{InstallPath=$WebPortProgPath;WasRunning=$wasRunning;Runtime=$runtime;HadModule=$hadModule;PreviousVersion=[string]$installedVersion;Uninstall=$true} |
            ConvertTo-Json | Set-Content -LiteralPath (Join-Path $backup 'backup.json') -Encoding utf8
        ok "Backup: $backup"

        step 'Stoppar WebPort'
        Stop-WebPortInstallation -InstallationPath $WebPortProgPath

        if($RemoveSettings -and ($presentKeys.Count -or $oldTables.Count)){
            step 'Tar bort tilläggets data i SQLite'
            $connection=New-WebPortIamSqliteConnection $databasePath $WebPortProgPath -Mode ReadWrite
            $transaction=$null
            try {
                $connection.Open();$transaction=$connection.BeginTransaction()
                foreach($key in $presentKeys){ [void](Invoke-WebPortSql $connection 'DELETE FROM Settings WHERE DeviceGuid=@dg COLLATE NOCASE AND "Key"=@key' @{'@dg'=$deviceGuid.ToString('D');'@key'=$key} -Transaction $transaction) }
                # Table names come from the fixed list above, never from input.
                foreach($table in $oldTables){ [void](Invoke-WebPortSql $connection "DROP TABLE IF EXISTS $table" -Transaction $transaction) }
                $transaction.Commit()
            } catch {
                if($null -ne $transaction){try{$transaction.Rollback()}catch{}}
                throw
            } finally {
                if($null -ne $transaction){$transaction.Dispose()}
                $connection.Dispose()
            }
            foreach($key in $presentKeys){ ok "Borttagen: Settings $key" }
            foreach($table in $oldTables){ ok "Borttagen: tabell $table" }
        }

        step 'Tar bort tillägget'
        if(Test-Path -LiteralPath $moduleDirectory){ Remove-Item -LiteralPath $moduleDirectory -Recurse -Force; ok "Borttagen: $moduleDirectory" }
        if($ownRedirects.Count){
            [void](Get-WebPortOwnBindingRedirects $config -Remove)
            $config.Save($configPath)
            foreach($redirect in $ownRedirects){ ok "bindingRedirect borttagen: $($redirect.Name) $($redirect.From) → $($redirect.To)" }
        }
        if($hasEventSource){ [Diagnostics.EventLog]::DeleteEventSource($eventSource); ok "Händelsekälla borttagen: $eventSource" }

        if($wasRunning){
            step 'Startar WebPort'
            Start-WebPortRuntime $runtime $WebPortProgPath
            if($statusUrl){
                $answer=$null
                for($attempt=0;$attempt -lt 15;$attempt++){
                    try {$answer=([string](Invoke-WebRequest -Uri $statusUrl -TimeoutSec 3 -SkipHttpErrorCheck).Content).Trim();break}catch{Start-Sleep -Seconds 2}
                }
                if($answer -like '*Entra PKCE module*'){ throw "WebPort still answers as the add-on: $answer" }
                elseif($null -eq $answer){ warn "WebPort svarade inte på $statusUrl inom 30 sekunder; kontrollera att WebPort fungerar" }
                else { ok 'WebPort svarar utan tillägget' }
            }
        } else { warn 'WebPort var stoppad före avinstallationen och lämnas stoppad' }

        Write-Host ''
        step 'Kvar att göra vid behov'
        Write-Host '   • Lås eller ta bort konton i WebPort som bara användes via Entra.'
        Write-Host '   • Ta bort appen i Entra: App registrations → <appens namn> → Delete.'
        if(-not $RemoveSettings -and $presentKeys.Count){ Write-Host '   • Ta bort tilläggets poster i Settings: kör igen med -Uninstall -RemoveSettings.' }
        Write-Host "`n✓ KLART" -ForegroundColor Green
        return
    }

    step 'Kontrollerar tillägget'
    if([string]::IsNullOrWhiteSpace($ModulePath)){
        $ModulePath=Join-Path $PSScriptRoot 'ExtEntraPkce.dll'
        if(-not (Test-Path -LiteralPath $ModulePath)){$ModulePath=Join-Path $PSScriptRoot 'Development\dist\EntraModule\ExtEntraPkce.dll'}
    }
    if(-not (Test-Path -LiteralPath $ModulePath -PathType Leaf)){
        err "Hittar inte ExtEntraPkce.dll. Lägg den bredvid scriptet eller ange -ModulePath."
        throw 'The add-on DLL is missing'
    }
    $ModulePath=(Resolve-Path -LiteralPath $ModulePath).Path
    $sourceName=[Reflection.AssemblyName]::GetAssemblyName($ModulePath)
    if($sourceName.Name -ne 'ExtEntraPkce'){throw 'Invalid module identity'}
    $sourceHash=(Get-FileHash -LiteralPath $ModulePath -Algorithm SHA256).Hash
    ok "ExtEntraPkce $($sourceName.Version), SHA256 $sourceHash"
    $hadModule=Test-Path -LiteralPath $targetPath
    $installedVersion=if($hadModule){[Reflection.AssemblyName]::GetAssemblyName($targetPath).Version}else{$null}
    # WebPort compatibility is checked by the add-on itself at startup; an incompatible installation never
    # reports active on /entra/status, which triggers the rollback below.

    step 'Läser Entra-inställningar från SQLite'
    $deviceGuid=Get-WebPortSettingsDeviceGuid $databasePath $WebPortProgPath
    $native=Read-WebPortIamSettings $databasePath $WebPortProgPath $deviceGuid
    $tenantGuid=[guid]::Empty;$clientGuid=[guid]::Empty
    $provider=[uri]$null
    $configured=[guid]::TryParse([string]$native['KIExternalTenantId'],[ref]$tenantGuid) -and $tenantGuid -ne [guid]::Empty -and
        [guid]::TryParse([string]$native['KIClientId'],[ref]$clientGuid) -and $clientGuid -ne [guid]::Empty -and
        [uri]::TryCreate([string]$native['KIProviderUri'],[UriKind]::Absolute,[ref]$provider) -and $provider.Host -eq 'login.microsoftonline.com' -and
        $provider.AbsolutePath.Trim('/').Split('/')[0] -ieq $tenantGuid.ToString('D')
    if(-not $configured){
        err 'Entra-inställningarna i SQLite saknas eller är ogiltiga. Kör .\Register-WebPortEntraApp.ps1 först.'
        throw 'Entra settings (KIClientId, KIExternalTenantId, KIProviderUri) are missing or invalid'
    }
    # Users sign in on the callback's origin (the Application Proxy address when EntraRedirectUri is set).
    # The health check stays on the server's own address, which the status page answers on as well.
    $internalOrigin=([uri]$native['WebPortServer Url']).GetLeftPart([UriPartial]::Authority)
    $origin=if($native['EntraRedirectUri']){([uri]$native['EntraRedirectUri']).GetLeftPart([UriPartial]::Authority)}else{$internalOrigin}
    if(-not $origin.StartsWith('https://',[StringComparison]::OrdinalIgnoreCase) -or -not $internalOrigin.StartsWith('https://',[StringComparison]::OrdinalIgnoreCase)){throw 'WebPortServer Url and EntraRedirectUri must be HTTPS addresses'}
    $statusUrl=$internalOrigin+'/entra/status'
    ok "Tenant $($tenantGuid.ToString('D')), klient-ID $($clientGuid.ToString('D'))"
    ok "Callback $origin/access/loginsso"
    if($origin -ne $internalOrigin){ ok "Intern adress $internalOrigin (inloggning som startar där skickas vidare till $origin)" }
    $connection=New-WebPortIamSqliteConnection $databasePath $WebPortProgPath
    try {
        $connection.Open()
        $secrets=@()
        $command=$connection.CreateCommand()
        try {
            $command.CommandText='SELECT "Value" FROM Settings WHERE DeviceGuid=@dg COLLATE NOCASE AND "Key"=''EntraClientSecretProtected'''
            [void]$command.Parameters.AddWithValue('@dg',$deviceGuid.ToString('D'))
            $reader=$command.ExecuteReader();try{while($reader.Read()){$secrets+=$reader.GetString(0)}}finally{$reader.Dispose()}
        } finally {$command.Dispose()}
        $oldTables=@($legacyTables | Where-Object { [int](Invoke-WebPortSql $connection "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@name" @{'@name'=$_} -Scalar) -gt 0 })
        # 1.2 could bind an identity to an account with any name. 1.3 finds the account by the UPN, so such
        # accounts must be renamed to the user's UPN, otherwise the user gets a new account.
        $nonUpnAccounts=@()
        if($oldTables -contains 'WebPortEntraUsers'){
            $command=$connection.CreateCommand()
            try {
                $command.CommandText="SELECT UserName FROM WebPortEntraUsers WHERE UserName NOT LIKE '%@%' ORDER BY UserName"
                $reader=$command.ExecuteReader();try{while($reader.Read()){$nonUpnAccounts+=$reader.GetString(0)}}finally{$reader.Dispose()}
            } finally {$command.Dispose()}
        }
    } finally {$connection.Dispose()}
    $hasEventSource=[Diagnostics.EventLog]::SourceExists($eventSource)
    if($secrets.Count -ne 1 -or -not $secrets[0].StartsWith('dpapi-user-v1:')){
        err 'Den krypterade klienthemligheten saknas i SQLite. Kör .\Register-WebPortEntraApp.ps1 (med -Force för en befintlig app).'
        throw 'EntraClientSecretProtected is missing, duplicated or in an unknown format'
    }
    ok 'Klienthemlighet: finns (EntraClientSecretProtected)'
    # The add-on decrypts the secret as the account WebPort runs as (DPAPI CurrentUser). When that account is
    # the one running this script, decryption is tried here, so a secret stored for another account stops the
    # installation instead of failing every sign-in.
    $currentSid=[Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    if($runtime.Sid -and $runtime.Sid -eq $currentSid){
        $plain=$null
        try {
            Add-Type -AssemblyName System.Security
            $entropy=[Text.Encoding]::UTF8.GetBytes('WebPort.Entra.SettingsSecret:'+$tenantGuid.ToString('D')+':'+$clientGuid.ToString('D'))
            $plain=[Security.Cryptography.ProtectedData]::Unprotect([Convert]::FromBase64String($secrets[0].Substring(14)),$entropy,[Security.Cryptography.DataProtectionScope]::CurrentUser)
            ok "Klienthemligheten kan läsas av $($runtime.Account), som kör WebPort"
        } catch {
            err "Klienthemligheten är krypterad för ett annat konto än $($runtime.Account), som kör WebPort."
            warn "Kör .\Register-WebPortEntraApp.ps1 -Force som $($runtime.Account) för att spara en ny hemlighet för det kontot."
            throw 'The client secret cannot be decrypted by the account WebPort runs as'
        } finally { if($plain){[Array]::Clear($plain,0,$plain.Length)} }
    } elseif($runtime.Mode -eq 'Program'){
        warn "WebPort körs som program av $($runtime.Account). Hemligheten måste vara sparad av Register-WebPortEntraApp.ps1 körd som det kontot."
    } elseif($runtime.Mode -eq 'Service' -and $runtime.Sid -ne 'S-1-5-18'){
        warn "Tjänsten kör som $($runtime.Account), inte LocalSystem. Register-WebPortEntraApp.ps1 sparar hemligheten för LocalSystem eller för programmets konto."
    }
    [xml]$config=Get-Content -LiteralPath $configPath
    if($config.configuration.appSettings.configSource){throw 'External appSettings are not supported'}
    $legacyEntries=@($config.configuration.appSettings.add | Where-Object { $legacyKeys -contains $_.key })
    step 'Jämför tilläggets beroenden med den installerade WebPort-versionen'
    $references=Get-WebPortAssemblyReferences $ModulePath $WebPortProgPath
    foreach($reference in $references){
        if($reference.Referenced -eq $reference.Installed){ ok "$($reference.Name) $($reference.Installed)" }
        else { warn "$($reference.Name): tillägget är byggt mot $($reference.Referenced), installerad är $($reference.Installed)" }
    }
    $redirects=@(Get-WebPortBindingRedirectPlan $config $references)
    $rolesEnabled=$native['EntraRolesEnabled']
    # The status page names the running version, so the health check proves that the copied DLL is the one loaded.
    $expectedStatus="WebPort Entra PKCE module $($sourceName.Version.ToString(3)): active"
    $backup=Join-Path $backupRoot ([DateTime]::Now.ToString('yyyyMMdd-HHmmss-fff'))

    Write-Host ''
    step 'Planerade åtgärder'
    setting 'Läge' $(if($hadModule){"Uppgradering (installerad version $installedVersion → $($sourceName.Version))"}else{'Ny installation'})
    setting 'Tilläggets DLL' $ModulePath
    detail "→ $targetPath"
    setting 'Entra-inställningar' 'Läses av tillägget från SQLite (ändras inte här)'
    setting 'SQLite-tabeller' 'Inga. Tillägget använder WebPort-kontot vars användarnamn är användarens UPN'
    if($oldTables.Count){ detail "Från version 1.2, används inte och lämnas orörda (för återställning): $($oldTables -join ', ')" }
    setting 'Rollmappning' 'Inbyggd: Deny=0, View=1, Basic=2, Advanced=3, System=4, Admin=5'
    setting 'Behörighetsläge' $(switch($rolesEnabled){'0'{'Lokala rättigheter (EntraRolesEnabled=0, oförändrat)'}'1'{'Entra-roller; den överflödiga posten EntraRolesEnabled=1 tas bort'}default{'Entra-roller (ingen post i Settings behövs)'}})
    setting 'Händelselogg' "Application, källa '$eventSource' $(if($hasEventSource){'(finns)'}else{'(registreras)'})"
    if($nonUpnAccounts.Count){
        warn "$($nonUpnAccounts.Count) konton var kopplade till Entra men heter inte som ett UPN. Byt namn på dem till användarens UPN"
        warn "(gemener) före nästa inloggning, annars får användaren ett nytt konto: $($nonUpnAccounts -join ', ')"
    }
    setting 'WebPortServer.exe.config' $(if($legacyEntries.Count -or $redirects.Count){'Ändras (WebPorts binärfiler ändras inte):'}else{'Ändras inte'})
    foreach($entry in $legacyEntries){ detail "Äldre nyckel tas bort: $($entry.key) = $($entry.value)" }
    foreach($redirect in $redirects){ detail "bindingRedirect läggs till: $($redirect.Name) $($redirect.From) → $($redirect.To)" }
    setting 'Backup' "$backup"
    detail "WebPortServer.exe.config$(if($hadModule){', ExtEntraPkce.dll'}), webport.sqlite"
    setting 'WebPort' $(switch($runtime.Mode){'Service'{"Tjänsten $(if($wasRunning){'stoppas och startas igen'}else{'är stoppad och startas'}) ($($runtime.Account))"}'Program'{"Programmet stoppas och startas igen ($($runtime.Account))"}default{'Körs inte; starta WebPort själv efter installationen'}})
    setting 'Hälsokontroll' "$statusUrl ska svara '$expectedStatus' inom 60 s"
    setting 'Vid fel' 'DLL och WebPortServer.exe.config återställs automatiskt; händelsekällan ligger kvar'
    Write-Host ''
    if(-not $PSCmdlet.ShouldProcess($WebPortProgPath,'Installera Entra-tillägget enligt planen ovan')){$backup=$null;return}

    step 'Säkerhetskopierar'
    New-Item -ItemType Directory -Path $backup -Force | Out-Null
    Copy-Item -LiteralPath $configPath -Destination $backup
    if($hadModule){Copy-Item -LiteralPath $targetPath -Destination $backup}
    $source=New-WebPortIamSqliteConnection $databasePath $WebPortProgPath
    $destination=New-WebPortIamSqliteConnection (Join-Path $backup 'webport.sqlite') $WebPortProgPath -Mode Create
    try {
        $source.Open();$destination.Open()
        # SQLite's backup API includes committed WAL data; copying only the .sqlite file would not.
        if($source.GetType().FullName -eq 'Microsoft.Data.Sqlite.SqliteConnection'){$source.BackupDatabase($destination)}
        else{$source.BackupDatabase($destination,'main','main',-1,$null,0)}
    } finally {$destination.Dispose();$source.Dispose()}
    [ordered]@{InstallPath=$WebPortProgPath;WasRunning=$wasRunning;Runtime=$runtime;HadModule=$hadModule;PreviousVersion=[string]$installedVersion;NewVersion=[string]$sourceName.Version;NewSha256=$sourceHash} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $backup 'backup.json') -Encoding utf8
    ok "Backup: $backup"

    step 'Stoppar WebPort'
    Stop-WebPortInstallation -InstallationPath $WebPortProgPath

    step 'Installerar tillägget'
    New-Item -ItemType Directory -Path $moduleDirectory -Force | Out-Null
    Copy-Item -LiteralPath $ModulePath -Destination $targetPath -Force
    ok "DLL kopierad: $targetPath"

    step 'Inställningar och händelselogg'
    if($rolesEnabled -eq '1'){
        # '1' means the same as no row (roles), so the redundant row from older installers is removed.
        $connection=New-WebPortIamSqliteConnection $databasePath $WebPortProgPath -Mode ReadWrite
        try {
            $connection.Open()
            [void](Invoke-WebPortSql $connection 'DELETE FROM Settings WHERE DeviceGuid=@dg COLLATE NOCASE AND "Key"=''EntraRolesEnabled'' AND "Value"=''1''' @{'@dg'=$deviceGuid.ToString('D')})
        } finally {$connection.Dispose()}
        ok 'EntraRolesEnabled=1 borttagen (rolläget gäller utan post)'
    } else {
        ok "Behörighetsläge: $(if($rolesEnabled -eq '0'){'lokala rättigheter (EntraRolesEnabled=0, oförändrat)'}else{'Entra-roller (ingen post behövs)'})"
    }
    if(-not $hasEventSource){
        [Diagnostics.EventLog]::CreateEventSource($eventSource,'Application')
        ok "Händelsekälla registrerad: Application / $eventSource"
    }

    if($legacyEntries.Count -or $redirects.Count){
        step 'Uppdaterar WebPortServer.exe.config'
        foreach($entry in $legacyEntries){ [void]$config.configuration.appSettings.RemoveChild($entry); ok "Borttagen: $($entry.key)" }
        foreach($redirect in $redirects){ Add-WebPortBindingRedirect $config $redirect; ok "bindingRedirect: $($redirect.Name) $($redirect.From) → $($redirect.To)" }
        $config.Save($configPath)
    }

    if($runtime.Mode -eq 'None'){
        warn 'WebPort körde inte och har ingen tjänst, så scriptet kan inte starta det.'
        warn "Starta WebPort och kontrollera att $statusUrl svarar '$expectedStatus'."
        Write-Host "`n✓ KLART (utan hälsokontroll)" -ForegroundColor Green
        return
    }
    step 'Startar WebPort'
    Start-WebPortRuntime $runtime $WebPortProgPath
    step "Väntar på att tillägget blir aktivt ($statusUrl)"
    $healthy=$false;$lastStatus=$null
    for($attempt=0;$attempt -lt 30;$attempt++){
        try {$probe=Invoke-WebRequest -Uri $statusUrl -TimeoutSec 3; $lastStatus=([string]$probe.Content).Trim(); if($probe.StatusCode -eq 200 -and $lastStatus -eq $expectedStatus){$healthy=$true;break}}catch{}
        Start-Sleep -Seconds 2
    }
    if(-not $healthy){
        err "Tillägget blev inte aktivt inom 60 sekunder$(if($lastStatus){" (svarade: '$lastStatus', förväntat: '$expectedStatus')"})."
        warn 'Vanliga orsaker: WebPort saknar interna delar som tillägget kräver (efter en WebPort-uppgradering),'
        warn 'ogiltiga Entra-inställningar, eller att servernamnet inte nås härifrån. Saknas händelsen'
        warn 'module_initialized helt kunde WebPort inte läsa in DLL:en; kontrollera bindingRedirect ovan.'
        warn "Se händelseloggen: Get-WinEvent -FilterHashtable @{LogName='Application';ProviderName='$eventSource'} -MaxEvents 20"
        throw 'Module health endpoint did not report the installed version as active'
    }
    ok "Tillägget svarar: $expectedStatus"

    Write-Host ''
    step 'Sammanfattning'
    setting 'Tillägg' "ExtEntraPkce $($sourceName.Version) i $moduleDirectory"
    setting 'Klient-ID' $clientGuid.ToString('D')
    setting 'Inloggning' "$origin/entra/login"
    setting 'Backup' $backup
    Write-Host ''
    step 'Nästa steg'
    Write-Host '   Tilldela användare eller grupper approller i Entra. Konton skapas i WebPort vid första inloggningen:'
    Write-Host '        Enterprise applications → <appens namn> → Users and groups' -ForegroundColor White
    Write-Host '   Inloggningar loggas i Windows händelselogg:'
    Write-Host "        Get-WinEvent -FilterHashtable @{LogName='Application';ProviderName='$eventSource'} -MaxEvents 20" -ForegroundColor White
    Write-Host "`n✓ KLART" -ForegroundColor Green
} catch {
    err "Installationen misslyckades: $($_.Exception.Message)"
    if($backup -and (Test-Path -LiteralPath (Join-Path $backup 'backup.json'))){
        try {Restore-InstallBackup $backup;ok 'Tidigare DLL och konfiguration återställda'}
        catch{err "Återställningen misslyckades. Kör manuellt: .\Install-EntraModule.ps1 -Rollback -BackupPath '$backup'"}
    }
    throw
}
