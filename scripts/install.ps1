# OneNote Sync Watcher — single-script INSTALL / UPDATE. Run as Administrator.
#
#   & "D:\Src\Personal\onenote-watcher\scripts\install.ps1"              # full install (first time)
#   & "D:\Src\Personal\onenote-watcher\scripts\install.ps1" -UpdateOnly  # after code changes: rebuild, swap binaries, restart
#
# ONE root folder holds everything: executables, config.ini, status.json, section-names.json and logs\.
# Default root: C:\ProgramData\OneNoteWatcher (both the SYSTEM collector and the user's tray can write there;
# Program Files cannot be written by the unelevated tray, which is why it is not used).
#
# Full install: builds Release, copies binaries, keeps an existing config.ini, grants Users modify on the
# root, registers two Scheduled Tasks (collector as SYSTEM at startup; tray at your logon), starts both.
# -UpdateOnly: everything except the task registration/ACL — needed whenever the INSTALLED binaries must
# change, because the tasks run the copies in the root folder, not your build output.
param(
    [string]$InstallDir = "C:\ProgramData\OneNoteWatcher",
    [string]$TrayUser   = "",     # default: the logged-on console user (works when elevated under another account)
    [switch]$UpdateOnly,
    [switch]$SkipBuild
)
$ErrorActionPreference = 'Stop'
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole('Administrators')) {
    Write-Host "NOT ELEVATED. Re-open PowerShell as Administrator and run again." -ForegroundColor Red; return
}
$repo = Split-Path $PSScriptRoot -Parent
$CollectorTask = "OneNoteWatcher Collector"; $TrayTask = "OneNoteWatcher Tray"
if (-not $TrayUser) { $TrayUser = (Get-CimInstance Win32_ComputerSystem).UserName; if (-not $TrayUser) { $TrayUser = whoami } }
Write-Host "Root folder : $InstallDir`nTray user   : $TrayUser`nMode        : $(if ($UpdateOnly) { 'update binaries only' } else { 'full install' })"

# build
if (-not $SkipBuild) {
    Write-Host "Building Release..." -ForegroundColor Cyan
    & dotnet build "$repo\OneNoteWatcher.slnx" -c Release --nologo -v q
    if ($LASTEXITCODE -ne 0) { throw "dotnet build failed" }
}
$trayBin = "$repo\src\OneNoteWatcher\bin\Release\net8.0-windows"
$colBin  = "$repo\src\OneNoteWatcher.Collector\bin\Release\net8.0-windows"
foreach ($p in @($trayBin, $colBin)) { if (-not (Test-Path $p)) { throw "missing build output: $p" } }

# stop running copies
Get-ScheduledTask -TaskName $CollectorTask, $TrayTask -ErrorAction SilentlyContinue | Stop-ScheduledTask -ErrorAction SilentlyContinue
Get-Process OneNoteWatcher, OneNoteWatcher.Collector -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep 1

# copy binaries into the root; keep config.ini, status, logs
New-Item -ItemType Directory -Force -Path $InstallDir, (Join-Path $InstallDir "logs") | Out-Null
$cfg = Join-Path $InstallDir "config.ini"
Copy-Item "$trayBin\*" $InstallDir -Recurse -Force -Exclude "config.ini"
Copy-Item "$colBin\*"  $InstallDir -Recurse -Force -Exclude "config.ini"
if (-not (Test-Path $cfg)) { Copy-Item "$repo\config.ini" $cfg -Force; Write-Host "Installed config.ini" } else { Write-Host "Kept existing config.ini" }
# migrate an old Program Files install / old flat log files
$old = "$env:ProgramFiles\OneNoteWatcher"
if (Test-Path $old) { Remove-Item $old -Recurse -Force -ErrorAction SilentlyContinue; Write-Host "Removed old install at $old" }
Get-ChildItem $InstallDir -File -Include "sync-history.log*","collector.log" -ErrorAction SilentlyContinue | Move-Item -Destination (Join-Path $InstallDir "logs") -Force -ErrorAction SilentlyContinue

if (-not $UpdateOnly) {
    # Users can read/write the root (tray runs unelevated)
    $acl = Get-Acl $InstallDir
    $acl.SetAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule("Users", "Modify", "ContainerInherit,ObjectInherit", "None", "Allow")))
    Set-Acl $InstallDir $acl

    # the collector (SYSTEM) must read THIS user's OneNote index / MRU / diagnostic logs
    $sid = (New-Object System.Security.Principal.NTAccount($TrayUser)).Translate([System.Security.Principal.SecurityIdentifier]).Value
    $profileDir = (Get-CimInstance Win32_UserProfile -Filter "SID='$sid'").LocalPath
    if (-not $profileDir) { $profileDir = Join-Path "C:\Users" ($TrayUser.Split('\')[-1]) }
    Write-Host "User profile: $profileDir"

    $colExe = Join-Path $InstallDir "OneNoteWatcher.Collector.exe"
    $colAction = New-ScheduledTaskAction -Execute $colExe -Argument "--user-profile `"$profileDir`"" -WorkingDirectory $InstallDir
    $colTrig   = @((New-ScheduledTaskTrigger -AtStartup), (New-ScheduledTaskTrigger -AtLogOn))
    $colPrinc  = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
    $colSet    = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable `
                   -RestartCount 5 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit (New-TimeSpan -Days 3650) -MultipleInstances IgnoreNew
    Register-ScheduledTask -TaskName $CollectorTask -Action $colAction -Trigger $colTrig -Principal $colPrinc -Settings $colSet -Force | Out-Null

    $trayExe = Join-Path $InstallDir "OneNoteWatcher.exe"
    $trayAction = New-ScheduledTaskAction -Execute $trayExe -WorkingDirectory $InstallDir
    $trayTrig   = New-ScheduledTaskTrigger -AtLogOn -User $TrayUser
    $trayPrinc  = New-ScheduledTaskPrincipal -UserId $TrayUser -LogonType Interactive -RunLevel Limited
    $traySet    = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Days 3650) -MultipleInstances IgnoreNew
    Register-ScheduledTask -TaskName $TrayTask -Action $trayAction -Trigger $trayTrig -Principal $trayPrinc -Settings $traySet -Force | Out-Null
    Write-Host "Registered scheduled tasks."
}

Start-ScheduledTask -TaskName $CollectorTask
Start-ScheduledTask -TaskName $TrayTask
Start-Sleep 3
$colOk  = $null -ne (Get-Process OneNoteWatcher.Collector -ErrorAction SilentlyContinue)
$trayOk = $null -ne (Get-Process OneNoteWatcher -ErrorAction SilentlyContinue)
Write-Host ("`nCollector running: {0}    Tray running: {1}" -f $colOk, $trayOk) -ForegroundColor ($(if ($colOk -and $trayOk) { 'Green' } else { 'Yellow' }))
Write-Host "Everything is in $InstallDir  (logs in $InstallDir\logs, config: $cfg)"
