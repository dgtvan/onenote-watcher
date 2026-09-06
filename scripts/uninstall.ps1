# OneNote Sync Watcher — single-script UNINSTALL. Run as Administrator.
#
#   & "D:\Src\Personal\onenote-watcher\scripts\uninstall.ps1"             # remove app + auto-start, keep logs\
#   & "D:\Src\Personal\onenote-watcher\scripts\uninstall.ps1" -PurgeData  # remove the whole root folder incl. logs
param(
    [string]$InstallDir = "C:\ProgramData\OneNoteWatcher",
    [switch]$PurgeData
)
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole('Administrators')) {
    Write-Host "NOT ELEVATED. Re-open PowerShell as Administrator and run again." -ForegroundColor Red; return
}
$tasks = "OneNoteWatcher Collector", "OneNoteWatcher Tray"
Get-ScheduledTask -TaskName $tasks -ErrorAction SilentlyContinue | Stop-ScheduledTask -ErrorAction SilentlyContinue
Get-Process OneNoteWatcher, OneNoteWatcher.Collector -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
& logman stop "OneNoteWatcher_Sync" -ets 2>$null | Out-Null
foreach ($t in $tasks) { if (Get-ScheduledTask -TaskName $t -ErrorAction SilentlyContinue) { Unregister-ScheduledTask -TaskName $t -Confirm:$false; Write-Host "Removed task '$t'" } }
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "OneNoteWatcher" -ErrorAction SilentlyContinue
$old = "$env:ProgramFiles\OneNoteWatcher"; if (Test-Path $old) { Remove-Item $old -Recurse -Force; Write-Host "Removed $old" }

if (Test-Path $InstallDir) {
    if ($PurgeData) { Remove-Item $InstallDir -Recurse -Force; Write-Host "Removed $InstallDir (incl. logs)" }
    else {
        Get-ChildItem $InstallDir -Force | Where-Object { $_.Name -ne "logs" } | Remove-Item -Recurse -Force
        Write-Host "Removed app from $InstallDir; kept $InstallDir\logs (re-run with -PurgeData to delete)"
    }
}
Remove-Item (Join-Path $env:LOCALAPPDATA "OneNoteWatcher") -Recurse -Force -ErrorAction SilentlyContinue   # Graph token cache
Write-Host "Uninstall complete." -ForegroundColor Green
