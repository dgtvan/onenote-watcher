# ETW feasibility probe — MUST be run as Administrator.
# Read-only: starts a real-time trace on the Office logging providers, captures 60s while YOU
# make a small edit in OneNote, then decodes and greps for sync events. Writes a result file.
# It changes nothing on the system and deletes its own trace session when done.
#
# HOW TO RUN:
#   1. Press Start, type "PowerShell", right-click "Windows PowerShell" -> "Run as administrator".
#   2. Paste:  & "D:\Src\Personal\onenote-watcher\experiments\etw_probe.ps1"
#   3. When it says "EDIT ONENOTE NOW", type a few characters into any OneNote page (and Ctrl+S).
#   4. When it finishes it prints where the result file is; tell Claude it's done.

$ErrorActionPreference = 'Continue'
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole('Administrators')) {
    Write-Host "NOT ELEVATED. Re-open PowerShell as Administrator and run again." -ForegroundColor Red
    return
}

$dir  = Join-Path $env:LOCALAPPDATA 'Temp\onwatch_etwtest'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$out  = Join-Path $dir 'etw_result.txt'
$etl  = Join-Path $dir 'onwatch.etl'
$sess = 'onwatch_probe'

function Log($m){ Write-Host $m; $m | Out-File -FilePath $out -Append -Encoding utf8 }
"ETW probe $(Get-Date -Format o)  elevated=True" | Out-File $out -Encoding utf8

& logman stop $sess -ets 2>&1 | Out-Null
Remove-Item $etl -ErrorAction SilentlyContinue

$providers = @(
  '{F50D9315-E17E-43C1-8370-3EDF6CC057BE}',  # OfficeLoggingLiblet
  '{8736922D-E8B2-47EB-8564-23E77E728CF3}'   # Microsoft-Office-Events
)
Log ("create: " + (& logman create trace $sess -o $etl -ets -nb 16 256 -bs 1024 -mode Circular -max 128 2>&1))
foreach($p in $providers){
  Log ("enable $p : " + (& logman update trace $sess -p $p 0xffffffffffffffff 0xff -ets 2>&1))
}

Write-Host "`n>>> EDIT ONENOTE NOW: type a few characters in any page, then Ctrl+S. Capturing 60s...`n" -ForegroundColor Yellow
for($i=60;$i -gt 0;$i-=10){ Write-Host "  $i s left..."; Start-Sleep 10 }

Log ("stop: " + (& logman stop $sess -ets 2>&1))
Start-Sleep 2
if(-not (Test-Path $etl)){ Log "NO ETL PRODUCED"; return }
Log ("etl size: {0} bytes" -f (Get-Item $etl).Length)

$csv = "$etl.csv"; $sum = "$etl.summary.txt"
& tracerpt $etl -o $csv -summary $sum -f CSV -y 2>&1 | Out-Null

Log "`n=== event-count summary (top providers/events) ==="
if(Test-Path $sum){ Get-Content $sum | Select-Object -First 70 | ForEach-Object { Log $_ } }

Log "`n=== CSV matches for sync/storage/onenote ==="
if(Test-Path $csv){
  $total = (Get-Content $csv | Measure-Object -Line).Lines
  $hits = Select-String -Path $csv -Pattern 'Sync','Storage','OneNote','Notebook' -SimpleMatch -ErrorAction SilentlyContinue
  Log ("csv lines: {0}; matches: {1}" -f $total, $hits.Count)
  $hits | Select-Object -First 30 | ForEach-Object { Log ("  " + $_.Line.Substring(0,[Math]::Min(240,$_.Line.Length))) }
}
Log "`n=== DONE. Result file: $out ==="
Write-Host "`nRESULT FILE: $out" -ForegroundColor Green
