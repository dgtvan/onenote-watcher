# Read-only probe of the Office diagnostic logs, the OTele store and the OAlerts event log.
# Reproduces the measurements in docs/detection-design.md. Run as the normal user; no admin.

$diag = Join-Path $env:LOCALAPPDATA 'Temp\Diagnostics\ONENOTE'
'=== 1. Which diagnostic logs can be opened while OneNote runs? ==='
Get-ChildItem $diag -Filter *.log | Sort-Object LastWriteTime | ForEach-Object {
    $r = 'OK'
    try {
        $fs = [IO.File]::Open($_.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        $fs.Close()
    } catch { $r = 'LOCKED' }
    '{0,-7} {1:yyyy-MM-dd HH:mm:ss}  {2}' -f $r, $_.LastWriteTime, $_.Name
}

'=== 2. Sync/error markers in the readable (rotated) logs ==='
$nul = [string][char]0
$text = Get-ChildItem $diag -Filter *.log | ForEach-Object {
    try { [IO.File]::ReadAllText($_.FullName).Replace($nul, '') } catch { '' }
} | Out-String
foreach ($p in 'Error.Code', 'Error.Type', 'jerrc', 'NotebookSyncResult', 'FSSHTTP') {
    '{0,-20} {1}' -f $p, ([regex]::Matches($text, [regex]::Escape($p))).Count
}
'readable bytes total: {0}' -f $text.Length

'=== 3. OTele store: live row count (expect 0 unless upload is failing) ==='
$otele = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\OTele'
Get-ChildItem $otele -Filter 'onenote.exe.db*' | ForEach-Object { '{0,10} B  {1:HH:mm:ss}  {2}' -f $_.Length, $_.LastWriteTime, $_.Name }
'(decode history with: python experiments\bond_decoder.py)'

'=== 4. OAlerts: OneNote dialog alerts ==='
Get-WinEvent -LogName OAlerts -MaxEvents 300 -ErrorAction SilentlyContinue |
    Where-Object { $_.Message -like 'Microsoft OneNote*' } |
    Select-Object -First 10 TimeCreated, @{n = 'Message'; e = { ($_.Message -replace "`r?`n", ' ').Substring(0, [Math]::Min(160, $_.Message.Length)) } } |
    Format-Table -Wrap | Out-String -Width 200

'=== 5. Privileges ==='
'admin token: {0}' -f (([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole('Administrators'))
'Performance Log Users: {0}' -f ((whoami /groups) -match 'S-1-5-32-559').Count
