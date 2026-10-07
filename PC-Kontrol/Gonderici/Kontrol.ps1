param([switch]$RetryOnly, [switch]$NoDelay)
$ErrorActionPreference='Stop'
$dataRoot=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PC-Kontrol'
$outbox=Join-Path $dataRoot 'Bekleyen-Raporlar'
$mutex=$null;$owns=$false;$worker=$null;$temporary=$null;$config=$null
try {
    if ($env:OS -ne 'Windows_NT') { throw 'This sender requires Windows 11.' }
    New-Item -ItemType Directory -Path $dataRoot -Force | Out-Null
    . (Join-Path $PSScriptRoot 'Ortak.ps1')
    $config=Read-KontrolConfig (Join-Path $PSScriptRoot 'Ayarlar.json')
    $mutex=New-Object Threading.Mutex($false,('Local\PC-Kontrol-Auto-'+([Environment]::UserName -replace '[^a-zA-Z0-9_-]','_')))
    try {$owns=$mutex.WaitOne(0)}catch [Threading.AbandonedMutexException]{$owns=$true}
    if (-not $owns) { return }
    if (-not $RetryOnly) {
        if (-not $NoDelay) { Start-Sleep -Seconds ([Math]::Min(60,[Math]::Max(0,[int]$config.StartupDelaySeconds))) }
        $temporary=Join-Path $dataRoot ('kontrol-'+[Guid]::NewGuid().ToString('N')+'.tmp')
        $powershell=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $arguments='-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "'+(Join-Path $PSScriptRoot 'Collect-Checks.ps1')+'" -OutFile "'+$temporary+'"'
        $worker=Start-Process $powershell -ArgumentList $arguments -PassThru -WindowStyle Hidden
        $limit=[Math]::Min(60,[Math]::Max(10,[int]$config.ChecksTimeoutSeconds))*1000
        if (-not $worker.WaitForExit($limit)) { $worker.Kill(); $snapshot=[pscustomobject]@{cpu=@{status='unknown'};gpu=@{status='unknown'};audio=@{status='unknown'};presence=@{status='unavailable'};devices=@();queryErrors=@('Unattended worker timed out.')} }
        elseif (-not (Test-Path $temporary)) { throw 'Unattended worker returned no check data.' }
        else { $snapshot=Get-Content $temporary -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop }
        $issues=@(Get-KontrolIssues $snapshot $config)
        $null=Add-KontrolProblem $outbox $snapshot $config ([Environment]::MachineName)
        Save-KontrolJson (Join-Path $dataRoot 'Son-Durum.json') ([ordered]@{checkedAtUtc=[DateTime]::UtcNow.ToString('o');problemCount=$issues.Count;hasProblems=($issues.Count -gt 0);note='Only automatically verifiable checks. Physical key/button/headset acoustics not assessed.'})
    }
    $null=Send-KontrolPending $outbox $config
} catch {
    try {
        New-Item -ItemType Directory -Path $dataRoot -Force | Out-Null
        [IO.File]::WriteAllText((Join-Path $dataRoot 'son-hata.txt'),[DateTime]::UtcNow.ToString('o')+' '+$_.Exception.Message,[Text.Encoding]::UTF8)
        if ($config) {
            . (Join-Path $PSScriptRoot 'Ortak.ps1')
            $snapshot=[pscustomobject]@{cpu=@{status='unknown'};gpu=@{status='unknown'};audio=@{status='unknown'};presence=@{status='unavailable'};devices=@();queryErrors=@('Check execution failed: '+$_.Exception.Message)}
            $null=Add-KontrolProblem $outbox $snapshot $config ([Environment]::MachineName)
            $null=Send-KontrolPending $outbox $config
        }
    } catch { } # Unattended mode must not open dialogs or block logon.
} finally {
    if ($worker) { try { if (-not $worker.HasExited) {$worker.Kill()} } catch { };$worker.Dispose() }
    if ($temporary -and (Test-Path $temporary)) { Remove-Item $temporary -Force -ErrorAction SilentlyContinue }
    if ($owns) {$mutex.ReleaseMutex()};if($mutex){$mutex.Dispose()}
}
