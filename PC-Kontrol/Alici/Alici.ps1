param([string]$Prefix, [string]$ReportRoot, [string]$ConfigPath)
$ErrorActionPreference='Stop'
$data=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PC-Kontrol-Alici'
$listener=$null;$mutex=$null;$owns=$false
try {
    New-Item -ItemType Directory -Path $data -Force | Out-Null
    . (Join-Path $PSScriptRoot 'Alici-Ortak.ps1')
    if (-not $ConfigPath) { $ConfigPath=Join-Path $PSScriptRoot 'Ayarlar.json' }
    $config=Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([string]$config.SharedKey -notmatch '^[a-f0-9]{64}$') { throw 'Invalid sender key configuration.' }
    if ([int]$config.ReceiverPort -lt 1024 -or [int]$config.ReceiverPort -gt 65535) { throw 'Invalid receiver port.' }
    if (-not $Prefix) { $Prefix='http://+:'+([int]$config.ReceiverPort)+'/pc-kontrol/rapor/' }
    if (-not $ReportRoot) { $ReportRoot=Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PC-Kontrol-Gelen-Raporlar' }
    $mutex=New-Object Threading.Mutex($false,('Local\PC-Kontrol-Alici-'+$config.ReceiverPort))
    try {$owns=$mutex.WaitOne(0)}catch [Threading.AbandonedMutexException]{$owns=$true}
    if (-not $owns) { return }
    $listener=New-Object Net.HttpListener
    $listener.Prefixes.Add($Prefix);$listener.Start()
    [IO.File]::WriteAllText((Join-Path $data 'process.json'),(@{pid=$PID;startedAtUtc=(Get-Process -Id $PID).StartTime.ToUniversalTime().ToString('o')} | ConvertTo-Json))
    Write-Host ('PC Kontrol alicisi hazir: '+$Prefix) -ForegroundColor Green
    Write-Host ('Sorun raporlari: '+$ReportRoot)
    Write-Host 'Sorunsuz bilgisayarlar rapor gondermez. Durdurmak icin Ctrl+C.'
    while ($listener.IsListening) {
        $next=$listener.GetContextAsync()
        while (-not $next.IsCompleted) { Start-Sleep -Milliseconds 150 }
        try { Invoke-AliciRequest $next.Result $config $ReportRoot }
        catch { $_.Exception.Message | Set-Content (Join-Path $data 'son-hata.txt') -Encoding UTF8 }
    }
} catch {
    $_.Exception.Message | Set-Content (Join-Path $data 'son-hata.txt') -Encoding UTF8
    Write-Host ('Alici baslatilamadi: '+$_.Exception.Message) -ForegroundColor Red
    Write-Host 'Windows ta once Aliciyi-Kur.cmd dosyasini calistir.'
    exit 1
} finally {
    if ($listener) {$listener.Close()}
    if ($owns) { Remove-Item (Join-Path $data 'process.json') -Force -ErrorAction SilentlyContinue;$mutex.ReleaseMutex() }
    if ($mutex) {$mutex.Dispose()}
}
