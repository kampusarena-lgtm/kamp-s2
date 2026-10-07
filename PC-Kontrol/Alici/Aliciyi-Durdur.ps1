$ErrorActionPreference='Stop'
$root=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PC-Kontrol-Alici'
$startup=Join-Path ([Environment]::GetFolderPath('Startup')) 'PC Kontrol Alici.lnk'
if(Test-Path $startup){Remove-Item $startup -Force}
$record=Join-Path $root 'process.json'
if(Test-Path $record){
    $saved=Get-Content $record -Raw | ConvertFrom-Json
    $process=Get-Process -Id ([int]$saved.pid) -ErrorAction SilentlyContinue
    if($process -and $process.StartTime.ToUniversalTime().ToString('o') -eq $saved.startedAtUtc){
        $info=Get-CimInstance Win32_Process -Filter ('ProcessId='+[int]$saved.pid)
        if($info.CommandLine -like '*Alici.ps1*'){Stop-Process -Id ([int]$saved.pid)}
    }
    Remove-Item $record -Force -ErrorAction SilentlyContinue
}
Write-Host 'Alici durduruldu; otomatik baslatma kapatildi. Raporlar korundu.'
