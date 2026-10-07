$ErrorActionPreference='Stop'
try{
    . (Join-Path $PSScriptRoot 'Alici-Ortak.ps1')
    $data=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PC-Kontrol-Alici\Alarm-Test'
    $reportId=[Guid]::NewGuid().ToString('D')
    $report=@{schemaVersion=2;appVersion='1.2.0';id=$reportId;computerName='ALARM TESTİ';occurredAtUtc=[DateTime]::UtcNow.ToString('o');hasProblems=$true;issues=@(@{component='alarm';kind='test';detail='Bu yalnızca yerel ses ve ekran alarmı denemesidir; donanım arızası raporu değildir.'})}
    $null=Save-GelenProblem ($report|ConvertTo-Json -Depth 8) $data
    $null=Start-KontrolAlarm -ReportRoot $data -ReportId $reportId -Config @{AlarmEnabled=$true;AlarmSoundEnabled=$true}
    Write-Host 'Alarm test penceresi açıldı. Ses ve görünür uyarıyı kontrol et.'
}catch{Write-Host $_.Exception.Message -ForegroundColor Red;exit 1}
