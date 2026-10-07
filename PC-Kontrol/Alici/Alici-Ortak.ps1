function Save-GelenProblem([string]$Json, [string]$ReportRoot) {
    $report=$Json | ConvertFrom-Json -ErrorAction Stop
    $id=[Guid]::Empty
    if (-not [Guid]::TryParse([string]$report.id,[ref]$id)) { throw 'Invalid report id.' }
    if ([int]$report.schemaVersion -ne 2 -or $report.hasProblems -ne $true -or @($report.issues).Count -lt 1 -or @($report.issues).Count -gt 256) { throw 'Only schema v2 problem reports are accepted.' }
    if ([string]::IsNullOrWhiteSpace([string]$report.computerName) -or ([string]$report.computerName).Length -gt 128) { throw 'Invalid computer name.' }
    $time=[DateTimeOffset]::MinValue
    if (-not [DateTimeOffset]::TryParse([string]$report.occurredAtUtc,[ref]$time)) { throw 'Invalid report time.' }
    foreach ($issue in @($report.issues)) {
        if ([string]::IsNullOrWhiteSpace([string]$issue.component) -or [string]::IsNullOrWhiteSpace([string]$issue.kind) -or $null -eq $issue.detail) { throw 'Invalid issue entry.' }
    }
    New-Item -ItemType Directory -Path $ReportRoot -Force | Out-Null
    $base=$id.ToString('D')
    $path=Join-Path $ReportRoot ($base+'.json')
    $normal=$report | ConvertTo-Json -Depth 14
    $duplicate=Test-Path $path
    if ($duplicate) {
        $existing=[IO.File]::ReadAllText($path,[Text.Encoding]::UTF8)
        if ($existing -cne $normal) { throw 'Report id already exists with different content.' }
    } else {
        # The GUID is the only client-controlled part of the filename.
        $temporary=$path+'.tmp'
        [IO.File]::WriteAllText($temporary,$normal,(New-Object Text.UTF8Encoding($false)))
        [IO.File]::Move($temporary,$path)
    }
    function Encode([object]$Value) { [Net.WebUtility]::HtmlEncode([string]$Value) }
    $rows=@($report.issues | ForEach-Object {
        $detail=[string]$_.detail
        if($null -ne $_.driverError){$detail+='; Windows hata kodu: '+[string]$_.driverError}
        '<tr><td>'+ (Encode $_.component) + '</td><td>'+(Encode $_.kind)+'</td><td>'+(Encode $detail)+'</td></tr>'
    }) -join ''
    $date=$time.ToOffset([TimeSpan]::FromHours(3)).ToString('yyyy-MM-dd HH:mm:ss')+' (UTC+03)'
    $html='<!doctype html><html lang="tr"><meta charset="utf-8"><title>PC Kontrol - Sorun raporu</title><style>body{font:15px/1.6 Segoe UI,Arial;margin:40px;color:#182333}table{border-collapse:collapse;width:100%}td,th{border:1px solid #ccd3de;padding:10px;text-align:left;overflow-wrap:anywhere}pre{white-space:pre-wrap;overflow-wrap:anywhere;background:#f3f5f8;padding:16px}th{background:#edf2f7}</style><h1>PC Kontrol - Sorun raporu</h1><p><b>Bilgisayar:</b> '+(Encode $report.computerName)+'<br><b>Kontrol zamani:</b> '+(Encode $date)+'<br><b>Rapor:</b> '+(Encode $base)+'</p><table><tr><th>Bilesen</th><th>Sorun turu</th><th>Aciklama</th></tr>'+ $rows +'</table><p>check_unavailable: Test tamamlanamadi; fiziksel donanim arizasi oldugu anlamina gelmez.</p><details><summary>Tum kontrol ayrintilari</summary><pre>'+(Encode $normal)+'</pre></details></html>'
    [IO.File]::WriteAllText((Join-Path $ReportRoot ($base+'.html')),$html,(New-Object Text.UTF8Encoding($false)))
    return [pscustomobject]@{received=$true;id=[string]$report.id;duplicate=[bool]$duplicate}
}

function Write-AliciResponse($Context,[int]$Code,$Value) {
    $response=$Context.Response
    try {
        $body=[Text.Encoding]::UTF8.GetBytes(($Value | ConvertTo-Json -Depth 5 -Compress))
        $response.StatusCode=$Code;$response.ContentType='application/json; charset=utf-8';$response.ContentLength64=$body.Length
        $response.Headers['Cache-Control']='no-store';$response.Headers['X-Content-Type-Options']='nosniff'
        $response.OutputStream.Write($body,0,$body.Length)
    } finally { $response.Close() }
}

function Invoke-AliciRequest($Context,$Config,[string]$ReportRoot) {
    $request=$Context.Request
    if ($request.HttpMethod -ne 'POST') { Write-AliciResponse $Context 405 @{error='POST required'};return }
    if ($request.Url.AbsolutePath -cne '/pc-kontrol/rapor/') { Write-AliciResponse $Context 404 @{error='Not found'};return }
    if (-not [string]::Equals($request.Headers['X-PC-Kontrol-Key'],[string]$Config.SharedKey,[StringComparison]::Ordinal)) { Write-AliciResponse $Context 403 @{error='Invalid sender key'};return }
    if ($request.Headers['Origin']) { Write-AliciResponse $Context 403 @{error='Browser submissions not accepted'};return }
    if ($request.ContentType -notlike 'application/json*') { Write-AliciResponse $Context 415 @{error='JSON required'};return }
    if ($request.ContentLength64 -le 0 -or $request.ContentLength64 -gt 262144) { Write-AliciResponse $Context 413 @{error='Invalid report size'};return }
    try {
        $bytes=New-Object byte[] ([int]$request.ContentLength64)
        $offset=0;$clock=[Diagnostics.Stopwatch]::StartNew()
        while ($offset -lt $bytes.Length) {
            $remaining=5000-[int]$clock.ElapsedMilliseconds
            if ($remaining -le 0) { throw 'Body read timeout.' }
            $read=$request.InputStream.ReadAsync($bytes,$offset,$bytes.Length-$offset)
            if (-not $read.Wait($remaining)) { throw 'Body read timeout.' }
            $n=$read.Result
            if ($n -eq 0) { throw 'Incomplete request body.' }
            $offset+=$n
        }
        $json=(New-Object Text.UTF8Encoding($false,$true)).GetString($bytes)
        $ack=Save-GelenProblem $json $ReportRoot
    } catch {
        $message=$_.Exception.Message
        if ($message -like 'Report id already*') { Write-AliciResponse $Context 409 @{error='Report id collision'} }
        elseif ($_.Exception -is [IO.IOException] -or $_.Exception -is [UnauthorizedAccessException]) { Write-AliciResponse $Context 503 @{error='Report could not be saved'} }
        else { Write-AliciResponse $Context 400 @{error='Invalid report'} }
        return
    }
    Write-AliciResponse $Context 200 $ack
    if(-not $ack.duplicate -and $Config.AlarmEnabled -ne $false){
        try { $null=Start-KontrolAlarm -ReportRoot $ReportRoot -ReportId $ack.id -Config $Config }
        catch {
            # A UI failure must never turn a successfully saved report into an HTTP failure.
            $alarmErrors=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PC-Kontrol-Alici'
            New-Item -ItemType Directory -Path $alarmErrors -Force | Out-Null
            $_.Exception.Message | Set-Content (Join-Path $alarmErrors 'son-alarm-hatasi.txt') -Encoding UTF8
        }
    }
}

function Save-AlarmState([string]$Path,$Value){
    $temp=$Path+'.tmp'
    [IO.File]::WriteAllText($temp,($Value|ConvertTo-Json -Depth 6),(New-Object Text.UTF8Encoding($false)))
    if(Test-Path $Path){[IO.File]::Replace($temp,$Path,[NullString]::Value)}else{[IO.File]::Move($temp,$Path)}
}

function Start-KontrolAlarm([string]$ReportRoot,[string]$ReportId,$Config,[string]$StateRoot,[switch]$NoLaunch){
    $id=[Guid]::Empty
    if(-not [Guid]::TryParse($ReportId,[ref]$id)){throw 'Invalid alarm report id.'}
    if($Config.AlarmEnabled -eq $false){return}
    if(-not $StateRoot){$StateRoot=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PC-Kontrol-Alici'}
    New-Item -ItemType Directory -Path $StateRoot -Force | Out-Null
    $requestFile=Join-Path $StateRoot 'alarm-istek.json'
    $ackFile=Join-Path $StateRoot 'alarm-onay.json'
    $pending=@();$acknowledged=@()
    if(Test-Path $requestFile){$pending=@((Get-Content $requestFile -Raw|ConvertFrom-Json).reportIds)}
    if(Test-Path $ackFile){$acknowledged=@((Get-Content $ackFile -Raw|ConvertFrom-Json).reportIds)}
    $pending=@(@($pending)+$id.ToString('D')|Where-Object {$_ -and $_ -notin $acknowledged}|Select-Object -Unique|Select-Object -Last 200)
    Save-AlarmState $requestFile @{reportIds=$pending;reportRoot=$ReportRoot;soundEnabled=($Config.AlarmSoundEnabled -ne $false);updatedUtc=[DateTime]::UtcNow.ToString('o')}
    if($NoLaunch -or $env:OS -ne 'Windows_NT'){return}
    $recordFile=Join-Path $StateRoot 'alarm-process.json'
    if(Test-Path $recordFile){
        try{
            $saved=Get-Content $recordFile -Raw|ConvertFrom-Json
            $process=Get-Process -Id ([int]$saved.pid) -ErrorAction SilentlyContinue
            if($process -and $process.StartTime.ToUniversalTime().ToString('o') -eq $saved.startedAtUtc){return}
        }catch{}
    }
    $powershell=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $arguments='-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "'+(Join-Path $PSScriptRoot 'Alarm.ps1')+'" -StateRoot "'+$StateRoot+'"'
    $process=Start-Process $powershell -ArgumentList $arguments -PassThru
    Save-AlarmState $recordFile @{pid=$process.Id;startedAtUtc=$process.StartTime.ToUniversalTime().ToString('o')}
}
