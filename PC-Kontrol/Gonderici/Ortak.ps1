function Read-KontrolConfig([string]$Path) {
    $config = Get-Content $Path -Raw -Encoding UTF8 | ConvertFrom-Json -ErrorAction Stop
    $url = [Uri]$config.ReceiverUrl
    if (-not $url.IsAbsoluteUri -or $url.Scheme -notin @('http','https')) { throw 'ReceiverUrl must be an absolute HTTP URL.' }
    if ([string]$config.SharedKey -notmatch '^[a-f0-9]{64}$') { throw 'SharedKey must contain 64 hex characters.' }
    if ([int]$config.TimeoutSeconds -lt 1 -or [int]$config.TimeoutSeconds -gt 30) { throw 'TimeoutSeconds must be between 1 and 30.' }
    if ([int]$config.MaxPendingPerRun -lt 1 -or [int]$config.MaxPendingPerRun -gt 50) { throw 'MaxPendingPerRun must be between 1 and 50.' }
    return $config
}

function Get-InputPresence($PnpDevices,$RawInput){
    $keyboard=@($PnpDevices|Where-Object Class -eq 'Keyboard').Count
    $mouse=@($PnpDevices|Where-Object Class -eq 'Mouse').Count
    $presence=[ordered]@{status='complete';keyboard=$keyboard;mouse=$mouse;rawInput=$RawInput;pnpKeyboard=$keyboard;pnpMouse=$mouse}
    if($RawInput.status -eq 'complete'){
        $presence.keyboard=[Math]::Max($keyboard,[int]$RawInput.keyboard)
        $presence.mouse=[Math]::Max($mouse,[int]$RawInput.mouse)
    }
    return [pscustomobject]$presence
}

function Get-KontrolIssues($Snapshot, $Config) {
    $issues = New-Object 'System.Collections.Generic.List[object]'
    foreach ($key in @('cpu','gpu')) {
        $test = $Snapshot.$key
        if ($test.status -eq 'fail') { $issues.Add([ordered]@{component=$key;kind='test_failure';detail=($test | ConvertTo-Json -Depth 8 -Compress)}) }
        elseif ($test.status -ne 'pass') { $issues.Add([ordered]@{component=$key;kind='check_unavailable';detail=($test | ConvertTo-Json -Depth 8 -Compress)}) }
    }
    foreach ($errorText in @($Snapshot.queryErrors)) {
        if ($errorText) { $issues.Add([ordered]@{component='system';kind='check_unavailable';detail=[string]$errorText}) }
    }
    foreach ($device in @($Snapshot.devices)) {
        # A deliberately disabled optional HDMI/DP audio adapter is not a failure
        # if Windows has another active default audio output. Other error codes remain alarms.
        if($device.class -eq 'MEDIA' -and $null -ne $device.driverError -and [int]$device.driverError -eq 22 -and
           $Snapshot.audio.status -eq 'complete' -and $Snapshot.audio.defaultOutputAvailable -and
           $Config.IgnoreDisabledOptionalAudio -ne $false){continue}
        $driverError = $null -ne $device.driverError -and [int]$device.driverError -ne 0
        if ($driverError -or $device.status -in @('Error','Degraded')) {
            $issues.Add([ordered]@{component=[string]$device.class;kind='driver_error';detail=([string]$device.name+'; Windows error code: '+[string]$device.driverError+'; status: '+[string]$device.status);instanceId=[string]$device.id;driverError=$device.driverError;deviceStatus=$device.status})
        }
    }
    if ($Snapshot.presence.status -eq 'complete') {
        foreach ($item in @(@{key='keyboard';required=$Config.RequiredDevices.Keyboard}, @{key='mouse';required=$Config.RequiredDevices.Mouse})) {
            if ($item.required -and [int]$Snapshot.presence.($item.key) -eq 0) {
                $raw=$Snapshot.presence.rawInput
                if($raw -and ($raw.status -ne 'complete' -or $raw.remoteSession)){
                    $issues.Add([ordered]@{component=$item.key;kind='check_unavailable';detail='PnP did not find this input device; Raw Input could not confirm physical presence, or the check ran in a remote session.'})
                }else{
                    $issues.Add([ordered]@{component=$item.key;kind='required_device_missing';detail='No device found by available PnP/Raw Input presence checks.'})
                }
            }
        }
    }
    if ($Snapshot.audio.status -eq 'complete') {
        if ($Config.RequiredDevices.AudioOutput -and -not $Snapshot.audio.defaultOutputAvailable) {
            $issues.Add([ordered]@{component='audio';kind='required_device_missing';detail='No active default Windows audio output endpoint.'})
        }
    } elseif ($Config.RequiredDevices.AudioOutput) {
        $issues.Add([ordered]@{component='audio';kind='check_unavailable';detail=($Snapshot.audio | ConvertTo-Json -Depth 6 -Compress)})
    }
    if ($Config.ExpectedAudioNamePattern -and $Snapshot.presence.status -eq 'complete') {
        $found = @($Snapshot.devices | Where-Object { $_.class -eq 'AudioEndpoint' -and $_.name -like $Config.ExpectedAudioNamePattern })
        if ($found.Count -eq 0) { $issues.Add([ordered]@{component='headset';kind='expected_device_not_enumerated';detail='Expected AudioEndpoint name pattern was not found in present Windows devices. This is not an acoustic test.'}) }
    }
    return $issues.ToArray()
}

function Save-KontrolJson([string]$Path, $Value) {
    $json = $Value | ConvertTo-Json -Depth 14
    $temp = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
    [IO.File]::WriteAllText($temp, $json, (New-Object Text.UTF8Encoding($false)))
    if (Test-Path $Path) { [IO.File]::Replace($temp, $Path, [NullString]::Value) } else { [IO.File]::Move($temp, $Path) }
}

function Add-KontrolProblem([string]$Outbox, $Snapshot, $Config, [string]$ComputerName) {
    $issues = @(Get-KontrolIssues $Snapshot $Config)
    if ($issues.Count -eq 0) { return $null }
    New-Item -ItemType Directory -Path $Outbox -Force | Out-Null
    $id = [Guid]::NewGuid().ToString('D')
    $report = [ordered]@{
        schemaVersion=2; appVersion='1.2.0'; id=$id; computerName=$ComputerName
        occurredAtUtc=[DateTime]::UtcNow.ToString('o'); hasProblems=$true; issues=$issues
        checks=$Snapshot
        limitations=@('No physical keyboard switch, mouse button or headset acoustic verification without user input.','Short CPU and default D3D11 adapter texture checks; no temperature or exhaustive stress test.')
    }
    $path = Join-Path $Outbox ($id + '.json')
    Save-KontrolJson $path $report
    return $path
}

function Send-KontrolPending([string]$Outbox, $Config) {
    $sent=0; $failed=0
    if (-not (Test-Path $Outbox)) { return [pscustomobject]@{sent=0;failed=0} }
    $files = @(Get-ChildItem $Outbox -Filter '*.json' -File | Sort-Object CreationTimeUtc | Select-Object -First ([int]$Config.MaxPendingPerRun))
    foreach ($file in $files) {
        try {
            $json = [IO.File]::ReadAllText($file.FullName, [Text.Encoding]::UTF8)
            $report = $json | ConvertFrom-Json -ErrorAction Stop
            if ($report.hasProblems -ne $true -or @($report.issues).Count -eq 0 -or [string]$report.id -notmatch '^[a-fA-F0-9-]{36}$') { throw 'Invalid queued problem report.' }
            $request = [Net.HttpWebRequest]::Create([Uri]$Config.ReceiverUrl)
            $request.Method='POST'; $request.ContentType='application/json; charset=utf-8'
            $request.Headers['X-PC-Kontrol-Key']=[string]$Config.SharedKey
            $request.Timeout=[int]$Config.TimeoutSeconds*1000; $request.ReadWriteTimeout=$request.Timeout
            $request.Proxy=$null; $request.AllowAutoRedirect=$false; $request.KeepAlive=$false
            $bytes=[Text.Encoding]::UTF8.GetBytes($json);$request.ContentLength=$bytes.Length
            $stream=$request.GetRequestStream()
            try { $stream.Write($bytes,0,$bytes.Length) } finally { $stream.Dispose() }
            $response=$request.GetResponse()
            try {
                if ([int]$response.StatusCode -ne 200) { throw 'Receiver did not acknowledge report.' }
                $reader=New-Object IO.StreamReader($response.GetResponseStream())
                try { $ack=$reader.ReadToEnd() | ConvertFrom-Json -ErrorAction Stop } finally { $reader.Dispose() }
                if ($ack.received -ne $true -or [string]$ack.id -cne [string]$report.id) { throw 'Receiver acknowledgement did not match the queued report.' }
            } finally { $response.Dispose() }
            Remove-Item $file.FullName -Force
            $sent++
        } catch {
            $failed++
            $errorFile=Join-Path $Outbox 'son-gonderim-hatasi.txt'
            [IO.File]::WriteAllText($errorFile, [DateTime]::UtcNow.ToString('o') + ' ' + $_.Exception.Message, [Text.Encoding]::UTF8)
            break # Receiver unavailable: keep all reports; avoid repeating identical connection failures.
        }
    }
    return [pscustomobject]@{sent=$sent;failed=$failed}
}
