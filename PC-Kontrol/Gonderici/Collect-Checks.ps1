param([Parameter(Mandatory=$true)][string]$OutFile)
$ErrorActionPreference='Stop'
try {
    Add-Type -Path @((Join-Path $PSScriptRoot 'app\Core.cs'), (Join-Path $PSScriptRoot 'app\NativeChecks.cs'))
    $cpu=[PcCheck.CpuRun]::new();$null=$cpu.Start()
    $clock=[Diagnostics.Stopwatch]::StartNew()
    while ($cpu.Active -and $clock.Elapsed.TotalSeconds -lt 10) { Start-Sleep -Milliseconds 50 }
    if ($cpu.Active) { $cpu.Cancel();$cpuResult=[pscustomobject]@{status='unknown';detail='CPU check timed out'} }
    else { $cpuResult=$cpu.ToJson() | ConvertFrom-Json }
    $gpu=[PcCheck.NativeChecks]::Gpu() | ConvertFrom-Json
    $audio=[PcCheck.NativeChecks]::AudioEndpoints() | ConvertFrom-Json
    $queryErrors=New-Object 'System.Collections.Generic.List[string]'
    $devices=@();$presence=[ordered]@{status='unavailable';keyboard=$null;mouse=$null}
    try {
        $present=@(Get-PnpDevice -PresentOnly -ErrorAction Stop | Where-Object { $_.Class -in @('Keyboard','Mouse','Display','MEDIA','AudioEndpoint','Processor') })
        $errors=@{}
        Get-CimInstance Win32_PnPEntity -OperationTimeoutSec 5 -ErrorAction Stop | ForEach-Object { $errors[$_.DeviceID]=$_.ConfigManagerErrorCode }
        $devices=@($present | ForEach-Object { [ordered]@{name=$_.FriendlyName;id=$_.InstanceId;class=$_.Class;status=$_.Status;driverError=$errors[$_.InstanceId]} })
        $presence=[ordered]@{status='complete';keyboard=@($present | Where-Object Class -eq 'Keyboard').Count;mouse=@($present | Where-Object Class -eq 'Mouse').Count}
    } catch { $queryErrors.Add('PnP enumeration: '+$_.Exception.Message) }
    $os=$null
    try { $os=Get-CimInstance Win32_OperatingSystem -OperationTimeoutSec 5 -ErrorAction Stop | Select-Object Caption,Version,BuildNumber }
    catch { $queryErrors.Add('OS query: '+$_.Exception.Message) }
    $snapshot=[ordered]@{status='complete';cpu=$cpuResult;gpu=$gpu;audio=$audio;presence=$presence;devices=$devices;os=$os;queryErrors=@($queryErrors.ToArray());physicalInputAndAcoustics='not_testable_unattended';temperature='not_measured'}
} catch {
    $snapshot=[ordered]@{status='unavailable';cpu=@{status='unknown'};gpu=@{status='unknown'};audio=@{status='unknown'};presence=@{status='unavailable'};devices=@();queryErrors=@($_.Exception.Message)}
}
[IO.File]::WriteAllText($OutFile,($snapshot | ConvertTo-Json -Depth 12 -Compress),(New-Object Text.UTF8Encoding($false)))
