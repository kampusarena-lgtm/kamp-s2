param([switch]$FirewallOnly,[string]$TargetUser)
$ErrorActionPreference='Stop'
try {
    $config=Get-Content (Join-Path $PSScriptRoot 'Ayarlar.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $port=[int]$config.ReceiverPort
    if($port -lt 1024 -or $port -gt 65535){throw 'Invalid port.'}
    if($FirewallOnly){
        $admin=New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
        if(-not $admin.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)){throw 'Administrator permission is required for the receiver firewall and URL reservation.'}
        if(-not $TargetUser){throw 'Target receiver user is required.'}
        $prefix='http://+:'+$port+'/pc-kontrol/rapor/'
        & netsh http delete urlacl ('url='+$prefix) 2>$null | Out-Null
        & netsh http add urlacl ('url='+$prefix) ('user='+$TargetUser) 'listen=yes' | Out-Null
        if($LASTEXITCODE -ne 0){throw 'Could not reserve HTTP URL.'}
        $rule='PC-Kontrol-Alici-'+$port
        Get-NetFirewallRule -Name $rule -ErrorAction SilentlyContinue | Remove-NetFirewallRule
        New-NetFirewallRule -Name $rule -DisplayName ('PC Kontrol rapor alicisi '+$port) -Direction Inbound -Action Allow -Protocol TCP -LocalPort $port -RemoteAddress LocalSubnet -Profile Private,Domain | Out-Null
        exit 0
    }
    $target=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PC-Kontrol-Alici\Uygulama'
    New-Item -ItemType Directory -Path $target -Force | Out-Null
    if([IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\') -ne [IO.Path]::GetFullPath($target).TrimEnd('\')){
        Get-ChildItem $PSScriptRoot | ForEach-Object {Copy-Item $_.FullName $target -Recurse -Force}
    }
    $powershell=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $who=[Security.Principal.WindowsIdentity]::GetCurrent().Name
    $elevate='-NoProfile -ExecutionPolicy Bypass -File "'+(Join-Path $target 'Alici-Kur.ps1')+'" -FirewallOnly -TargetUser "'+$who+'"'
    Write-Host 'Alim icin TCP portu ve HTTP adresi kaydedilecek. Windows yonetici izni isteyecek.'
    $setup=Start-Process $powershell -Verb RunAs -ArgumentList $elevate -PassThru -Wait
    if($setup.ExitCode -ne 0){throw 'Receiver network setup did not complete.'}
    $shell=New-Object -ComObject WScript.Shell
    $args='-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "'+(Join-Path $target 'Alici.ps1')+'"'
    $link=$shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Startup')) 'PC Kontrol Alici.lnk'))
    $link.TargetPath=$powershell;$link.Arguments=$args;$link.WorkingDirectory=$target;$link.WindowStyle=7;$link.Save()
    Write-Host ('Alici kuruldu. TCP '+$port+' dinleniyor; gelen sorunlar Belgeler\PC-Kontrol-Gelen-Raporlar klasorune kaydedilir.') -ForegroundColor Green
    Start-Process $powershell -ArgumentList $args
}catch{Write-Host ('Kurulum tamamlanamadi: '+$_.Exception.Message) -ForegroundColor Red;exit 1}
