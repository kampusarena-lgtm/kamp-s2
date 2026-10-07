$ErrorActionPreference='Stop'
try {
    $target=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'PC-Kontrol\Uygulama'
    New-Item -ItemType Directory -Path $target -Force | Out-Null
    if ([IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\') -ne [IO.Path]::GetFullPath($target).TrimEnd('\')) {
        Get-ChildItem $PSScriptRoot | ForEach-Object { Copy-Item $_.FullName $target -Recurse -Force }
    }
    $shell=New-Object -ComObject WScript.Shell
    $powershell=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $args='-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "'+(Join-Path $target 'Kontrol.ps1')+'"'
    foreach ($folder in @([Environment]::GetFolderPath('Startup'),[Environment]::GetFolderPath('Desktop'))) {
        $link=$shell.CreateShortcut((Join-Path $folder 'PC Kontrol.lnk'))
        $link.TargetPath=$powershell;$link.Arguments=$args;$link.WorkingDirectory=$target;$link.WindowStyle=7
        $link.Description='Katilimsiz kontrol; sadece sorun varsa HTTP raporu';$link.Save()
    }
    Write-Host 'Kuruldu. Oturum acilinca sessiz kontrol yapilacak. Sorun yoksa rapor gonderilmez.' -ForegroundColor Green
    Start-Process $powershell -ArgumentList $args
} catch {Write-Host $_.Exception.Message -ForegroundColor Red;exit 1}
