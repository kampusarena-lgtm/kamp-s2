param([switch]$NoRestart)
$ErrorActionPreference='Stop'
try{
    $localData=[Environment]::GetFolderPath('LocalApplicationData')
    $roles=@(@{name='Gonderici';target=(Join-Path $localData 'PC-Kontrol\Uygulama')},@{name='Alici';target=(Join-Path $localData 'PC-Kontrol-Alici\Uygulama')})
    $count=0
    foreach($role in $roles){
        $settings=Join-Path $role.target 'Ayarlar.json'
        if(-not (Test-Path $settings)){continue}
        $oldSettings=[IO.File]::ReadAllBytes($settings)
        $startupBytes=$null;$startupPath=$null
        if($role.name -eq 'Alici' -and -not $NoRestart){
            $startupPath=Join-Path ([Environment]::GetFolderPath('Startup')) 'PC Kontrol Alici.lnk'
            if(Test-Path $startupPath){$startupBytes=[IO.File]::ReadAllBytes($startupPath)}
            $stopScript=Join-Path $role.target 'Aliciyi-Durdur.ps1'
            if(Test-Path $stopScript){
                $powershell=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
                $stop=Start-Process $powershell -ArgumentList ('-NoProfile -ExecutionPolicy Bypass -File "'+$stopScript+'"') -PassThru -Wait -WindowStyle Hidden
                if($stop.ExitCode -ne 0){throw 'Calisan alici durdurulamadi.'}
            }
        }
        $source=Join-Path $PSScriptRoot $role.name
        Get-ChildItem $source | Where-Object Name -ne 'Ayarlar.json' | ForEach-Object {Copy-Item $_.FullName $role.target -Recurse -Force}
        if([Convert]::ToBase64String([IO.File]::ReadAllBytes($settings)) -cne [Convert]::ToBase64String($oldSettings)){throw 'Mevcut ayarlar beklenmedik sekilde degisti.'}
        if($startupBytes){[IO.File]::WriteAllBytes($startupPath,$startupBytes)}
        Write-Host ($role.name+' guncellendi. Anahtar ve ayarlar korundu.') -ForegroundColor Green
        $count++
        if(-not $NoRestart){
            $powershell=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
            $main=if($role.name -eq 'Alici'){'Alici.ps1'}else{'Kontrol.ps1'}
            $arguments='-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "'+(Join-Path $role.target $main)+'"'
            if($role.name -eq 'Gonderici'){$arguments+=' -NoDelay'}
            Start-Process $powershell -ArgumentList $arguments
        }
    }
    if($count -eq 0){throw 'Bu Windows kullanicisinda kurulu PC Kontrol bulunamadi. Ilk kurulum icin Eslesmeyi-Hazirla.cmd adimini kullan.'}
    Write-Host 'Guncelleme tamamlandi. Alicida Alarm-Test.cmd ile ses ve ekran uyarisini deneyebilirsin.'
}catch{Write-Host ('Guncelleme tamamlanamadi: '+$_.Exception.Message) -ForegroundColor Red;exit 1}
