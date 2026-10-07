$link=Join-Path ([Environment]::GetFolderPath('Startup')) 'PC Kontrol.lnk'
if(Test-Path $link){Remove-Item $link -Force}
Write-Host 'Otomatik kontrol kapatildi. Bekleyen raporlar ve uygulama silinmedi.'
