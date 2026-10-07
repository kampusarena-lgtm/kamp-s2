param([switch]$YeniAnahtar)
$ErrorActionPreference='Stop'
try {
    $paths=@((Join-Path $PSScriptRoot 'Gonderici\Ayarlar.json'),(Join-Path $PSScriptRoot 'Alici\Ayarlar.json'))
    $configs=@($paths | ForEach-Object {Get-Content $_ -Raw -Encoding UTF8 | ConvertFrom-Json})
    $keys=@($configs | ForEach-Object {[string]$_.SharedKey} | Where-Object {$_ -match '^[a-f0-9]{64}$'} | Select-Object -Unique)
    if(-not $YeniAnahtar -and $keys.Count -gt 1){throw 'Iki farkli anahtar var. Mevcut eslesmeyi korumak icin dosyalari kontrol et.'}
    if(-not $YeniAnahtar -and $keys.Count -eq 1){$key=$keys[0]}
    else {
        $random=[Security.Cryptography.RandomNumberGenerator]::Create()
        try{$bytes=New-Object byte[] 32;$random.GetBytes($bytes)}finally{$random.Dispose()}
        $key=[BitConverter]::ToString($bytes).Replace('-','').ToLowerInvariant()
    }
    for($i=0;$i -lt $configs.Count;$i++){
        $configs[$i].SharedKey=$key
        $temp=$paths[$i]+'.tmp'
        [IO.File]::WriteAllText($temp,($configs[$i]|ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($true)))
        [IO.File]::Replace($temp,$paths[$i],[NullString]::Value)
    }
    Write-Host 'Eslesme hazir. Bu Alici ve Gonderici klasorlerini ilgili bilgisayarlara kopyalayip kur.' -ForegroundColor Green
    Write-Host 'Hazirlanmis Ayarlar.json dosyalarini GitHub a yukleme; gonderim anahtari icerir.'
}catch{Write-Host ('Hazirlanamadi: '+$_.Exception.Message) -ForegroundColor Red;exit 1}
