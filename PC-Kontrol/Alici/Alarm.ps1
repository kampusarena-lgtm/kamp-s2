param([Parameter(Mandatory=$true)][string]$StateRoot)
$ErrorActionPreference='Stop'
$alarmMutex=$null;$alarmOwns=$false;$alarmTimer=$null;$alarmForm=$null
$shownIds=@();$alarmRequest=$null
try{
    . (Join-Path $PSScriptRoot 'Alici-Ortak.ps1')
    $alarmMutex=New-Object Threading.Mutex($false,'Local\PC-Kontrol-Gorunur-Alarm')
    try{$alarmOwns=$alarmMutex.WaitOne(5000)}catch [Threading.AbandonedMutexException]{$alarmOwns=$true}
    if(-not $alarmOwns){return}
    Add-Type -AssemblyName System.Windows.Forms,System.Drawing
    [Windows.Forms.Application]::EnableVisualStyles()
    $alarmForm=New-Object Windows.Forms.Form
    $alarmForm.Text='PC Kontrol — Sorun alarmı';$alarmForm.Size=New-Object Drawing.Size(850,550)
    $alarmForm.StartPosition='CenterScreen';$alarmForm.TopMost=$true;$alarmForm.ShowInTaskbar=$true
    $alarmForm.Icon=[Drawing.SystemIcons]::Warning;$alarmForm.MinimumSize=New-Object Drawing.Size(650,400)
    $title=New-Object Windows.Forms.Label
    $title.Location=New-Object Drawing.Point(20,16);$title.Size=New-Object Drawing.Size(780,38)
    $title.Font=New-Object Drawing.Font('Segoe UI',18,[Drawing.FontStyle]::Bold);$title.ForeColor=[Drawing.Color]::DarkRed
    $title.Text='Bilgisayardan sorun raporu geldi';$alarmForm.Controls.Add($title)
    $sub=New-Object Windows.Forms.Label
    $sub.Location=New-Object Drawing.Point(22,58);$sub.Size=New-Object Drawing.Size(780,50)
    $sub.Font=New-Object Drawing.Font('Segoe UI',10);$sub.Text='Testin çalıştırılamaması, fiziksel arıza teşhisi değildir.';$alarmForm.Controls.Add($sub)
    $grid=New-Object Windows.Forms.DataGridView
    $grid.Location=New-Object Drawing.Point(20,112);$grid.Size=New-Object Drawing.Size(795,300)
    $grid.Anchor='Top, Bottom, Left, Right';$grid.ReadOnly=$true;$grid.AllowUserToAddRows=$false
    $grid.AllowUserToDeleteRows=$false;$grid.RowHeadersVisible=$false;$grid.AutoSizeRowsMode='AllCells'
    $grid.DefaultCellStyle.WrapMode='True';$grid.Font=New-Object Drawing.Font('Segoe UI',10)
    foreach($column in @(@('computer','Bilgisayar',145),@('component','Bileşen',90),@('kind','Sorun türü',150),@('detail','Açıklama',375))){
        $null=$grid.Columns.Add($column[0],$column[1]);$grid.Columns[$column[0]].Width=$column[2]
    }
    $alarmForm.Controls.Add($grid)
    $open=New-Object Windows.Forms.Button
    $open.Text='Rapor klasörünü aç';$open.Location=New-Object Drawing.Point(20,438);$open.Size=New-Object Drawing.Size(180,38);$open.Anchor='Bottom, Left'
    $open.Add_Click({if($script:alarmRequest){Start-Process explorer.exe -ArgumentList ('"'+$script:alarmRequest.reportRoot+'"')}});$alarmForm.Controls.Add($open)
    $ok=New-Object Windows.Forms.Button
    $ok.Text='Alarmı onayla ve kapat';$ok.Location=New-Object Drawing.Point(560,438);$ok.Size=New-Object Drawing.Size(255,38);$ok.Anchor='Bottom, Right'
    $ok.Add_Click({$alarmForm.Close()});$alarmForm.Controls.Add($ok)
    $script:shownIds=@();$script:lastSound=[DateTime]::MinValue
    function Refresh-KontrolAlarm {
        $file=Join-Path $StateRoot 'alarm-istek.json'
        if(-not (Test-Path $file)){return}
        $script:alarmRequest=Get-Content $file -Raw -Encoding UTF8|ConvertFrom-Json
        $ack=@();$ackFile=Join-Path $StateRoot 'alarm-onay.json'
        if(Test-Path $ackFile){$ack=@((Get-Content $ackFile -Raw|ConvertFrom-Json).reportIds)}
        $ids=@($script:alarmRequest.reportIds|Where-Object {$_ -notin $ack})
        if($ids.Count -eq 0){$alarmForm.Close();return}
        if(($ids -join ',') -cne ($script:shownIds -join ',')){
            $grid.Rows.Clear();$machines=New-Object 'System.Collections.Generic.List[string]';$validIds=@()
            foreach($idText in $ids){
                $parsed=[Guid]::Empty;if(-not [Guid]::TryParse([string]$idText,[ref]$parsed)){continue}
                $reportPath=Join-Path $script:alarmRequest.reportRoot ($parsed.ToString('D')+'.json')
                if(-not (Test-Path $reportPath)){continue}
                $report=Get-Content $reportPath -Raw -Encoding UTF8|ConvertFrom-Json
                if(-not $report.hasProblems){continue}
                $validIds+=$parsed.ToString('D');$machines.Add([string]$report.computerName)
                foreach($issue in $report.issues){
                    $detail=[string]$issue.detail;if($null -ne $issue.driverError){$detail+='; Windows kodu: '+$issue.driverError}
                    $null=$grid.Rows.Add([string]$report.computerName,[string]$issue.component,[string]$issue.kind,$detail)
                }
            }
            $script:shownIds=$validIds
            $sub.Text=($machines.ToArray()|Select-Object -Unique) -join ', '
            $sub.Text+=' — '+$validIds.Count+' rapor. check_unavailable = kontrol tamamlanamadı.'
            $script:lastSound=[DateTime]::MinValue
            $alarmForm.Activate()
        }
        if($script:shownIds.Count -gt 0 -and $script:alarmRequest.soundEnabled -and ([DateTime]::UtcNow-$script:lastSound).TotalSeconds -ge 15){
            [Media.SystemSounds]::Exclamation.Play()
            try{[Console]::Beep(950,180);[Console]::Beep(1250,180);[Console]::Beep(950,180)}catch{}
            $script:lastSound=[DateTime]::UtcNow
        }
    }
    $alarmForm.Add_FormClosing({
        $old=@();$ackFile=Join-Path $StateRoot 'alarm-onay.json'
        if(Test-Path $ackFile){$old=@((Get-Content $ackFile -Raw|ConvertFrom-Json).reportIds)}
        Save-AlarmState $ackFile @{reportIds=@(@($old)+@($script:shownIds)|Select-Object -Unique|Select-Object -Last 1000);acknowledgedUtc=[DateTime]::UtcNow.ToString('o')}
        Remove-Item (Join-Path $StateRoot 'alarm-process.json') -Force -ErrorAction SilentlyContinue
    })
    $alarmTimer=New-Object Windows.Forms.Timer;$alarmTimer.Interval=1000
    $alarmTimer.Add_Tick({try{Refresh-KontrolAlarm}catch{$_|Out-String|Set-Content (Join-Path $StateRoot 'son-alarm-hatasi.txt') -Encoding UTF8}})
    $alarmForm.Add_Shown({Refresh-KontrolAlarm;$alarmTimer.Start()})
    $null=$alarmForm.ShowDialog()
}catch{$_|Out-String|Set-Content (Join-Path $StateRoot 'son-alarm-hatasi.txt') -Encoding UTF8}
finally{
    if($alarmTimer){$alarmTimer.Stop();$alarmTimer.Dispose()};if($alarmForm){$alarmForm.Dispose()}
    if($alarmOwns){$alarmMutex.ReleaseMutex()};if($alarmMutex){$alarmMutex.Dispose()}
    # A report that arrived while the window was closing must not be lost.
    if($alarmOwns -and $alarmRequest){
        try{
            $latest=Get-Content (Join-Path $StateRoot 'alarm-istek.json') -Raw|ConvertFrom-Json
            $ack=(Get-Content (Join-Path $StateRoot 'alarm-onay.json') -Raw|ConvertFrom-Json).reportIds
            $unseen=@($latest.reportIds|Where-Object {$_ -notin @($ack)})
            if($unseen.Count -gt 0 -and $script:shownIds.Count -gt 0){
                $null=Start-KontrolAlarm -ReportRoot $latest.reportRoot -ReportId $unseen[-1] -Config @{AlarmEnabled=$true;AlarmSoundEnabled=$latest.soundEnabled} -StateRoot $StateRoot
            }
        }catch{}
    }
}
