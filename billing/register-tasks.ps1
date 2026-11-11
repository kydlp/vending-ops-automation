# 月次の請求処理を Task Scheduler に登録する。
# 実運用で定時実行に載っているのは ingest（毎月2日 06:00）だけ。remind は現在は手動実行で、定時実行への切り替えを予定。draft は人が中身を見ながら手で流している。
# 実行: powershell -ExecutionPolicy Bypass -File .\register-tasks.ps1
# 解除: powershell -ExecutionPolicy Bypass -File .\register-tasks.ps1 -Remove
param([switch]$Remove)

$ErrorActionPreference = "Stop"
$here   = $PSScriptRoot
$python = (Get-Command python).Source   # PATH 上の Python を使う
$script = Join-Path $here "monthly_run.py"

# 日／時刻／サブコマンド。案B（5日はドラフト出力、発行は人が承認してから）
$tasks = @(
    @{ Name = "VMS\Invoice-Ingest"; Day = 2; Time = "06:00"; Cmd = "ingest";
       Desc = "前月分の売上を VMS から取得し、Drive のブックに追記する" },
    @{ Name = "VMS\Invoice-Remind"; Day = 4; Time = "09:00"; Cmd = "remind";
       Desc = "請求額を Discord に流して確認を促す" },
    @{ Name = "VMS\Invoice-Draft";  Day = 5; Time = "09:00"; Cmd = "draft";
       Desc = "検証を通してから請求書PDFをドラフト出力する（NGなら出力しない）" }
)

foreach ($t in $tasks) {
    if ($Remove) {
        schtasks /Delete /TN $t.Name /F 2>$null | Out-Null
        Write-Host ("削除: " + $t.Name)
        continue
    }

    $tr = '"{0}" "{1}" {2}' -f $python, $script, $t.Cmd
    schtasks /Create /TN $t.Name /TR $tr /SC MONTHLY /D $t.Day /ST $t.Time /F | Out-Null
    if ($LASTEXITCODE -ne 0) { throw ("登録に失敗: " + $t.Name) }

    # PCがスリープ・電源OFFで実行時刻を逃しても、次に起動したときに走らせる。
    # 月1回しか動かないので、取りこぼすと1ヶ月穴があく。
    $st = Get-ScheduledTask -TaskName ($t.Name -replace '^VMS\\', '') -TaskPath "\VMS\"
    $st.Settings.StartWhenAvailable = $true
    $st.Settings.ExecutionTimeLimit = "PT2H"
    $st.Description = $t.Desc
    Set-ScheduledTask -InputObject $st | Out-Null

    Write-Host ("登録: {0}  毎月{1}日 {2}  → monthly_run.py {3}" -f $t.Name, $t.Day, $t.Time, $t.Cmd)
}

if (-not $Remove) {
    Write-Host ""
    Write-Host "確認:   schtasks /Query /TN ""VMS\Invoice-Draft"" /FO LIST"
    Write-Host "手動実行: schtasks /Run /TN ""VMS\Invoice-Ingest"""
}
