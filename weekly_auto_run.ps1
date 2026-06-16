$PYTHON = "C:\Program Files\PyManager\python.exe"
$BASE_DIR = "C:\Users\chiehyi\OneDrive - 大豐環保科技股份有限公司\文件\Claude\分析資料\週追蹤報告"
$LOG_FILE = "$BASE_DIR\weekly_auto_run.log"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"========================================" | Out-File $LOG_FILE -Encoding UTF8
"[$timestamp] Start" | Out-File $LOG_FILE -Append -Encoding UTF8
Set-Location $BASE_DIR
"[Step 1] generate_report.py ..." | Out-File $LOG_FILE -Append -Encoding UTF8
& $PYTHON "$BASE_DIR\generate_report.py" 2>&1 | Out-File $LOG_FILE -Append -Encoding UTF8
if ($LASTEXITCODE -ne 0) { "[ERROR] generate_report.py failed" | Out-File $LOG_FILE -Append; exit 1 }
"[Step 1] Done" | Out-File $LOG_FILE -Append -Encoding UTF8
"[Step 2] weekly_summary.py ..." | Out-File $LOG_FILE -Append -Encoding UTF8
& $PYTHON "$BASE_DIR\weekly_summary.py" 2>&1 | Out-File $LOG_FILE -Append -Encoding UTF8
if ($LASTEXITCODE -ne 0) { "[ERROR] weekly_summary.py failed" | Out-File $LOG_FILE -Append; exit 1 }
"[Step 2] Done" | Out-File $LOG_FILE -Append -Encoding UTF8
"[Step 3] hr_tracker.py ..." | Out-File $LOG_FILE -Append -Encoding UTF8
& $PYTHON "$BASE_DIR\hr_tracker.py" 2>&1 | Out-File $LOG_FILE -Append -Encoding UTF8
# hr_tracker 偶發 Outlook COM 斷線會失敗，但其產出非後續上傳所需，僅記 WARN 繼續（避免漏傳/漏 commit）
if ($LASTEXITCODE -ne 0) { "[WARN] hr_tracker.py failed (continue)" | Out-File $LOG_FILE -Append -Encoding UTF8 }
"[Step 3] Done" | Out-File $LOG_FILE -Append -Encoding UTF8
# [Step OD] SharePoint 上傳前確認 OneDrive 有在執行，沒在跑就自動啟動（否則檔案放進同步夾也不會上傳）
"[Step OD] Check OneDrive ..." | Out-File $LOG_FILE -Append -Encoding UTF8
if (-not (Get-Process OneDrive -ErrorAction SilentlyContinue)) {
    $odExe = "C:\Program Files\Microsoft OneDrive\OneDrive.exe"
    if (Test-Path $odExe) {
        Start-Process $odExe -ArgumentList "/background"
        Start-Sleep -Seconds 15
        "[Step OD] OneDrive 未執行，已自動啟動" | Out-File $LOG_FILE -Append -Encoding UTF8
    } else {
        "[WARN] 找不到 OneDrive.exe，略過自動啟動" | Out-File $LOG_FILE -Append -Encoding UTF8
    }
} else {
    "[Step OD] OneDrive 執行中" | Out-File $LOG_FILE -Append -Encoding UTF8
}
"[Step SP] sharepoint_archive.py ..." | Out-File $LOG_FILE -Append -Encoding UTF8
& $PYTHON "$BASE_DIR\sharepoint_archive.py" 2>&1 | Out-File $LOG_FILE -Append -Encoding UTF8
if ($LASTEXITCODE -ne 0) { "[WARN] sharepoint_archive.py failed (continue)" | Out-File $LOG_FILE -Append -Encoding UTF8 }
"[Step SP] Done" | Out-File $LOG_FILE -Append -Encoding UTF8
"[Step 4] git commit ..." | Out-File $LOG_FILE -Append -Encoding UTF8
$week = Get-Date -Format "MM/dd"
git -C $BASE_DIR add report.html weekly_dashboard.html hr_weekly_report.html 2>&1 | Out-File $LOG_FILE -Append -Encoding UTF8
git -C $BASE_DIR commit -m "chore: auto weekly report $week" 2>&1 | Out-File $LOG_FILE -Append -Encoding UTF8
"[Step 4] Done" | Out-File $LOG_FILE -Append -Encoding UTF8
$timestamp2 = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"[$timestamp2] All done!" | Out-File $LOG_FILE -Append -Encoding UTF8
Add-Type -AssemblyName System.Windows.Forms
$n = New-Object System.Windows.Forms.NotifyIcon
$n.Icon = [System.Drawing.SystemIcons]::Information
$n.Visible = $true
$n.ShowBalloonTip(8000, "Weekly Report Done", "Please review then git push.", [System.Windows.Forms.ToolTipIcon]::Info)
Start-Sleep -Seconds 3
$n.Dispose()