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
if ($LASTEXITCODE -ne 0) { "[ERROR] hr_tracker.py failed" | Out-File $LOG_FILE -Append; exit 1 }
"[Step 3] Done" | Out-File $LOG_FILE -Append -Encoding UTF8
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