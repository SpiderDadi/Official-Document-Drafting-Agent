Add-Type -AssemblyName System.Windows.Forms
$notify = New-Object System.Windows.Forms.NotifyIcon
$notify.Icon = [System.Drawing.SystemIcons]::Information
$notify.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
$notify.BalloonTipTitle = "📋 每日政策简报已生成"
$notify.BalloonTipText = "哥哥，今天16:00到啦！政策简报已自动搜索完成，快打开Claude看看，选一条我做深度报告和PPT吧~"
$notify.Visible = $true
$notify.ShowBalloonTip(15000)
Start-Sleep 10
$notify.Dispose()
