param(
    [switch]$NoPush,
    [switch]$NoPull,
    [switch]$Quiet
)

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectDir

if (-not $Quiet) {
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "    公文助手 - 同步脚本" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
}

# Step 1: Pull latest
if (-not $NoPull) {
    if (-not $Quiet) { Write-Host "[1/4] 正在拉取远程更新..." -ForegroundColor Yellow }

    git pull --rebase 2>$null
    if ($LASTEXITCODE -ne 0) {
        git pull 2>$null
        if ($LASTEXITCODE -ne 0) {
            if (-not $Quiet) {
                Write-Host "[失败] 拉取失败！可能有冲突，请手动处理：" -ForegroundColor Red
                Write-Host "   cd $projectDir" -ForegroundColor Gray
                Write-Host "   git pull" - ForegroundColor Gray
                Write-Host "   解决冲突后运行 sync.ps1" -ForegroundColor Gray
            }
            exit 1
        }
    }
    if (-not $Quiet) { Write-Host "   OK - 拉取完成" -ForegroundColor Green }
}

# Step 2: Check for changes
$status = git status --porcelain
if (-not $status) {
    if (-not $Quiet) {
        Write-Host "[2/4] 没有需要提交的更改" -ForegroundColor Green
    }
    if (-not $NoPush) {
        git push 2>$null
        if ($LASTEXITCODE -eq 0) {
            if (-not $Quiet) { Write-Host "完成 - 一切最新！" -ForegroundColor Green }
        }
    }
    exit 0
}

# Step 3: Add and commit
if (-not $Quiet) { Write-Host "[2/4] 正在提交更改..." -ForegroundColor Yellow }

git add .
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"
git commit -m "auto sync $timestamp"

if (-not $Quiet) {
    $status -split "`n" | ForEach-Object {
        $f = $_ -replace '^.. ', ''
        if ($f) { Write-Host "   $f" -ForegroundColor Gray }
    }
}

# Step 4: Push
if (-not $NoPush) {
    if (-not $Quiet) { Write-Host "[3/4] 正在推送到 GitHub..." -ForegroundColor Yellow }

    git push 2>$null
    if ($LASTEXITCODE -eq 0) {
        $fileCount = ($status -split "`n").Count
        if (-not $Quiet) {
            Write-Host "[4/4] 完成！共同步 $fileCount 个文件" -ForegroundColor Green
            Write-Host "      提交时间: $timestamp" -ForegroundColor Gray
        }
    }
    else {
        if (-not $Quiet) {
            Write-Host "[失败] 推送失败，请检查网络连接" -ForegroundColor Red
        }
        exit 1
    }
}

if (-not $Quiet) {
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "    同步完成！" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
}
