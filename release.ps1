# WordWriter 快速发布脚本 (PowerShell)
# 用法: .\release.ps1 4.0.1 "Release notes"

param(
    [Parameter(Mandatory=$true)]
    [string]$Version,
    
    [Parameter(Mandatory=$false)]
    [string]$Message = "Release v$Version"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "WordWriter 发布脚本 v$Version" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# 1. 检查版本号格式
if ($Version -notmatch '^\d+\.\d+\.\d+$') {
    Write-Host "❌ 错误：版本号格式不正确，应为 X.Y.Z 格式" -ForegroundColor Red
    exit 1
}

Write-Host "`n步骤 1/6: 更新版本号..." -ForegroundColor Yellow

# 更新 setup.py
$setupContent = Get-Content "setup.py" -Raw
$setupContent = $setupContent -replace "version='[\d\.]+',", "version='$Version',"
Set-Content "setup.py" -Value $setupContent
Write-Host "  ✓ 更新 setup.py" -ForegroundColor Green

# 更新 pyproject.toml
$pyprojectContent = Get-Content "pyproject.toml" -Raw
$pyprojectContent = $pyprojectContent -replace 'version = "[\d\.]+"', "version = `"$Version`""
Set-Content "pyproject.toml" -Value $pyprojectContent
Write-Host "  ✓ 更新 pyproject.toml" -ForegroundColor Green

# 更新 __init__.py
$initContent = Get-Content "WordWriter\__init__.py" -Raw
$initContent = $initContent -replace "__version__ = '[\d\.]+'", "__version__ = '$Version'"
Set-Content "WordWriter\__init__.py" -Value $initContent
Write-Host "  ✓ 更新 WordWriter\__init__.py" -ForegroundColor Green

# 2. Git 状态检查
Write-Host "`n步骤 2/6: 检查 Git 状态..." -ForegroundColor Yellow
$gitStatus = git status --porcelain
if ($gitStatus) {
    Write-Host "  ℹ 发现未提交的更改" -ForegroundColor Cyan
} else {
    Write-Host "  ✓ 工作目录干净" -ForegroundColor Green
}

# 3. 提交更改
Write-Host "`n步骤 3/6: 提交更改到 Git..." -ForegroundColor Yellow
git add .
git commit -m "Release v$Version"
if ($LASTEXITCODE -eq 0) {
    Write-Host "  ✓ 提交成功" -ForegroundColor Green
} else {
    Write-Host "  ⚠ 没有新的更改需要提交" -ForegroundColor Yellow
}

# 4. 推送到远程
Write-Host "`n步骤 4/6: 推送到 GitHub..." -ForegroundColor Yellow
git push origin main
if ($LASTEXITCODE -eq 0) {
    Write-Host "  ✓ 推送成功" -ForegroundColor Green
} else {
    Write-Host "  ❌ 推送失败" -ForegroundColor Red
    exit 1
}

# 5. 创建标签
Write-Host "`n步骤 5/6: 创建 Git 标签..." -ForegroundColor Yellow
git tag "v$Version"
git push origin "v$Version"
if ($LASTEXITCODE -eq 0) {
    Write-Host "  ✓ 标签 v$Version 创建并推送成功" -ForegroundColor Green
} else {
    Write-Host "  ❌ 标签创建失败" -ForegroundColor Red
    exit 1
}

# 6. 提示创建 Release
Write-Host "`n步骤 6/6: 创建 GitHub Release..." -ForegroundColor Yellow
Write-Host "  ℹ 请访问以下链接创建 Release：" -ForegroundColor Cyan
Write-Host "  https://github.com/pzweuj/WordWriter/releases/new?tag=v$Version" -ForegroundColor Cyan
Write-Host ""
Write-Host "  或使用 GitHub CLI：" -ForegroundColor Cyan
Write-Host "  gh release create v$Version --title 'Release v$Version' --notes '$Message'" -ForegroundColor White

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "✅ 发布准备完成！" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "接下来：" -ForegroundColor Yellow
Write-Host "1. 在 GitHub 上创建 Release" -ForegroundColor White
Write-Host "2. GitHub Actions 将自动发布到 PyPI" -ForegroundColor White
Write-Host "3. 几分钟后检查 https://pypi.org/project/WordWriter/" -ForegroundColor White
Write-Host ""
