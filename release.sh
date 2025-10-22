#!/bin/bash
# WordWriter 快速发布脚本 (Bash)
# 用法: ./release.sh 4.0.1 "Release notes"

set -e  # 遇到错误立即退出

VERSION=$1
MESSAGE=${2:-"Release v$VERSION"}

if [ -z "$VERSION" ]; then
    echo "❌ 错误：请提供版本号"
    echo "用法: ./release.sh 4.0.1 \"Release notes\""
    exit 1
fi

# 检查版本号格式
if ! [[ $VERSION =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]]; then
    echo "❌ 错误：版本号格式不正确，应为 X.Y.Z 格式"
    exit 1
fi

echo "========================================"
echo "WordWriter 发布脚本 v$VERSION"
echo "========================================"

# 1. 更新版本号
echo ""
echo "步骤 1/6: 更新版本号..."

# 更新 setup.py
sed -i.bak "s/version='[0-9.]*',/version='$VERSION',/" setup.py
echo "  ✓ 更新 setup.py"

# 更新 pyproject.toml
sed -i.bak "s/version = \"[0-9.]*\"/version = \"$VERSION\"/" pyproject.toml
echo "  ✓ 更新 pyproject.toml"

# 更新 __init__.py
sed -i.bak "s/__version__ = '[0-9.]*'/__version__ = '$VERSION'/" WordWriter/__init__.py
echo "  ✓ 更新 WordWriter/__init__.py"

# 删除备份文件
rm -f setup.py.bak pyproject.toml.bak WordWriter/__init__.py.bak

# 2. Git 状态检查
echo ""
echo "步骤 2/6: 检查 Git 状态..."
if [[ -n $(git status --porcelain) ]]; then
    echo "  ℹ 发现未提交的更改"
else
    echo "  ✓ 工作目录干净"
fi

# 3. 提交更改
echo ""
echo "步骤 3/6: 提交更改到 Git..."
git add .
if git commit -m "Release v$VERSION"; then
    echo "  ✓ 提交成功"
else
    echo "  ⚠ 没有新的更改需要提交"
fi

# 4. 推送到远程
echo ""
echo "步骤 4/6: 推送到 GitHub..."
if git push origin main; then
    echo "  ✓ 推送成功"
else
    echo "  ❌ 推送失败"
    exit 1
fi

# 5. 创建标签
echo ""
echo "步骤 5/6: 创建 Git 标签..."
git tag "v$VERSION"
if git push origin "v$VERSION"; then
    echo "  ✓ 标签 v$VERSION 创建并推送成功"
else
    echo "  ❌ 标签创建失败"
    exit 1
fi

# 6. 提示创建 Release
echo ""
echo "步骤 6/6: 创建 GitHub Release..."
echo "  ℹ 请访问以下链接创建 Release："
echo "  https://github.com/pzweuj/WordWriter/releases/new?tag=v$VERSION"
echo ""
echo "  或使用 GitHub CLI："
echo "  gh release create v$VERSION --title 'Release v$VERSION' --notes '$MESSAGE'"

echo ""
echo "========================================"
echo "✅ 发布准备完成！"
echo "========================================"
echo ""
echo "接下来："
echo "1. 在 GitHub 上创建 Release"
echo "2. GitHub Actions 将自动发布到 PyPI"
echo "3. 几分钟后检查 https://pypi.org/project/WordWriter/"
echo ""
