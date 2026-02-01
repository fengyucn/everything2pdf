#!/bin/bash
# Everything to PDF - Linux 打包脚本

set -e

echo "=== Everything to PDF 打包脚本 (Linux) ==="

# 检查PyInstaller
if ! command -v pyinstaller &> /dev/null; then
    echo "正在安装 PyInstaller..."
    pip install pyinstaller
fi

# 清理旧的构建
echo "清理旧的构建文件..."
rm -rf build dist

# 执行打包
echo "开始打包..."
pyinstaller build_linux.spec

# 检查结果
if [ -f "dist/everything2pdf" ]; then
    echo ""
    echo "=== 打包成功! ==="
    echo "可执行文件: dist/everything2pdf"
    ls -lh dist/everything2pdf
else
    echo "打包失败!"
    exit 1
fi
