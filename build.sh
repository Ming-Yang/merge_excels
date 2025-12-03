#!/usr/bin/env bash

set -euo pipefail

echo "============================================================"
echo "Excel/CSV文件合并工具 - Linux 打包脚本"
echo "============================================================"
echo

# 检查 Python 可执行文件
if command -v python3 >/dev/null 2>&1; then
    PYTHON_BIN="python3"
elif command -v python >/dev/null 2>&1; then
    PYTHON_BIN="python"
else
    echo "错误: 未找到 Python，请先安装 Python 3"
    exit 1
fi

echo "使用 Python 解释器: ${PYTHON_BIN}"

echo "检查并安装依赖..."
"${PYTHON_BIN}" -m pip install --upgrade pip
"${PYTHON_BIN}" -m pip install -r requirements.txt
"${PYTHON_BIN}" -m pip install pyinstaller

echo
echo "开始打包..."
echo "提示: 可以使用 --no-cleanup 参数跳过清理步骤"
"${PYTHON_BIN}" build.py "$@"

echo
echo "打包完成！构建产物位于 dist 目录"

