#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
打包脚本 - 使用 PyInstaller 将项目打包成可执行文件
"""

import os
import sys
import shutil
import subprocess
import argparse
from pathlib import Path

def check_pyinstaller():
    """检查 PyInstaller 是否已安装"""
    try:
        import PyInstaller
        print("[OK] PyInstaller 已安装")
        return True
    except ImportError:
        print("[WARN] PyInstaller 未安装")
        print("正在安装 PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("[OK] PyInstaller 安装完成")
        return True

def clean_build_dirs():
    """清理之前的构建目录"""
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            print(f"清理目录: {dir_name}")
            shutil.rmtree(dir_name)
    
    # 清理 .spec 文件（可选，如果需要重新生成）
    spec_file = 'merge_excel_pyside6.spec'
    if os.path.exists(spec_file):
        print(f"保留 spec 文件: {spec_file}")

def build_executable(enable_cleanup: bool = True):
    """
    构建可执行文件
    
    Args:
        enable_cleanup: 是否在打包后执行清理脚本，默认为 True
    """
    print("\n" + "=" * 60)
    print("开始打包...")
    print("=" * 60)
    
    # PyInstaller 命令参数
    cmd = [
        'pyinstaller',
        'merge_excel_pyside6.spec',
        '--clean',  # 清理临时文件
        '--noconfirm',  # 覆盖输出目录而不询问
    ]
    
    print(f"执行命令: {' '.join(cmd)}")
    result = subprocess.run(cmd, check=True)
    
    if result.returncode == 0:
        print("\n" + "=" * 60)
        print("[OK] 打包成功！")
        print("=" * 60)
        print(f"\n可执行文件位置: {os.path.abspath('dist')}")
        print("\n提示:")
        print("- 可执行文件在 dist 目录中")
        print("- 可以将整个 dist 目录分发给其他用户")
        print("- 首次运行可能需要几秒钟加载时间")
        
        # 根据参数决定是否执行清理脚本
        if enable_cleanup:
            cleanup_script = Path(__file__).parent / 'cleanup_deps.py'
            if cleanup_script.exists():
                print("\n" + "=" * 60)
                print("开始清理不需要的依赖文件...")
                print("=" * 60)
                try:
                    subprocess.run([sys.executable, str(cleanup_script)], check=False)
                except Exception as e:
                    print(f"[警告] 清理脚本执行失败: {e}")
                    print("您可以稍后手动运行 cleanup_deps.py")
            else:
                print("\n[提示] 清理脚本不存在，跳过清理步骤")
        else:
            print("\n[提示] 已跳过清理步骤（使用 --no-cleanup 参数）")
    else:
        print("\n[ERROR] 打包失败")
        sys.exit(1)

def main():
    """主函数"""
    # 解析命令行参数
    parser = argparse.ArgumentParser(
        description='Excel/CSV文件合并工具 - 打包脚本',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python build.py              # 打包并执行清理（默认）
  python build.py --no-cleanup  # 打包但不执行清理
  python build.py --cleanup     # 明确指定执行清理
        """
    )
    parser.add_argument(
        '--no-cleanup',
        action='store_true',
        help='打包后不执行清理脚本（默认会执行清理）'
    )
    parser.add_argument(
        '--cleanup',
        action='store_true',
        help='打包后执行清理脚本（默认行为）'
    )
    
    args = parser.parse_args()
    
    # 确定是否执行清理：如果指定了 --no-cleanup 则不清理，否则清理
    enable_cleanup = not args.no_cleanup
    
    print("=" * 60)
    print("Excel/CSV文件合并工具 - 打包脚本")
    print("=" * 60)
    if not enable_cleanup:
        print("[提示] 已禁用自动清理，打包完成后不会执行 cleanup_deps.py")
    
    # 检查当前目录
    if not os.path.exists('main.py'):
        print("错误: 请在项目根目录运行此脚本")
        sys.exit(1)
    
    # 检查 PyInstaller
    check_pyinstaller()
    
    # 清理构建目录
    print("\n清理构建目录...")
    clean_build_dirs()
    
    # 构建可执行文件
    build_executable(enable_cleanup=enable_cleanup)

if __name__ == "__main__":
    main()


