#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
清理脚本 - 删除打包后不需要的依赖文件
在打包完成后调用此脚本，可以删除 dist/Excel合并工具/_internal 中不需要的文件
"""

import os
import sys
import shutil
from pathlib import Path


# 要删除的文件/目录列表（相对于 dist/Excel合并工具/_internal 的路径）
# 用户可以在这里添加要删除的文件路径
FILES_TO_DELETE = [
    'sqlite3.dll',
    'libssl-1_1.dll',
    '_ssl.pyd',
    '_sqlite3.pyd',
    'PySide6/opengl32sw.dll',
    'PySide6/Qt6Quick.dll',
    'PySide6/Qt6Pdf.dll',
    'PySide6/Qt6Qml.dll',
    'PySide6/Qt6Network.dll',
    'PySide6/QtNetwork.pyd',
    'PySide6/Qt6OpenGL.dll',
    'PySide6/Qt6QmlModels.dll',
    'PySide6/Qt6VirtualKeyboard.dll',
    'PySide6/translations/',
    'PIL',
    'numpy/fft',
]


def delete_file_or_dir(base_path: Path, relative_path: str) -> bool:
    """
    删除文件或目录（如果是目录，会递归删除整个文件夹及其所有内容）
    
    Args:
        base_path: 基础路径 (dist/Excel合并工具/_internal)
        relative_path: 相对路径（会自动规范化，去除末尾斜杠）
        
    Returns:
        bool: 是否成功删除
    """
    # 规范化路径：去除末尾的斜杠和反斜杠，统一处理
    normalized_path = relative_path.rstrip('/\\')
    target_path = base_path / normalized_path
    
    if not target_path.exists():
        print(f"  [跳过] 路径不存在: {normalized_path}")
        return False
    
    try:
        if target_path.is_file():
            # 删除单个文件
            target_path.unlink()
            print(f"  [删除] 文件: {normalized_path}")
        elif target_path.is_dir():
            # 递归删除整个目录及其所有内容
            shutil.rmtree(target_path)
            print(f"  [删除] 目录（已完全删除）: {normalized_path}")
        else:
            # 其他类型（如符号链接等）
            target_path.unlink()
            print(f"  [删除] 其他类型: {normalized_path}")
        return True
    except Exception as e:
        print(f"  [错误] 删除失败 {normalized_path}: {e}")
        return False


def cleanup_dependencies(files_to_delete: list = None):
    """
    清理依赖文件
    
    Args:
        files_to_delete: 要删除的文件列表，如果为 None 则使用默认列表
    """
    if files_to_delete is None:
        files_to_delete = FILES_TO_DELETE
    
    # 确定基础路径
    project_root = Path(__file__).parent
    base_path = project_root / 'dist' / 'Excel合并工具' / '_internal'
    
    if not base_path.exists():
        print(f"[错误] 目录不存在: {base_path}")
        print("请先运行打包脚本生成 dist 目录")
        return False
    
    print("=" * 60)
    print("清理依赖文件")
    print("=" * 60)
    print(f"基础路径: {base_path}")
    print(f"要删除的文件数量: {len(files_to_delete)}")
    print()
    
    if not files_to_delete:
        print("[提示] 没有要删除的文件")
        print("请在脚本的 FILES_TO_DELETE 列表中添加要删除的文件路径")
        return True
    
    deleted_count = 0
    skipped_count = 0
    
    for relative_path in files_to_delete:
        if delete_file_or_dir(base_path, relative_path):
            deleted_count += 1
        else:
            skipped_count += 1
    
    print()
    print("=" * 60)
    print(f"清理完成: 删除 {deleted_count} 个，跳过 {skipped_count} 个")
    print("=" * 60)
    
    return True


def main():
    """主函数"""
    # 可以从命令行参数读取文件列表
    if len(sys.argv) > 1:
        # 如果提供了命令行参数，使用参数作为文件列表
        files_to_delete = sys.argv[1:]
        cleanup_dependencies(files_to_delete)
    else:
        # 否则使用脚本中定义的列表
        cleanup_dependencies()


if __name__ == "__main__":
    main()

