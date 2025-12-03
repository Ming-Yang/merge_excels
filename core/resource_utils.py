#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
资源路径工具
用于处理打包后和开发环境中的资源文件路径
"""

import sys
from pathlib import Path


def get_resource_path(relative_path: str) -> Path:
    """
    获取资源文件的绝对路径
    
    在开发环境中，返回项目根目录下的资源路径
    在打包后的环境中，返回临时解压目录下的资源路径
    
    Args:
        relative_path: 相对于 resources 目录的路径，例如 "icon.svg" 或 "icon.ico"
    
    Returns:
        Path: 资源文件的绝对路径
    """
    # 判断是否在 PyInstaller 打包后的环境中
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # 打包后的环境：资源文件在临时解压目录中
        base_path = Path(sys._MEIPASS)
    else:
        # 开发环境：资源文件在项目根目录中
        base_path = Path(__file__).parent.parent
    
    resource_path = base_path / "resources" / relative_path
    return resource_path

