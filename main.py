#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel/CSV文件合并工具 - PySide6版本
主入口文件
"""

import sys
import traceback
from pathlib import Path
from PySide6.QtWidgets import QApplication, QMessageBox
from PySide6.QtGui import QIcon

from ui.main_window import MainWindow
from core.resource_utils import get_resource_path


def main():
    """主函数"""
    app = QApplication(sys.argv)
    
    # 设置应用图标（使用资源路径工具，支持打包后的环境）
    icon_path = get_resource_path("icon.svg")
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))
    
    # 设置应用样式
    app.setStyle('Fusion')
    
    try:
        print("=" * 60)
        print("Excel/CSV文件合并工具")
        print("=" * 60)
        
        # 显示主窗口（保持打开，直到用户关闭）
        print("\n打开文件选择窗口...")
        main_window = MainWindow()
        main_window.show()
        
        # 运行应用事件循环，直到窗口关闭
        sys.exit(app.exec())
    
    except Exception as e:
        error_msg = f"程序执行出错: {e}\n\n{traceback.format_exc()}"
        print(error_msg)
        QMessageBox.critical(
            None,
            "错误",
            f"程序执行出错: {e}\n\n请检查文件并重试"
        )


if __name__ == "__main__":
    main()

