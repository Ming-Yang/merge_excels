#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
表头选择对话框
"""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QPushButton, QScrollArea, QWidget,
    QFrame, QMessageBox
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont
from typing import Optional, List, Dict


class HeaderSelectionDialog(QDialog):
    """表头选择对话框"""
    
    def __init__(self, headers_info: List[Dict], parent=None):
        """
        初始化对话框
        
        Args:
            headers_info: 表头信息列表
            parent: 父窗口
        """
        super().__init__(parent)
        self.headers_info = headers_info
        self.selected_headers: Optional[List[str]] = None
        
        self.setWindowTitle("选择表头")
        self.setGeometry(100, 100, 800, 600)
        
        self._setup_ui()
    
    def _setup_ui(self):
        """设置UI"""
        layout = QVBoxLayout(self)
        
        # 标题
        title_label = QLabel("请选择要使用的表头：")
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(12)
        title_label.setFont(title_font)
        layout.addWidget(title_label)
        
        # 滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        scroll_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        
        # 为每个文件创建选择按钮
        for idx, info in enumerate(self.headers_info):
            frame = QFrame()
            frame.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Raised)
            frame.setLineWidth(2)
            
            frame_layout = QVBoxLayout(frame)
            
            # 文件信息
            file_label = QLabel(f"文件: {info['file']}")
            file_font = QFont()
            file_font.setBold(True)
            file_font.setPointSize(10)
            file_label.setFont(file_font)
            frame_layout.addWidget(file_label)
            
            sheet_label = QLabel(f"Sheet: {info['sheet']}")
            sheet_label.setFont(QFont("Arial", 9))
            frame_layout.addWidget(sheet_label)
            
            headers_text = ", ".join(info['headers'])
            headers_label = QLabel(f"列名: {headers_text}")
            headers_label.setFont(QFont("Arial", 9))
            headers_label.setWordWrap(True)
            frame_layout.addWidget(headers_label)
            
            # 选择按钮
            btn = QPushButton("选择此表头")
            btn.clicked.connect(lambda checked, i=idx: self._on_select(i))
            frame_layout.addWidget(btn)
            
            scroll_layout.addWidget(frame)
        
        scroll_area.setWidget(scroll_widget)
        layout.addWidget(scroll_area)
    
    def _on_select(self, index: int):
        """选择表头"""
        self.selected_headers = self.headers_info[index]['headers']
        self.accept()
    
    def get_selected_headers(self) -> Optional[List[str]]:
        """获取选中的表头"""
        return self.selected_headers
    
    def closeEvent(self, event):
        """窗口关闭事件"""
        if self.selected_headers is None:
            self.reject()
        event.accept()

