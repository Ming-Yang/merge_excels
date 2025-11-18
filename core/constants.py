#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
常量定义
"""

# 支持的文件格式
SUPPORTED_EXTENSIONS = {'.xlsx', '.xls', '.et', '.csv'}

# CSV文件编码列表（按优先级排序）
CSV_ENCODINGS = ['utf-8-sig', 'utf-8', 'gbk', 'gb2312', 'latin1']

# 默认文件名
DEFAULT_OUTPUT_FILENAME = "合并结果.xlsx"

