#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
文件读取模块
"""

import pandas as pd
from pathlib import Path
from typing import Dict, Optional
from .constants import SUPPORTED_EXTENSIONS, CSV_ENCODINGS


def read_file_sheets(file_path: str) -> Dict[str, pd.DataFrame]:
    """
    读取文件的所有sheet，返回字典 {sheet_name: DataFrame}
    
    Args:
        file_path: 文件路径
        
    Returns:
        字典，键为sheet名称，值为DataFrame。如果读取失败，返回空字典
    """
    file_path = Path(file_path)
    ext = file_path.suffix.lower()
    sheets_data = {}
    
    try:
        if ext == '.csv':
            # CSV文件只有一个sheet，尝试多种编码
            df = None
            for encoding in CSV_ENCODINGS:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue
            if df is None or df.empty:
                return {}
            sheets_data['Sheet1'] = df
        else:
            # Excel文件可能有多个sheet
            # xlsx使用openpyxl，xls和et让pandas自动选择引擎
            engine = 'openpyxl' if ext in ['.xlsx', '.et'] else None
            excel_file = None
            try:
                excel_file = pd.ExcelFile(file_path, engine=engine)
            except Exception:
                # 如果指定引擎失败，尝试默认引擎
                try:
                    excel_file = pd.ExcelFile(file_path)
                    engine = None  # 使用默认引擎
                except Exception as e:
                    print(f"无法读取Excel文件 {file_path}: {e}")
                    return {}
            
            if excel_file is None:
                return {}
            
            for sheet_name in excel_file.sheet_names:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, engine=engine)
                    if not df.empty:
                        sheets_data[sheet_name] = df
                except Exception as e:
                    print(f"警告: 读取 {file_path} 的 {sheet_name} sheet 时出错: {e}")
                    continue
    except Exception as e:
        print(f"错误: 读取文件 {file_path} 时出错: {e}")
        return {}
    
    return sheets_data


def get_all_headers(files_data: Dict[str, Dict[str, pd.DataFrame]]) -> list:
    """
    获取所有文件的表头信息
    
    Args:
        files_data: 文件数据字典，格式为 {file_path: {sheet_name: DataFrame}}
        
    Returns:
        表头信息列表，每个元素包含文件、sheet、表头等信息
    """
    headers_info = []
    for file_path, sheets_data in files_data.items():
        for sheet_name, df in sheets_data.items():
            headers = list(df.columns)
            headers_info.append({
                'file': Path(file_path).name,
                'sheet': sheet_name,
                'headers': headers,
                'file_path': file_path
            })
    return headers_info


def check_headers_consistency(headers_info: list) -> tuple:
    """
    检查所有表头是否一致
    
    Args:
        headers_info: 表头信息列表
        
    Returns:
        (是否一致, 结果)
        如果一致，返回 (True, 第一个表头列表)
        如果不一致，返回 (False, headers_info)
    """
    if not headers_info:
        return True, None
    
    first_headers = headers_info[0]['headers']
    for info in headers_info[1:]:
        if info['headers'] != first_headers:
            return False, headers_info
    return True, first_headers

