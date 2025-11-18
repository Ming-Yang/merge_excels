#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
数据合并模块
"""

import pandas as pd
from pathlib import Path
from typing import Dict, Tuple, List


def merge_data(files_data: Dict[str, Dict[str, pd.DataFrame]], 
               target_headers: List[str]) -> Tuple[pd.DataFrame, List[Dict]]:
    """
    合并数据
    
    Args:
        files_data: 文件数据字典，格式为 {file_path: {sheet_name: DataFrame}}
        target_headers: 目标表头列表
        
    Returns:
        (合并后的DataFrame, 统计信息列表)
    """
    merged_data = []
    statistics = []
    
    for file_path, sheets_data in files_data.items():
        for sheet_name, df in sheets_data.items():
            original_rows = len(df)
            
            # 对齐列：使用reindex方法，缺失的列会自动填充NaN
            aligned_df = df.reindex(columns=target_headers)
            
            merged_data.append(aligned_df)
            
            statistics.append({
                'file': Path(file_path).name,
                'sheet': sheet_name,
                'rows': original_rows
            })
    
    if merged_data:
        result_df = pd.concat(merged_data, ignore_index=True)
        return result_df, statistics
    else:
        return pd.DataFrame(), statistics


def save_result(df: pd.DataFrame, 
                output_path: str) -> bool:
    """
    保存合并结果
    
    Args:
        df: 要保存的DataFrame
        output_path: 保存路径
        
    Returns:
        是否保存成功
    """
    try:
        if output_path.lower().endswith('.csv'):
            df.to_csv(output_path, index=False, encoding='utf-8-sig')
        else:
            df.to_excel(output_path, index=False, engine='openpyxl')
        
        # 验证文件是否存在且大小大于0
        import os
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            return file_size > 0
        return False
    except Exception as e:
        print(f"保存文件时出错: {e}")
        return False

