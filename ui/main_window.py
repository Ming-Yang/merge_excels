#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
主窗口
"""

from PySide6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QFileDialog, QMessageBox, QHeaderView,
    QAbstractItemView
)
from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtGui import QFont
from pathlib import Path
from typing import Dict, List, Optional
from functools import partial
import pandas as pd

from core.constants import SUPPORTED_EXTENSIONS, DEFAULT_OUTPUT_FILENAME
from core.file_reader import read_file_sheets, get_all_headers, check_headers_consistency
from core.data_merger import merge_data, save_result
from ui.header_selection_dialog import HeaderSelectionDialog
import os


class FileReaderWorker(QObject):
    """文件读取工作线程"""
    finished = Signal(str, object, bool)  # file_path, sheets_data, failed
    
    def __init__(self, file_path: str):
        super().__init__()
        self.file_path = file_path
    
    def read(self):
        """读取文件"""
        print(f"[Worker] 开始读取文件: {self.file_path}")
        try:
            sheets_data = read_file_sheets(self.file_path)
            print(f"[Worker] 文件读取完成: {self.file_path}, sheets数量: {len(sheets_data)}")
            if sheets_data:
                print(f"[Worker] 发送成功信号: {self.file_path}")
                self.finished.emit(self.file_path, sheets_data, False)
            else:
                print(f"[Worker] 发送失败信号（无数据）: {self.file_path}")
                self.finished.emit(self.file_path, {}, True)
        except Exception as e:
            print(f"[Worker] 读取文件 {self.file_path} 时出错: {e}")
            import traceback
            traceback.print_exc()
            print(f"[Worker] 发送失败信号（异常）: {self.file_path}")
            self.finished.emit(self.file_path, {}, True)


class MainWindow(QDialog):
    """主窗口"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.all_files: List[str] = []
        self.files_data_cache: Dict[str, Dict] = {}
        self.reading_files: set = set()
        self.reader_threads: Dict[str, QThread] = {}
        self.reader_workers: Dict[str, FileReaderWorker] = {}  # 保存worker对象，避免被垃圾回收
        
        self.setWindowTitle("Excel/CSV文件合并工具")
        self.setGeometry(100, 100, 700, 500)
        
        self._setup_ui()
    
    def _setup_ui(self):
        """设置UI"""
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # 标题
        title_label = QLabel("选择要合并的文件")
        title_font = QFont()
        title_font.setBold(True)
        title_font.setPointSize(12)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title_label)
        
        # 提示信息
        info_label = QLabel(
            "支持格式：.xlsx, .xls, .et, .csv"
        )
        info_label.setStyleSheet("color: gray;")
        info_label.setFont(QFont("Arial", 8))
        layout.addWidget(info_label)
        
        # 文件列表表格
        self.file_table = QTableWidget()
        self.file_table.setColumnCount(2)
        self.file_table.setHorizontalHeaderLabels(["文件名", "行数"])
        self.file_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.file_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.file_table.horizontalHeader().setStretchLastSection(False)
        self.file_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.file_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.file_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        layout.addWidget(self.file_table)
        
        # 统计信息和删除按钮
        info_button_layout = QHBoxLayout()
        
        self.count_label = QLabel("已选择 0 个文件")
        info_button_layout.addWidget(self.count_label)
        
        self.total_rows_label = QLabel("总行数: 0")
        total_rows_font = QFont()
        total_rows_font.setBold(True)
        self.total_rows_label.setFont(total_rows_font)
        self.total_rows_label.setStyleSheet("color: blue;")
        info_button_layout.addWidget(self.total_rows_label)
        
        info_button_layout.addStretch()
        
        self.btn_delete = QPushButton("删除选中")
        self.btn_delete.setStyleSheet("background-color: #ff9800; color: white;")
        self.btn_delete.clicked.connect(self._delete_selected)
        info_button_layout.addWidget(self.btn_delete)
        
        layout.addLayout(info_button_layout)
        
        # 按钮区域
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        self.btn_select_files = QPushButton("选择文件")
        self.btn_select_files.setMinimumHeight(40)
        self.btn_select_files.clicked.connect(self._select_files)
        button_layout.addWidget(self.btn_select_files)
        
        self.btn_select_folder = QPushButton("选择路径")
        self.btn_select_folder.setMinimumHeight(40)
        self.btn_select_folder.clicked.connect(self._select_folder)
        button_layout.addWidget(self.btn_select_folder)
        
        self.btn_start = QPushButton("开始处理")
        self.btn_start.setMinimumHeight(40)
        self.btn_start.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.btn_start.clicked.connect(self._start_process)
        button_layout.addWidget(self.btn_start)
        
        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.setMinimumHeight(40)
        self.btn_cancel.setStyleSheet("background-color: #f44336; color: white;")
        self.btn_cancel.clicked.connect(self.close)
        button_layout.addWidget(self.btn_cancel)
        
        layout.addLayout(button_layout)
    
    def _format_rows_display(self, rows: int) -> str:
        """格式化行数显示"""
        if rows == -1:
            return "失败"
        elif rows == 0:
            return "-"
        else:
            return str(rows)
    
    def _add_file_to_list(self, file_path: str, rows: int = 0, failed: bool = False):
        """添加文件到列表"""
        if failed:
            rows = -1
        
        row = self.file_table.rowCount()
        self.file_table.insertRow(row)
        
        file_item = QTableWidgetItem(file_path)
        self.file_table.setItem(row, 0, file_item)
        
        rows_item = QTableWidgetItem(self._format_rows_display(rows))
        rows_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.file_table.setItem(row, 1, rows_item)
        
        # 存储文件信息
        if file_path not in self.files_data_cache:
            self.files_data_cache[file_path] = {'rows': rows}
        
        self._update_count_label()
        self._update_total_rows()
    
    def _update_file_rows(self, file_path: str, rows: int, failed: bool = False):
        """更新文件行数"""
        if failed:
            rows = -1
        
        # 查找对应的行
        for row in range(self.file_table.rowCount()):
            item = self.file_table.item(row, 0)
            if item and item.text() == file_path:
                rows_item = QTableWidgetItem(self._format_rows_display(rows))
                rows_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.file_table.setItem(row, 1, rows_item)
                
                if file_path in self.files_data_cache:
                    self.files_data_cache[file_path]['rows'] = rows
                break
        
        self._update_total_rows()
    
    def _update_count_label(self):
        """更新文件计数标签"""
        count = len(self.all_files)
        self.count_label.setText(f"已选择 {count} 个文件")
    
    def _update_total_rows(self):
        """更新总行数显示"""
        total_rows = sum(
            info.get('rows', 0) 
            for info in self.files_data_cache.values() 
            if info.get('rows', 0) > 0
        )
        self.total_rows_label.setText(f"总行数: {total_rows}")
    
    def _delete_selected(self):
        """删除选中的文件"""
        selected_rows = self.file_table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.information(self, "提示", "请先选择一个或多个文件")
            return
        
        # 收集要删除的文件路径（从后往前删除，避免索引变化）
        rows_to_delete = sorted([row.row() for row in selected_rows], reverse=True)
        files_to_delete = []
        
        for row in rows_to_delete:
            item = self.file_table.item(row, 0)
            if item:
                file_path = item.text()
                if file_path in self.all_files:
                    files_to_delete.append(file_path)
        
        if files_to_delete:
            # 从列表中删除
            for file_path in files_to_delete:
                self.all_files.remove(file_path)
                if file_path in self.files_data_cache:
                    del self.files_data_cache[file_path]
                if file_path in self.reading_files:
                    self.reading_files.discard(file_path)
                # 停止正在读取的线程
                if file_path in self.reader_threads:
                    thread = self.reader_threads[file_path]
                    thread.quit()  # 请求线程退出
                    thread.wait(3000)  # 等待最多3秒，确保线程完全退出
                    # 如果线程还在运行，强制终止（不推荐，但作为最后手段）
                    if thread.isRunning():
                        print(f"警告: 线程仍在运行，强制终止: {file_path}")
                        thread.terminate()
                        thread.wait(1000)
                    # 清理引用（线程已完全退出）
                    if file_path in self.reader_threads:
                        del self.reader_threads[file_path]
                if file_path in self.reader_workers:
                    del self.reader_workers[file_path]
            
            # 从表格中删除
            for row in rows_to_delete:
                self.file_table.removeRow(row)
            
            self._update_count_label()
            self._update_total_rows()
            print(f"已删除 {len(files_to_delete)} 个文件")
    
    def _read_file_async(self, file_path: str):
        """异步读取文件"""
        if file_path in self.reading_files:
            return
        
        self.reading_files.add(file_path)
        
        # 创建工作线程
        thread = QThread()
        worker = FileReaderWorker(file_path)
        worker.moveToThread(thread)
        
        # 连接信号 - 确保使用队列连接（跨线程通信）
        thread.started.connect(worker.read)
        # 显式指定连接类型为队列连接，确保跨线程信号能正确传递
        worker.finished.connect(
            self._on_file_read_finished,
            type=Qt.ConnectionType.QueuedConnection
        )
        worker.finished.connect(thread.quit, type=Qt.ConnectionType.QueuedConnection)
        # 线程完全退出后再删除线程对象和清理引用
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(partial(self._cleanup_thread, file_path))
        
        # 保存引用，避免被垃圾回收
        self.reader_threads[file_path] = thread
        self.reader_workers[file_path] = worker
        
        thread.start()
        print(f"启动线程读取文件: {file_path}")
    
    def _cleanup_thread(self, file_path: str):
        """清理线程和worker引用（在线程完全退出后调用）"""
        if file_path in self.reader_threads:
            del self.reader_threads[file_path]
            print(f"已清理线程引用: {file_path}")
        if file_path in self.reader_workers:
            del self.reader_workers[file_path]
            print(f"已清理worker引用: {file_path}")
    
    def _on_file_read_finished(self, file_path: str, sheets_data: Dict[str, pd.DataFrame], failed: bool):
        """文件读取完成回调"""
        print(f"回调函数被触发: {file_path}, 失败: {failed}, 数据: {bool(sheets_data)}")
        self.reading_files.discard(file_path)
        
        # 注意：不要在这里删除线程引用，让线程自己通过 finished 信号来清理
        # 这样可以确保线程完全退出后再清理，避免 "Destroyed while thread is still running" 错误
        
        # 检查文件是否还在列表中
        if file_path not in self.all_files:
            print(f"文件 {file_path} 已从列表中移除，跳过处理")
            return
        
        print(f"文件读取结束: {file_path}, 失败: {failed}")
        
        if failed or not sheets_data:
            self._update_file_rows(file_path, 0, failed=True)
            self.files_data_cache[file_path] = {}
        else:
            # 保存数据到缓存
            self.files_data_cache[file_path] = {'data': sheets_data}
            # 计算总行数
            total_rows = sum(len(df) for df in sheets_data.values())
            self._update_file_rows(file_path, total_rows, failed=False)
            # 更新缓存
            self.files_data_cache[file_path]['rows'] = total_rows
    
    def _add_file_with_async_read(self, file_path: str):
        """添加文件到列表并启动异步读取"""
        if file_path not in self.all_files:
            self.all_files.append(file_path)
            self._add_file_to_list(file_path, 0)  # 初始显示"-"
            self._read_file_async(file_path)
    
    def _select_files(self):
        """选择文件"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择要合并的文件",
            "",
            "所有支持格式 (*.xlsx *.xls *.et *.csv);;Excel文件 (*.xlsx *.xls *.et);;CSV文件 (*.csv);;所有文件 (*.*)"
        )
        
        if files:
            for file_path in files:
                self._add_file_with_async_read(file_path)
            print(f"添加了 {len(files)} 个文件")
    
    def _select_folder(self):
        """选择文件夹"""
        folder = QFileDialog.getExistingDirectory(
            self,
            "选择包含文件的文件夹（将合并该文件夹下所有支持格式的文件）"
        )
        
        if folder:
            folder_path = Path(folder)
            new_files = []
            for ext in SUPPORTED_EXTENSIONS:
                new_files.extend(folder_path.glob(f"*{ext}"))
                new_files.extend(folder_path.glob(f"*{ext.upper()}"))
            
            new_files = [str(f) for f in new_files]
            if new_files:
                for file_path in new_files:
                    self._add_file_with_async_read(file_path)
                print(f"从文件夹添加了 {len(new_files)} 个文件")
            else:
                QMessageBox.warning(self, "警告", f"文件夹 {folder} 中没有找到支持格式的文件")
    
    def _start_process(self):
        """开始处理"""
        if len(self.all_files) == 0:
            QMessageBox.warning(self, "警告", "请至少选择一个文件或文件夹")
            return
        
        # 检查是否有文件还在读取中
        if self.reading_files:
            reading_count = len(self.reading_files)
            QMessageBox.warning(
                self,
                "文件读取中",
                f"还有 {reading_count} 个文件正在读取中，请等待读取完成后再开始处理。"
            )
            return
        
        # 检查是否有有效数据
        valid_files_data = {}
        for file_path in self.all_files:
            if file_path in self.files_data_cache:
                cache = self.files_data_cache[file_path]
                if 'data' in cache and cache['data']:
                    valid_files_data[file_path] = cache['data']
        
        if not valid_files_data:
            QMessageBox.warning(self, "警告", "没有读取到任何有效数据\n\n请检查文件是否正确，或重新选择文件")
            return
        
        # 开始处理流程（不关闭窗口）
        self._process_files(valid_files_data)
    
    def _process_files(self, files_data: Dict[str, Dict[str, pd.DataFrame]]):
        """处理文件合并流程"""
        try:
            print(f"\n总共选择了 {len(self.all_files)} 个文件")
            
            # 确定默认保存路径
            last_selected_folder = self.get_last_selected_folder()
            if last_selected_folder:
                default_save_dir = Path(last_selected_folder)
            else:
                default_save_dir = Path(self.all_files[-1]).parent if self.all_files else Path.cwd()
            
            # 获取所有表头
            headers_info = get_all_headers(files_data)
            print(f"\n共找到 {len(headers_info)} 个表/Sheet")
            
            # 检查表头一致性
            is_consistent, result = check_headers_consistency(headers_info)
            
            if is_consistent:
                target_headers = result
                print("所有表头一致，直接合并")
            else:
                print("表头不一致，需要用户选择")
                header_dialog = HeaderSelectionDialog(headers_info, self)
                if header_dialog.exec() != QDialog.DialogCode.Accepted:
                    QMessageBox.warning(
                        self,
                        "警告",
                        "未选择表头\n\n请重新点击\"开始处理\"并选择表头"
                    )
                    return
                
                target_headers = header_dialog.get_selected_headers()
                if not target_headers:
                    QMessageBox.warning(
                        self,
                        "警告",
                        "未选择表头\n\n请重新点击\"开始处理\"并选择表头"
                    )
                    return
            
            # 合并数据
            print("\n正在合并数据...")
            merged_df, statistics = merge_data(files_data, target_headers)
            
            if merged_df.empty:
                QMessageBox.warning(
                    self,
                    "警告",
                    "合并结果为空\n\n请检查文件数据是否正确"
                )
                return
            
            # 检查重复行并询问是否去重
            original_rows = len(merged_df)
            duplicate_count = merged_df.duplicated().sum()
            
            if duplicate_count > 0:
                print(f"\n检测到 {duplicate_count} 行完全重复的数据")
                reply = QMessageBox.question(
                    self,
                    "发现重复数据",
                    f"检测到 {duplicate_count} 行完全重复的数据。\n\n是否要去除重复行？\n\n是(Y) - 去除重复行\n否(N) - 保留所有数据",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes
                )
                
                if reply == QMessageBox.StandardButton.Yes:
                    merged_df = merged_df.drop_duplicates()
                    deduplicated_rows = len(merged_df)
                    removed_rows = original_rows - deduplicated_rows
                    print(f"已去除 {removed_rows} 行重复数据，剩余 {deduplicated_rows} 行")
                else:
                    print("保留所有数据（包括重复行）")
            else:
                print("\n未发现重复数据")
            
            # 显示统计信息
            total_rows = len(merged_df)
            print("\n" + "=" * 60)
            print("合并统计:")
            print("=" * 60)
            for stat in statistics:
                print(f"  {stat['file']} - {stat['sheet']}: {stat['rows']} 行")
            print(f"\n合并后总计: {original_rows} 行")
            if duplicate_count > 0:
                if total_rows < original_rows:
                    removed_count = original_rows - total_rows
                    print(f"去重后总计: {total_rows} 行 (已去除 {removed_count} 行重复数据)")
                else:
                    print(f"最终总计: {total_rows} 行 (检测到 {duplicate_count} 行重复，但未去重)")
            else:
                print(f"最终总计: {total_rows} 行")
            print("=" * 60)
            
            # 保存结果
            default_filename = DEFAULT_OUTPUT_FILENAME
            output_path = None
            
            while True:
                output_path, _ = QFileDialog.getSaveFileName(
                    self,
                    "保存合并结果 - 请选择格式、路径和文件名",
                    str(default_save_dir / default_filename),
                    "Excel文件 (*.xlsx);;CSV文件 (*.csv);;所有文件 (*.*)"
                )
                
                if not output_path:
                    print("\n未保存文件")
                    return
                
                # 尝试保存文件
                if save_result(merged_df, output_path):
                    # 保存成功，验证文件
                    if os.path.exists(output_path):
                        file_size = os.path.getsize(output_path)
                        if file_size > 0:
                            print(f"\n结果已保存到: {output_path}")
                            QMessageBox.information(
                                self,
                                "完成",
                                f"合并完成！\n\n共合并 {total_rows} 行数据\n\n结果已保存到:\n{output_path}"
                            )
                            break
                        else:
                            # 文件大小为0
                            reply = QMessageBox.question(
                                self,
                                "保存失败",
                                "文件保存后大小为0，可能保存失败。\n\n是否重新选择保存位置？\n\n是(Y) - 重新选择\n否(N) - 取消保存",
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                QMessageBox.StandardButton.Yes
                            )
                            if reply == QMessageBox.StandardButton.No:
                                return
                            # 更新默认路径
                            default_save_dir = Path(output_path).parent
                            default_filename = Path(output_path).name
                            continue
                    else:
                        # 文件不存在
                        reply = QMessageBox.question(
                            self,
                            "保存失败",
                            "文件保存失败，文件不存在。\n\n是否重新选择保存位置？\n\n是(Y) - 重新选择\n否(N) - 取消保存",
                            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                            QMessageBox.StandardButton.Yes
                        )
                        if reply == QMessageBox.StandardButton.No:
                            return
                        default_save_dir = Path(output_path).parent
                        default_filename = Path(output_path).name
                        continue
                else:
                    # 保存失败（异常已在save_result中处理）
                    reply = QMessageBox.question(
                        self,
                        "保存失败",
                        "保存文件时出错。\n\n是否重新选择保存位置？\n\n是(Y) - 重新选择\n否(N) - 取消保存",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                        QMessageBox.StandardButton.Yes
                    )
                    if reply == QMessageBox.StandardButton.No:
                        return
                    default_save_dir = Path(output_path).parent if output_path else default_save_dir
                    default_filename = Path(output_path).name if output_path else default_filename
                    continue
                    
        except Exception as e:
            import traceback
            error_msg = f"程序执行出错: {e}\n\n{traceback.format_exc()}"
            print(error_msg)
            QMessageBox.critical(
                self,
                "错误",
                f"程序执行出错: {e}\n\n请检查文件并重试"
            )
    
    def get_files_data(self) -> Dict[str, Dict[str, pd.DataFrame]]:
        """获取所有文件的数据"""
        files_data = {}
        for file_path in self.all_files:
            if file_path in self.files_data_cache:
                cache = self.files_data_cache[file_path]
                if 'data' in cache and cache['data']:
                    files_data[file_path] = cache['data']
        return files_data
    
    def get_last_selected_folder(self) -> Optional[str]:
        """获取最后选择的文件夹（简化实现，返回最后一个文件的文件夹）"""
        if self.all_files:
            return str(Path(self.all_files[-1]).parent)
        return None
    
    def closeEvent(self, event):
        """窗口关闭事件"""
        # 停止所有读取线程
        for file_path, thread in list(self.reader_threads.items()):
            thread.quit()  # 请求线程退出
            thread.wait(3000)  # 等待最多3秒
            # 如果线程还在运行，强制终止
            if thread.isRunning():
                print(f"警告: 关闭窗口时线程仍在运行，强制终止: {file_path}")
                thread.terminate()
                thread.wait(1000)
        # 清理所有引用
        self.reader_threads.clear()
        self.reader_workers.clear()
        event.accept()

