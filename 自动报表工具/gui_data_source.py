#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI数据源管理组件
负责数据源配置相关的UI界面和逻辑
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, Any, Optional

class DataSourceManager:
    """数据源管理器"""
    
    def __init__(self, parent_notebook):
        self.parent_notebook = parent_notebook
        self.data_source_config = {}  # 存储数据源配置
        
        # 创建数据源标签页
        self.tab = ttk.Frame(self.parent_notebook)
        self.parent_notebook.add(self.tab, text="数据源配置")
        
        self._create_widgets()
        
    def _create_widgets(self):
        """创建数据源配置界面"""
        # 数据源类型
        ttk.Label(self.tab, text="数据源类型：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.data_source_type = ttk.Combobox(self.tab, values=["excel", "csv", "sql", "api"], state="readonly", width=15)
        self.data_source_type.set("excel")
        self.data_source_type.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        self.data_source_type.bind("<<ComboboxSelected>>", self._on_data_source_type_change)
        
        # 数据源路径/URL
        ttk.Label(self.tab, text="数据源路径/URL：").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.data_source_path = ttk.Entry(self.tab, width=60)
        self.data_source_path.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        self.browse_btn = ttk.Button(self.tab, text="浏览...", command=self._browse_file)
        self.browse_btn.grid(row=1, column=2, padx=5, pady=5)
        
        # 数据源参数框架
        self.data_source_params_frame = ttk.LabelFrame(self.tab, text="数据源参数", padding="10")
        self.data_source_params_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=10, sticky=tk.W+tk.E)
        
        # 为不同数据源类型创建参数面板
        self._create_excel_params()
        self._create_csv_params()
        self._create_sql_params()
        self._create_api_params()
        
        # 初始显示Excel参数
        self._show_params_panel("excel")
        
    def _create_excel_params(self):
        """创建Excel数据源参数面板"""
        self.excel_params_frame = ttk.Frame(self.data_source_params_frame)
        
        ttk.Label(self.excel_params_frame, text="工作表名称：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.excel_sheet = ttk.Entry(self.excel_params_frame, width=20)
        self.excel_sheet.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(self.excel_params_frame, text="分块大小：").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.excel_chunksize = ttk.Entry(self.excel_params_frame, width=10)
        self.excel_chunksize.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        
    def _create_csv_params(self):
        """创建CSV数据源参数面板"""
        self.csv_params_frame = ttk.Frame(self.data_source_params_frame)
        
        ttk.Label(self.csv_params_frame, text="编码：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.csv_encoding = ttk.Combobox(self.csv_params_frame, values=["utf-8", "gbk", "gb2312", "latin1"], state="readonly", width=15)
        self.csv_encoding.set("utf-8")
        self.csv_encoding.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(self.csv_params_frame, text="分隔符：").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.csv_separator = ttk.Entry(self.csv_params_frame, width=5)
        self.csv_separator.insert(0, ",")
        self.csv_separator.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(self.csv_params_frame, text="分块大小：").grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        self.csv_chunksize = ttk.Entry(self.csv_params_frame, width=10)
        self.csv_chunksize.grid(row=0, column=5, padx=5, pady=5, sticky=tk.W)
        
    def _create_sql_params(self):
        """创建SQL数据源参数面板"""
        self.sql_params_frame = ttk.Frame(self.data_source_params_frame)
        
        ttk.Label(self.sql_params_frame, text="数据库连接字符串：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.sql_connection = ttk.Entry(self.sql_params_frame, width=50)
        self.sql_connection.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(self.sql_params_frame, text="SQL查询：").grid(row=1, column=0, padx=5, pady=5, sticky=tk.NW)
        self.sql_query = tk.Text(self.sql_params_frame, width=60, height=5)
        self.sql_query.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        scrollbar = ttk.Scrollbar(self.sql_params_frame, command=self.sql_query.yview)
        scrollbar.grid(row=1, column=2, sticky=tk.NS)
        self.sql_query.config(yscrollcommand=scrollbar.set)
        
    def _create_api_params(self):
        """创建API数据源参数面板"""
        self.api_params_frame = ttk.Frame(self.data_source_params_frame)
        
        ttk.Label(self.api_params_frame, text="请求方法：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.api_method = ttk.Combobox(self.api_params_frame, values=["GET", "POST"], state="readonly", width=10)
        self.api_method.set("GET")
        self.api_method.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(self.api_params_frame, text="请求头：").grid(row=1, column=0, padx=5, pady=5, sticky=tk.NW)
        self.api_headers = tk.Text(self.api_params_frame, width=40, height=3)
        self.api_headers.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(self.api_params_frame, text="查询参数：").grid(row=2, column=0, padx=5, pady=5, sticky=tk.NW)
        self.api_params = tk.Text(self.api_params_frame, width=40, height=3)
        self.api_params.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(self.api_params_frame, text="响应数据键：").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.api_response_key = ttk.Entry(self.api_params_frame, width=30)
        self.api_response_key.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        
    def _on_data_source_type_change(self, event=None):
        """数据源类型改变事件处理"""
        source_type = self.data_source_type.get()
        self._show_params_panel(source_type)
        
    def _show_params_panel(self, source_type: str):
        """显示对应的参数面板"""
        # 隐藏所有参数面板
        for frame in [self.excel_params_frame, self.csv_params_frame, 
                     self.sql_params_frame, self.api_params_frame]:
            frame.pack_forget()
            
        # 显示对应的参数面板
        if source_type == "excel":
            self.excel_params_frame.pack(fill=tk.X)
        elif source_type == "csv":
            self.csv_params_frame.pack(fill=tk.X)
        elif source_type == "sql":
            self.sql_params_frame.pack(fill=tk.X)
        elif source_type == "api":
            self.api_params_frame.pack(fill=tk.X)
            
    def _browse_file(self):
        """浏览文件对话框"""
        source_type = self.data_source_type.get()
        
        if source_type in ["excel", "csv"]:
            file_path = filedialog.askopenfilename(
                title="选择数据文件",
                filetypes=[
                    ("Excel文件", "*.xlsx *.xls"),
                    ("CSV文件", "*.csv"),
                    ("所有文件", "*.*")
                ]
            )
            if file_path:
                self.data_source_path.delete(0, tk.END)
                self.data_source_path.insert(0, file_path)
        else:
            messagebox.showinfo("提示", f"{source_type.upper()}数据源不需要浏览文件")
            
    def get_data_source_config(self) -> Dict[str, Any]:
        """获取数据源配置"""
        config = {
            'type': self.data_source_type.get(),
            'path': self.data_source_path.get()
        }
        
        source_type = self.data_source_type.get()
        
        if source_type == "excel":
            config.update({
                'sheet_name': self.excel_sheet.get(),
                'chunksize': int(self.excel_chunksize.get()) if self.excel_chunksize.get() else None
            })
        elif source_type == "csv":
            config.update({
                'encoding': self.csv_encoding.get(),
                'separator': self.csv_separator.get(),
                'chunksize': int(self.csv_chunksize.get()) if self.csv_chunksize.get() else None
            })
        elif source_type == "sql":
            config.update({
                'connection_string': self.sql_connection.get(),
                'query': self.sql_query.get("1.0", tk.END).strip()
            })
        elif source_type == "api":
            try:
                headers = eval(self.api_headers.get("1.0", tk.END).strip()) if self.api_headers.get("1.0", tk.END).strip() else {}
            except:
                headers = {}
            
            try:
                params = eval(self.api_params.get("1.0", tk.END).strip()) if self.api_params.get("1.0", tk.END).strip() else {}
            except:
                params = {}
                
            config.update({
                'method': self.api_method.get(),
                'headers': headers,
                'params': params,
                'response_key': self.api_response_key.get()
            })
            
        return config
        
    def set_data_source_config(self, config: Dict[str, Any]):
        """设置数据源配置"""
        self.data_source_type.set(config.get('type', 'excel'))
        self.data_source_path.delete(0, tk.END)
        self.data_source_path.insert(0, config.get('path', ''))
        
        source_type = config.get('type', 'excel')
        
        if source_type == "excel":
            self.excel_sheet.delete(0, tk.END)
            self.excel_sheet.insert(0, config.get('sheet_name', ''))
            self.excel_chunksize.delete(0, tk.END)
            self.excel_chunksize.insert(0, str(config.get('chunksize', '')))
        elif source_type == "csv":
            self.csv_encoding.set(config.get('encoding', 'utf-8'))
            self.csv_separator.delete(0, tk.END)
            self.csv_separator.insert(0, config.get('separator', ','))
            self.csv_chunksize.delete(0, tk.END)
            self.csv_chunksize.insert(0, str(config.get('chunksize', '')))
        elif source_type == "sql":
            self.sql_connection.delete(0, tk.END)
            self.sql_connection.insert(0, config.get('connection_string', ''))
            self.sql_query.delete("1.0", tk.END)
            self.sql_query.insert("1.0", config.get('query', ''))
        elif source_type == "api":
            self.api_method.set(config.get('method', 'GET'))
            self.api_headers.delete("1.0", tk.END)
            self.api_headers.insert("1.0", str(config.get('headers', {})))
            self.api_params.delete("1.0", tk.END)
            self.api_params.insert("1.0", str(config.get('params', {})))
            self.api_response_key.delete(0, tk.END)
            self.api_response_key.insert(0, config.get('response_key', ''))
            
        # 更新显示的参数面板
        self._show_params_panel(source_type)