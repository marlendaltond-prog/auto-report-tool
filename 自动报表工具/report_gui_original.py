#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动化报表工具GUI界面
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import json
import os
# 添加YAML支持
import yaml
import threading
from datetime import datetime

# 导入现有功能
from auto_report import (
    ReportConfig, AutoReportEngine, ConfigManager,
    third_party_available, logger, schedule_manager
)

class ReportGUI:
    """自动化报表工具GUI界面类"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("自动化报表工具 - AutoReport Pro")
        self.root.geometry("1200x800")
        self.root.resizable(True, True)
        
        # 直接以管理员身份登录，无需登录界面
        self.current_user = "admin"
        
        # 设置主题
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        # 检查依赖
        if not third_party_available:
            messagebox.showerror("错误", "缺少必要的依赖包！请安装：\npip install pandas openpyxl sqlalchemy jinja2 reportlab requests")
            self.root.destroy()
            return
        
        # 全局配置
        self.config_manager = ConfigManager()
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标签页
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 创建各标签页
        self.create_data_source_tab()
        self.create_report_config_tab()
        self.create_data_entry_tab()
        self.create_output_tab()
        self.create_charts_tab()
        self.create_calculations_tab()
        self.create_schedule_tab()
        
        # 创建底部按钮栏
        self.create_button_bar()
        
        # 创建状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("准备就绪")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def create_data_source_tab(self):
        """创建数据源配置标签页"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="数据源配置")
        
        # 数据源类型
        ttk.Label(tab, text="数据源类型：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.data_source_type = ttk.Combobox(tab, values=["excel", "csv", "sql", "api"], state="readonly", width=15)
        self.data_source_type.set("excel")
        self.data_source_type.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        self.data_source_type.bind("<<ComboboxSelected>>", self.on_data_source_type_change)
        
        # 数据源路径/URL
        ttk.Label(tab, text="数据源路径/URL：").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.data_source_path = ttk.Entry(tab, width=60)
        self.data_source_path.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        self.browse_btn = ttk.Button(tab, text="浏览...", command=self.browse_file)
        self.browse_btn.grid(row=1, column=2, padx=5, pady=5)
        
        # 数据源参数框架
        self.data_source_params_frame = ttk.LabelFrame(tab, text="数据源参数", padding="10")
        self.data_source_params_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=10, sticky=tk.W+tk.E)
        
        # 为不同数据源类型创建参数面板
        self.create_excel_params()
        self.create_csv_params()
        self.create_sql_params()
        self.create_api_params()
        
        # 初始显示Excel参数
        self.show_params_panel("excel")
    
    def create_excel_params(self):
        """创建Excel数据源参数面板"""
        self.excel_params_frame = ttk.Frame(self.data_source_params_frame)
        
        ttk.Label(self.excel_params_frame, text="工作表名称：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.excel_sheet = ttk.Entry(self.excel_params_frame, width=20)
        self.excel_sheet.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(self.excel_params_frame, text="分块大小：").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.excel_chunksize = ttk.Entry(self.excel_params_frame, width=10)
        self.excel_chunksize.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
    
    def create_csv_params(self):
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
    
    def create_sql_params(self):
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
    
    def create_api_params(self):
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
    
    def create_report_config_tab(self):
        """创建报表配置标签页"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="报表配置")
        
        # 报表名称
        ttk.Label(tab, text="报表名称：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.report_name = ttk.Entry(tab, width=30)
        self.report_name.insert(0, "自动化报表")
        self.report_name.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 配置文件
        ttk.Label(tab, text="配置文件：").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.config_file = ttk.Entry(tab, width=50)
        self.config_file.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Button(tab, text="选择...", command=self.browse_config_file).grid(row=1, column=2, padx=5, pady=5)
        
        # 配置模板下拉框
        ttk.Label(tab, text="配置模板：").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.config_templates = ttk.Combobox(tab, values=["空模板", "销售报表", "财务报表", "库存报表"], state="readonly", width=15)
        self.config_templates.set("空模板")
        self.config_templates.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Button(tab, text="应用模板", command=self.apply_config_template).grid(row=2, column=2, padx=5, pady=5)
        
        # 调度设置（可选）
        ttk.LabelFrame(tab, text="调度设置（可选）").grid(row=3, column=0, columnspan=3, padx=5, pady=10, sticky=tk.W+tk.E)
        ttk.Label(tab, text="Cron表达式：").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        self.cron_expression = ttk.Entry(tab, width=30)
        self.cron_expression.grid(row=4, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 邮件接收者（可选）
        ttk.LabelFrame(tab, text="邮件设置（可选）").grid(row=5, column=0, columnspan=3, padx=5, pady=10, sticky=tk.W+tk.E)
        ttk.Label(tab, text="接收者邮箱：").grid(row=6, column=0, padx=5, pady=5, sticky=tk.W)
        self.email_recipients = ttk.Entry(tab, width=50)
        self.email_recipients.grid(row=6, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Label(tab, text="（逗号分隔）").grid(row=6, column=2, padx=5, pady=5, sticky=tk.W)
    
    def create_output_tab(self):
        """创建输出配置标签页"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="输出配置")
        
        # 输出格式
        ttk.Label(tab, text="输出格式：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        self.output_formats = {}
        formats = ["excel", "pdf", "html", "email"]
        for i, fmt in enumerate(formats):
            var = tk.BooleanVar()
            var.set(True if fmt == "excel" else False)
            self.output_formats[fmt] = var
            ttk.Checkbutton(tab, text=fmt, variable=var).grid(row=0, column=i+1, padx=10, pady=5, sticky=tk.W)
        
        # 输出目录
        ttk.Label(tab, text="输出目录：").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.output_dir = ttk.Entry(tab, width=50)
        self.output_dir.insert(0, self.config_manager.get("output_dir", "reports"))
        self.output_dir.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Button(tab, text="浏览...", command=self.browse_output_dir).grid(row=1, column=2, padx=5, pady=5)
        
        # 模板类型
        ttk.Label(tab, text="模板类型：").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        
        # 获取可用模板列表
        templates_dir = os.path.join(os.getcwd(), 'templates')
        available_templates = ['default']
        if os.path.exists(templates_dir):
            for filename in os.listdir(templates_dir):
                if filename.endswith('.html'):
                    available_templates.append(os.path.splitext(filename)[0])
        
        self.template_type = ttk.Combobox(tab, values=available_templates, state="readonly", width=20)
        self.template_type.set('default')
        self.template_type.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 模板文件（可选）
        ttk.Label(tab, text="自定义模板：").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.template_file = ttk.Entry(tab, width=50)
        self.template_file.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Button(tab, text="浏览...", command=self.browse_template_file).grid(row=3, column=2, padx=5, pady=5)
    
    def create_charts_tab(self):
        """创建图表配置标签页"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="图表配置")
        
        # 图表列表
        self.charts_list = ttk.Treeview(tab, columns=("type", "title", "x", "y"), show="headings", height=5)
        self.charts_list.heading("type", text="图表类型")
        self.charts_list.heading("title", text="标题")
        self.charts_list.heading("x", text="X轴字段")
        self.charts_list.heading("y", text="Y轴字段")
        self.charts_list.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W+tk.E)
        
        # 图表配置
        ttk.LabelFrame(tab, text="图表设置").grid(row=1, column=0, columnspan=3, padx=5, pady=10, sticky=tk.W+tk.E)
        
        ttk.Label(tab, text="图表类型：").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.chart_type = ttk.Combobox(tab, values=["bar", "line", "pie", "scatter"], state="readonly", width=15)
        self.chart_type.set("bar")
        self.chart_type.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(tab, text="图表标题：").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.chart_title = ttk.Entry(tab, width=30)
        self.chart_title.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(tab, text="X轴字段：").grid(row=4, column=0, padx=5, pady=5, sticky=tk.W)
        self.chart_x_field = ttk.Entry(tab, width=20)
        self.chart_x_field.grid(row=4, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(tab, text="Y轴字段：").grid(row=5, column=0, padx=5, pady=5, sticky=tk.W)
        self.chart_y_field = ttk.Entry(tab, width=20)
        self.chart_y_field.grid(row=5, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 按钮
        ttk.Button(tab, text="添加图表", command=self.add_chart).grid(row=2, column=2, padx=5, pady=5)
        ttk.Button(tab, text="删除图表", command=self.delete_chart).grid(row=3, column=2, padx=5, pady=5)
        ttk.Button(tab, text="清空图表", command=self.clear_charts).grid(row=4, column=2, padx=5, pady=5)
    
    def create_data_entry_tab(self):
        """创建数据填报标签页"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="数据填报")
        
        # 数据填报表格
        self.data_entry_tree = ttk.Treeview(tab, show="headings", height=10)
        self.data_entry_tree.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W+tk.E+tk.N+tk.S)
        
        # 垂直滚动条
        v_scrollbar = ttk.Scrollbar(tab, orient=tk.VERTICAL, command=self.data_entry_tree.yview)
        v_scrollbar.grid(row=0, column=2, sticky=tk.NS)
        self.data_entry_tree.configure(yscrollcommand=v_scrollbar.set)
        
        # 水平滚动条
        h_scrollbar = ttk.Scrollbar(tab, orient=tk.HORIZONTAL, command=self.data_entry_tree.xview)
        h_scrollbar.grid(row=1, column=0, columnspan=2, sticky=tk.EW)
        self.data_entry_tree.configure(xscrollcommand=h_scrollbar.set)
        
        # 数据操作按钮
        button_frame = ttk.Frame(tab)
        button_frame.grid(row=2, column=0, columnspan=3, pady=10, sticky=tk.W)
        
        ttk.Button(button_frame, text="新建表格", command=self.create_new_table).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="导入数据", command=self.import_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="导出数据", command=self.export_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="添加行", command=self.add_row).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="删除行", command=self.delete_row).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="验证数据", command=self.validate_data).pack(side=tk.LEFT, padx=5)
        
        # 数据验证结果
        self.validation_result = tk.Text(tab, width=80, height=5, wrap=tk.WORD)
        self.validation_result.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W+tk.E)
        
        # 初始化数据存储
        self.data_entry_data = []
        self.data_entry_columns = []
    
    def create_calculations_tab(self):
        """创建计算字段标签页"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="计算字段")
        
        # 计算字段列表
        self.calculations_list = ttk.Treeview(tab, columns=("column", "formula"), show="headings", height=5)
        self.calculations_list.heading("column", text="字段名")
        self.calculations_list.heading("formula", text="计算公式")
        self.calculations_list.grid(row=0, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W+tk.E)
        
        # 垂直滚动条
        scrollbar = ttk.Scrollbar(tab, orient=tk.VERTICAL, command=self.calculations_list.yview)
        scrollbar.grid(row=0, column=3, sticky=tk.NS)
        self.calculations_list.configure(yscrollcommand=scrollbar.set)
        
        # 计算字段设置
        ttk.LabelFrame(tab, text="计算字段设置").grid(row=1, column=0, columnspan=3, padx=5, pady=10, sticky=tk.W+tk.E)
        
        ttk.Label(tab, text="字段名：").grid(row=2, column=0, padx=5, pady=5, sticky=tk.W)
        self.calc_column = ttk.Entry(tab, width=20)
        self.calc_column.grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(tab, text="计算公式：").grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)
        self.calc_formula = ttk.Entry(tab, width=40)
        self.calc_formula.grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        ttk.Label(tab, text="（使用df['字段名']语法）").grid(row=3, column=2, padx=5, pady=5, sticky=tk.W)
        
        # 按钮
        ttk.Button(tab, text="添加计算字段", command=self.add_calculation).grid(row=4, column=1, padx=5, pady=5)
        ttk.Button(tab, text="删除计算字段", command=self.delete_calculation).grid(row=4, column=2, padx=5, pady=5)
        ttk.Button(tab, text="清空计算字段", command=self.clear_calculations).grid(row=5, column=1, padx=5, pady=5)
    
    def create_button_bar(self):
        """创建底部按钮栏"""
        button_bar = ttk.Frame(self.main_frame)
        button_bar.pack(side=tk.BOTTOM, fill=tk.X, pady=10)
        
        # 左对齐按钮
        left_buttons = ttk.Frame(button_bar)
        left_buttons.pack(side=tk.LEFT)
        
        ttk.Button(left_buttons, text="保存配置", command=self.save_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(left_buttons, text="加载配置", command=self.load_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(left_buttons, text="重置", command=self.reset_config).pack(side=tk.LEFT, padx=5)
        
        # 右对齐按钮
        right_buttons = ttk.Frame(button_bar)
        right_buttons.pack(side=tk.RIGHT)
        
        ttk.Button(right_buttons, text="退出", command=self.root.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(right_buttons, text="生成报表", command=self.generate_report).pack(side=tk.RIGHT, padx=5)
    
    def on_data_source_type_change(self, event):
        """数据源类型改变时的处理函数"""
        data_source_type = self.data_source_type.get()
        self.show_params_panel(data_source_type)
    
    def show_params_panel(self, data_source_type):
        """显示指定数据源类型的参数面板"""
        # 隐藏所有参数面板
        for frame in [self.excel_params_frame, self.csv_params_frame, self.sql_params_frame, self.api_params_frame]:
            frame.grid_forget()
        
        # 显示选定的参数面板
        if data_source_type == "excel":
            self.excel_params_frame.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        elif data_source_type == "csv":
            self.csv_params_frame.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        elif data_source_type == "sql":
            self.sql_params_frame.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        elif data_source_type == "api":
            self.api_params_frame.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
    
    def browse_file(self):
        """浏览文件"""
        filetypes = []
        data_source_type = self.data_source_type.get()
        
        if data_source_type == "excel":
            filetypes = [ ("Excel文件", "*.xlsx;*.xls"), ("所有文件", "*.*") ]
        elif data_source_type == "csv":
            filetypes = [ ("CSV文件", "*.csv"), ("所有文件", "*.*") ]
        
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            self.data_source_path.delete(0, tk.END)
            self.data_source_path.insert(0, filename)
    
    def browse_config_file(self):
        """浏览配置文件"""
        filename = filedialog.askopenfilename(filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")])
        if filename:
            self.config_file.delete(0, tk.END)
            self.config_file.insert(0, filename)
    
    def browse_output_dir(self):
        """浏览输出目录"""
        dirname = filedialog.askdirectory()
        if dirname:
            self.output_dir.delete(0, tk.END)
            self.output_dir.insert(0, dirname)
    
    def browse_template_file(self):
        """浏览模板文件"""
        filename = filedialog.askopenfilename(filetypes=[("模板文件", "*.html;*.jinja2"), ("所有文件", "*.*")])
        if filename:
            self.template_file.delete(0, tk.END)
            self.template_file.insert(0, filename)
    
    def add_chart(self):
        """添加图表"""
        chart_type = self.chart_type.get()
        title = self.chart_title.get()
        x_field = self.chart_x_field.get()
        y_field = self.chart_y_field.get()
        
        if not title or not x_field or not y_field:
            messagebox.showwarning("警告", "请填写完整的图表信息")
            return
        
        self.charts_list.insert("", tk.END, values=(chart_type, title, x_field, y_field))
        
        # 清空输入框
        self.chart_title.delete(0, tk.END)
        self.chart_x_field.delete(0, tk.END)
        self.chart_y_field.delete(0, tk.END)
    
    def delete_chart(self):
        """删除图表"""
        selected_item = self.charts_list.selection()
        if selected_item:
            self.charts_list.delete(selected_item)
        else:
            messagebox.showwarning("警告", "请先选择要删除的图表")
    
    def clear_charts(self):
        """清空图表列表"""
        self.charts_list.delete(*self.charts_list.get_children())
    
    def add_calculation(self):
        """添加计算字段"""
        column = self.calc_column.get()
        formula = self.calc_formula.get()
        
        if not column or not formula:
            messagebox.showwarning("警告", "请填写完整的计算字段信息")
            return
        
        self.calculations_list.insert("", tk.END, values=(column, formula))
        
        # 清空输入框
        self.calc_column.delete(0, tk.END)
        self.calc_formula.delete(0, tk.END)
    
    def delete_calculation(self):
        """删除计算字段"""
        selected_item = self.calculations_list.selection()
        if selected_item:
            self.calculations_list.delete(selected_item)
        else:
            messagebox.showwarning("警告", "请先选择要删除的计算字段")
    
    def clear_calculations(self):
        """清空计算字段列表"""
        self.calculations_list.delete(*self.calculations_list.get_children())
    
    def create_new_table(self):
        """创建新的数据填报表格"""
        # 创建一个简单的对话框来设置列名
        dialog = tk.Toplevel(self.root)
        dialog.title("新建表格")
        dialog.geometry("400x200")
        
        ttk.Label(dialog, text="列名（逗号分隔）：").pack(pady=10)
        
        columns_entry = ttk.Entry(dialog, width=40)
        columns_entry.pack(pady=5)
        columns_entry.insert(0, "日期,产品,数量,单价,金额")
        
        def on_ok():
            columns = [col.strip() for col in columns_entry.get().split(",") if col.strip()]
            if not columns:
                messagebox.showerror("错误", "请输入至少一个列名")
                return
            
            # 清空现有表格
            for col in self.data_entry_tree.get_children():
                self.data_entry_tree.delete(col)
            for col in self.data_entry_tree["columns"]:
                self.data_entry_tree.delete(col)
            
            # 设置新列
            self.data_entry_tree["columns"] = columns
            self.data_entry_columns = columns
            
            for col in columns:
                self.data_entry_tree.heading(col, text=col)
                self.data_entry_tree.column(col, width=100)
            
            # 清空数据
            self.data_entry_data = []
            
            dialog.destroy()
        
        ttk.Button(dialog, text="确定", command=on_ok).pack(side=tk.LEFT, padx=10, pady=10)
        ttk.Button(dialog, text="取消", command=dialog.destroy).pack(side=tk.RIGHT, padx=10, pady=10)
    
    def import_data(self):
        """导入数据"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("CSV文件", "*.csv"), ("JSON文件", "*.json")]
        )
        
        if not file_path:
            return
        
        try:
            import pandas as pd
            
            # 读取文件
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            elif file_path.endswith('.json'):
                df = pd.read_json(file_path)
            else:  # Excel
                df = pd.read_excel(file_path)
            
            # 清空现有表格
            for col in self.data_entry_tree.get_children():
                self.data_entry_tree.delete(col)
            for col in self.data_entry_tree["columns"]:
                self.data_entry_tree.delete(col)
            
            # 设置列
            columns = list(df.columns)
            self.data_entry_tree["columns"] = columns
            self.data_entry_columns = columns
            
            for col in columns:
                self.data_entry_tree.heading(col, text=col)
                self.data_entry_tree.column(col, width=100)
            
            # 填充数据
            self.data_entry_data = df.to_dict('records')
            
            for i, row in enumerate(self.data_entry_data):
                values = [str(row[col]) if pd.notna(row[col]) else '' for col in columns]
                self.data_entry_tree.insert('', tk.END, values=values)
            
            self.status_var.set(f"成功导入 {len(self.data_entry_data)} 行数据")
            
        except Exception as e:
            messagebox.showerror("导入错误", f"导入数据失败：{str(e)}")
            logger.error(f"数据导入失败: {str(e)}")
    
    def export_data(self):
        """导出数据"""
        if not self.data_entry_columns:
            messagebox.showwarning("警告", "没有数据可以导出")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("CSV文件", "*.csv"), ("JSON文件", "*.json")]
        )
        
        if not file_path:
            return
        
        try:
            import pandas as pd
            
            # 创建DataFrame
            df = pd.DataFrame(self.data_entry_data)
            
            # 导出文件
            if file_path.endswith('.csv'):
                df.to_csv(file_path, index=False)
            elif file_path.endswith('.json'):
                df.to_json(file_path, orient='records', indent=2)
            else:  # Excel
                df.to_excel(file_path, index=False)
            
            self.status_var.set(f"成功导出 {len(self.data_entry_data)} 行数据")
            
        except Exception as e:
            messagebox.showerror("导出错误", f"导出数据失败：{str(e)}")
            logger.error(f"数据导出失败: {str(e)}")
    
    def add_row(self):
        """添加新行"""
        if not self.data_entry_columns:
            messagebox.showwarning("警告", "请先创建表格或导入数据")
            return
        
        # 创建空行数据
        new_row = {col: '' for col in self.data_entry_columns}
        self.data_entry_data.append(new_row)
        
        # 在表格中显示新行
        values = ['' for _ in self.data_entry_columns]
        self.data_entry_tree.insert('', tk.END, values=values)
    
    def delete_row(self):
        """删除选中的行"""
        selected_item = self.data_entry_tree.selection()
        if not selected_item:
            messagebox.showwarning("警告", "请选择要删除的行")
            return
        
        # 获取选中行的索引
        index = self.data_entry_tree.index(selected_item[0])
        
        # 删除数据和表格行
        del self.data_entry_data[index]
        self.data_entry_tree.delete(selected_item[0])
    
    def validate_data(self):
        """验证数据"""
        if not self.data_entry_columns:
            messagebox.showwarning("警告", "没有数据需要验证")
            return
        
        validation_results = []
        
        for i, row in enumerate(self.data_entry_data):
            row_errors = []
            
            # 检查必填字段
            for col in self.data_entry_columns:
                if not str(row[col]).strip():
                    row_errors.append(f"{col} 不能为空")
            
            # 简单的数字验证（如果列名包含数量、金额、单价等关键词）
            for col, value in row.items():
                if col in ["数量", "金额", "单价", "价格", "数值", "amount", "price", "quantity"]:
                    try:
                        float(value)
                    except (ValueError, TypeError):
                        if str(value).strip():
                            row_errors.append(f"{col} 必须是数字")
            
            # 日期验证（如果列名包含日期关键词）
            for col, value in row.items():
                if col in ["日期", "时间", "date", "time", "datetime"]:
                    try:
                        if str(value).strip():
                            pd.to_datetime(value)
                    except (ValueError, TypeError):
                        row_errors.append(f"{col} 不是有效的日期格式")
            
            if row_errors:
                validation_results.append(f"行 {i+1} 错误: {'; '.join(row_errors)}")
        
        # 显示验证结果
        self.validation_result.delete(1.0, tk.END)
        if validation_results:
            self.validation_result.insert(tk.END, "数据验证失败：\n\n")
            for result in validation_results:
                self.validation_result.insert(tk.END, f"• {result}\n")
            self.validation_result.tag_add("error", 1.0, tk.END)
            self.validation_result.tag_config("error", foreground="red")
            self.status_var.set(f"数据验证失败，发现 {len(validation_results)} 个错误")
        else:
            self.validation_result.insert(tk.END, "数据验证成功！所有数据格式正确。")
            self.validation_result.tag_add("success", 1.0, tk.END)
            self.validation_result.tag_config("success", foreground="green")
            self.status_var.set("数据验证成功")
        
        # 更新表格中的数据
        for item in self.data_entry_tree.get_children():
            self.data_entry_tree.delete(item)
        
        for row in self.data_entry_data:
            values = [str(row[col]) for col in self.data_entry_columns]
            self.data_entry_tree.insert('', tk.END, values=values)
    
    def get_report_config(self):
        """获取报表配置"""
        # 基础配置
        report_name = self.report_name.get()
        data_source_type = self.data_source_type.get()
        data_source_path = self.data_source_path.get()
        
        # 输出格式
        output_formats = [fmt for fmt, var in self.output_formats.items() if var.get()]
        
        # 调度和邮件
        schedule = self.cron_expression.get() or None
        recipients = [r.strip() for r in self.email_recipients.get().split(",") if r.strip()]
        
        # 模板文件
        template_path = self.template_file.get() or None
        
        # 参数配置
        parameters = self.get_data_source_parameters()
        
        # 图表配置
        charts = []
        for item in self.charts_list.get_children():
            chart_type, title, x_field, y_field = self.charts_list.item(item)["values"]
            charts.append({
                "type": chart_type,
                "title": title,
                "x_field": x_field,
                "y_field": y_field
            })
        
        # 计算字段配置
        calculations = []
        for item in self.calculations_list.get_children():
            column, formula = self.calculations_list.item(item)["values"]
            calculations.append({
                "column": column,
                "formula": formula
            })
        
        # 获取模板类型
        template_type = self.template_type.get() if self.template_type.get() != 'default' else None
        
        # 创建配置对象
        return ReportConfig(
            report_name=report_name,
            template_type=template_type,
            data_source_type=data_source_type,
            data_source_path=data_source_path,
            output_format=output_formats,
            schedule=schedule,
            recipients=recipients,
            template_path=template_path,
            filters={},  # 简化处理，未实现过滤功能
            calculations=calculations,
            charts=charts,
            parameters=parameters
        )
    
    def get_data_source_parameters(self):
        """获取数据源参数"""
        data_source_type = self.data_source_type.get()
        parameters = {}
        
        if data_source_type == "excel":
            if self.excel_sheet.get():
                parameters["sheet_name"] = self.excel_sheet.get()
            if self.excel_chunksize.get():
                parameters["chunksize"] = int(self.excel_chunksize.get())
        
        elif data_source_type == "csv":
            parameters["encoding"] = self.csv_encoding.get()
            parameters["separator"] = self.csv_separator.get()
            if self.csv_chunksize.get():
                parameters["chunksize"] = int(self.csv_chunksize.get())
        
        elif data_source_type == "sql":
            parameters["query"] = self.sql_query.get("1.0", tk.END).strip()
        
        elif data_source_type == "api":
            parameters["method"] = self.api_method.get()
            try:
                if self.api_headers.get("1.0", tk.END).strip():
                    parameters["headers"] = json.loads(self.api_headers.get("1.0", tk.END).strip())
                if self.api_params.get("1.0", tk.END).strip():
                    parameters["params"] = json.loads(self.api_params.get("1.0", tk.END).strip())
            except json.JSONDecodeError:
                messagebox.showerror("错误", "请求头或查询参数格式错误，必须是有效的JSON格式")
                return None
            
            if self.api_response_key.get():
                parameters["response_key"] = self.api_response_key.get()
        
        return parameters
    
    def save_config(self):
        """保存配置到文件（支持JSON和YAML）"""
        config = self.get_report_config()
        
        # 转换为字典
        config_dict = {
            "report_name": config.report_name,
            "data_source_type": config.data_source_type,
            "data_source_path": config.data_source_path,
            "output_format": config.output_format,
            "schedule": config.schedule,
            "recipients": config.recipients,
            "template_path": config.template_path,
            "filters": config.filters,
            "calculations": config.calculations,
            "charts": config.charts,
            "parameters": config.parameters
        }
        
        # 保存到文件
        filename = filedialog.asksaveasfilename(filetypes=[("JSON文件", "*.json"), ("YAML文件", "*.yaml"), ("YAML文件", "*.yml"), ("所有文件", "*.*")], defaultextension=".json")
        if filename:
            try:
                file_ext = os.path.splitext(filename)[1].lower()
                
                if file_ext == '.json':
                    with open(filename, "w", encoding="utf-8") as f:
                        json.dump(config_dict, f, indent=2, ensure_ascii=False)
                elif file_ext in ['.yaml', '.yml']:
                    with open(filename, "w", encoding="utf-8") as f:
                        yaml.dump(config_dict, f, default_flow_style=False, allow_unicode=True)
                else:
                    messagebox.showerror("错误", "不支持的文件格式")
                    return
                    
                messagebox.showinfo("成功", "配置已保存")
                self.config_file.delete(0, tk.END)
                self.config_file.insert(0, filename)
            except Exception as e:
                messagebox.showerror("错误", f"保存配置失败：{e}")
    
    def load_config(self):
        """从文件加载配置（支持JSON和YAML）"""
        filename = filedialog.askopenfilename(filetypes=[("JSON文件", "*.json"), ("YAML文件", "*.yaml"), ("YAML文件", "*.yml"), ("所有文件", "*.*")])
        if filename:
            try:
                file_ext = os.path.splitext(filename)[1].lower()
                
                if file_ext == '.json':
                    with open(filename, "r", encoding="utf-8") as f:
                        config_dict = json.load(f)
                elif file_ext in ['.yaml', '.yml']:
                    with open(filename, "r", encoding="utf-8") as f:
                        config_dict = yaml.safe_load(f)
                else:
                    messagebox.showerror("错误", "不支持的文件格式")
                    return
                
                # 加载配置
                self.report_name.delete(0, tk.END)
                self.report_name.insert(0, config_dict.get("report_name", "自动化报表"))
                
                self.data_source_type.set(config_dict.get("data_source_type", "excel"))
                self.data_source_path.delete(0, tk.END)
                self.data_source_path.insert(0, config_dict.get("data_source_path", ""))
                
                # 加载输出格式
                for fmt, var in self.output_formats.items():
                    var.set(fmt in config_dict.get("output_format", ["excel"]))
                
                # 加载调度和邮件
                self.cron_expression.delete(0, tk.END)
                self.cron_expression.insert(0, config_dict.get("schedule", ""))
                self.email_recipients.delete(0, tk.END)
                self.email_recipients.insert(0, ",".join(config_dict.get("recipients", [])))
                
                # 加载模板文件
                self.template_file.delete(0, tk.END)
                self.template_file.insert(0, config_dict.get("template_path", ""))
                
                # 加载参数
                parameters = config_dict.get("parameters", {})
                self.load_data_source_parameters(parameters)
                
                # 加载图表
                self.clear_charts()
                for chart in config_dict.get("charts", []):
                    self.charts_list.insert("", tk.END, values=(
                        chart.get("type", "bar"),
                        chart.get("title", ""),
                        chart.get("x_field", ""),
                        chart.get("y_field", "")
                    ))
                
                # 加载计算字段
                self.clear_calculations()
                for calc in config_dict.get("calculations", []):
                    self.calculations_list.insert("", tk.END, values=(
                        calc.get("column", ""),
                        calc.get("formula", "")
                    ))
                
                # 更新当前配置文件路径
                self.config_file.delete(0, tk.END)
                self.config_file.insert(0, filename)
                
                messagebox.showinfo("成功", "配置已加载")
                
            except Exception as e:
                messagebox.showerror("错误", f"加载配置失败：{e}")
    
    def load_data_source_parameters(self, parameters):
        """加载数据源参数"""
        data_source_type = self.data_source_type.get()
        
        if data_source_type == "excel":
            self.excel_sheet.delete(0, tk.END)
            self.excel_sheet.insert(0, parameters.get("sheet_name", ""))
            self.excel_chunksize.delete(0, tk.END)
            self.excel_chunksize.insert(0, parameters.get("chunksize", ""))
        
        elif data_source_type == "csv":
            self.csv_encoding.set(parameters.get("encoding", "utf-8"))
            self.csv_separator.delete(0, tk.END)
            self.csv_separator.insert(0, parameters.get("separator", ","))
            self.csv_chunksize.delete(0, tk.END)
            self.csv_chunksize.insert(0, parameters.get("chunksize", ""))
        
        elif data_source_type == "sql":
            self.sql_connection.delete(0, tk.END)
            self.sql_connection.insert(0, parameters.get("connection", ""))
            self.sql_query.delete("1.0", tk.END)
            self.sql_query.insert("1.0", parameters.get("query", ""))
        
        elif data_source_type == "api":
            self.api_method.set(parameters.get("method", "GET"))
            self.api_headers.delete("1.0", tk.END)
            self.api_headers.insert("1.0", json.dumps(parameters.get("headers", {}), indent=2) if "headers" in parameters else "")
            self.api_params.delete("1.0", tk.END)
            self.api_params.insert("1.0", json.dumps(parameters.get("params", {}), indent=2) if "params" in parameters else "")
            self.api_response_key.delete(0, tk.END)
            self.api_response_key.insert(0, parameters.get("response_key", ""))
        
        # 显示相应的参数面板
        self.show_params_panel(data_source_type)
    
    def apply_config_template(self):
        """应用配置模板"""
        template_name = self.config_templates.get()
        
        # 根据模板名称设置默认值
        if template_name == "销售报表":
            self.report_name.delete(0, tk.END)
            self.report_name.insert(0, "销售业绩报表")
            self.data_source_type.set("excel")
            self.show_params_panel("excel")
            
            # 设置输出格式
            for fmt, var in self.output_formats.items():
                var.set(False)
            self.output_formats["excel"].set(True)
            self.output_formats["email"].set(True)
            
            messagebox.showinfo("成功", "销售报表模板已应用")
        elif template_name == "财务报表":
            self.report_name.delete(0, tk.END)
            self.report_name.insert(0, "财务分析报表")
            self.data_source_type.set("excel")
            self.show_params_panel("excel")
            
            # 设置输出格式
            for fmt, var in self.output_formats.items():
                var.set(False)
            self.output_formats["pdf"].set(True)
            self.output_formats["excel"].set(True)
            
            messagebox.showinfo("成功", "财务报表模板已应用")
        elif template_name == "库存报表":
            self.report_name.delete(0, tk.END)
            self.report_name.insert(0, "库存状态报表")
            self.data_source_type.set("sql")
            self.show_params_panel("sql")
            
            # 设置默认SQL查询
            default_query = "SELECT product_name, category, stock_quantity, unit_price FROM products"
            self.sql_query.delete("1.0", tk.END)
            self.sql_query.insert("1.0", default_query)
            
            # 设置输出格式
            for fmt, var in self.output_formats.items():
                var.set(False)
            self.output_formats["excel"].set(True)
            self.output_formats["html"].set(True)
            
            messagebox.showinfo("成功", "库存报表模板已应用")
        else:
            # 空模板不做任何操作
            pass
    
    def reset_config(self):
        """重置配置"""
        if messagebox.askyesno("确认", "确定要重置所有配置吗？"):
            # 重置所有输入框
            self.report_name.delete(0, tk.END)
            self.report_name.insert(0, "自动化报表")
            
            self.data_source_type.set("excel")
            self.data_source_path.delete(0, tk.END)
            
            for fmt, var in self.output_formats.items():
                var.set(True if fmt == "excel" else False)
            
            self.cron_expression.delete(0, tk.END)
            self.email_recipients.delete(0, tk.END)
            
            self.template_file.delete(0, tk.END)
            
            # 重置数据源参数
            self.excel_sheet.delete(0, tk.END)
            self.excel_chunksize.delete(0, tk.END)
            
            self.csv_encoding.set("utf-8")
            self.csv_separator.delete(0, tk.END)
            self.csv_separator.insert(0, ",")
            self.csv_chunksize.delete(0, tk.END)
            
            self.sql_connection.delete(0, tk.END)
            self.sql_query.delete("1.0", tk.END)
            
            self.api_method.set("GET")
            self.api_headers.delete("1.0", tk.END)
            self.api_params.delete("1.0", tk.END)
            self.api_response_key.delete(0, tk.END)
            
            # 清空图表和计算字段
            self.clear_charts()
            self.clear_calculations()
            
            # 显示Excel参数面板
            self.show_params_panel("excel")
            
            messagebox.showinfo("成功", "配置已重置")
    
    def generate_report(self):
        """生成报表"""
        # 验证输入
        if not self.data_source_path.get():
            messagebox.showwarning("警告", "请选择数据源文件路径")
            return
        
        if not self.output_dir.get():
            messagebox.showwarning("警告", "请选择输出目录")
            return
        
        # 更新配置管理器
        self.config_manager.config["output_dir"] = self.output_dir.get()
        
        # 获取配置
        config = self.get_report_config()
        
        # 检查输出格式
        if not config.output_format:
            messagebox.showwarning("警告", "请至少选择一种输出格式")
            return
        
        # 异步生成报表
        self.status_var.set("正在生成报表...")
        self.root.update()
        
        # 创建线程执行报表生成
        thread = threading.Thread(target=self._generate_report_thread, args=(config,))
        thread.daemon = True
        thread.start()
    
    def _generate_report_thread(self, config):
        """报表生成线程"""
        try:
            # 创建引擎并运行
            engine = AutoReportEngine(config)
            result = engine.run()
            
            # 显示结果
            self.root.after(0, lambda: self.show_report_result(result))
            
        except Exception as e:
            logger.error(f"生成报表失败: {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"生成报表失败：{e}"))
            self.root.after(0, lambda: self.status_var.set("生成报表失败"))
    
    def show_report_result(self, result):
        """显示报表生成结果"""
        # 显示成功信息
        success_msg = "报表生成成功！生成的文件：\n\n"
        for fmt, path in result.items():
            success_msg += f"{fmt}: {path}\n"
        
        messagebox.showinfo("成功", success_msg)
        self.status_var.set("准备就绪")
    
    def create_schedule_tab(self):
        """创建调度管理标签页"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="调度管理")
        
        # 调度任务列表
        ttk.Label(tab, text="调度任务列表：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.schedule_list = ttk.Treeview(tab, columns=("id", "report_name", "cron", "status"), show="headings", height=8)
        self.schedule_list.heading("id", text="任务ID")
        self.schedule_list.heading("report_name", text="报表名称")
        self.schedule_list.heading("cron", text="Cron表达式")
        self.schedule_list.heading("status", text="状态")
        self.schedule_list.grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W+tk.E)
        
        # 任务操作按钮
        button_frame = ttk.Frame(tab)
        button_frame.grid(row=2, column=0, columnspan=3, pady=10)
        
        ttk.Button(button_frame, text="添加当前配置为调度任务", command=self.add_schedule_task).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="删除选中任务", command=self.delete_schedule_task).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="刷新任务列表", command=self.update_schedule_list).pack(side=tk.LEFT, padx=5)
        
        # 调度器控制
        control_frame = ttk.LabelFrame(tab, text="调度器控制")
        control_frame.grid(row=3, column=0, columnspan=3, padx=5, pady=10, sticky=tk.W+tk.E)
        
        ttk.Button(control_frame, text="启动调度器", command=self.start_scheduler).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(control_frame, text="停止调度器", command=self.stop_scheduler).pack(side=tk.LEFT, padx=5, pady=5)
        
        # 初始化时更新任务列表
        self.update_schedule_list()
    
    def update_schedule_list(self):
        """更新调度任务列表"""
        # 清空现有列表
        for item in self.schedule_list.get_children():
            self.schedule_list.delete(item)
        
        # 获取所有调度任务
        tasks = schedule_manager.get_all_tasks()
        
        # 添加任务到列表
        for task in tasks:
            status = "运行中" if schedule_manager.is_scheduler_running() else "已停止"
            self.schedule_list.insert("", tk.END, values=(task['id'], task['report_name'], task['cron_expression'], status))
    
    def add_schedule_task(self):
        """添加当前配置为调度任务"""
        # 获取当前报表配置
        report_name = self.report_name.get()
        cron_expr = self.cron_expression.get()
        
        if not report_name:
            messagebox.showerror("错误", "请输入报表名称")
            return
        
        if not cron_expr:
            messagebox.showerror("错误", "请输入Cron表达式")
            return
        
        try:
            # 获取完整的报表配置
            report_config = self.get_report_config()
            report_config.schedule = cron_expr
            
            # 添加调度任务
            task_id = schedule_manager.add_task(report_name, cron_expr, report_config)
            
            if task_id:
                messagebox.showinfo("成功", "调度任务添加成功")
                self.update_schedule_list()
            else:
                messagebox.showerror("错误", "调度任务添加失败")
        except Exception as e:
            messagebox.showerror("错误", f"添加调度任务失败: {str(e)}")
    
    def delete_schedule_task(self):
        """删除选中的调度任务"""
        selected_item = self.schedule_list.selection()
        if not selected_item:
            messagebox.showwarning("警告", "请选择要删除的任务")
            return
        
        # 获取选中任务的ID
        task_id = self.schedule_list.item(selected_item[0])['values'][0]
        
        if messagebox.askyesno("确认", "确定要删除选中的调度任务吗？"):
            # 删除任务
            if schedule_manager.remove_task(task_id):
                messagebox.showinfo("成功", "调度任务删除成功")
                self.update_schedule_list()
            else:
                messagebox.showerror("错误", "调度任务删除失败")
    
    def start_scheduler(self):
        """启动调度器"""
        try:
            if schedule_manager.start_scheduler():
                self.status_var.set("调度器已启动")
                messagebox.showinfo("成功", "调度器已启动")
                self.update_schedule_list()
            else:
                messagebox.showwarning("警告", "调度器已经在运行中")
        except Exception as e:
            self.status_var.set(f"启动调度器失败：{str(e)}")
            messagebox.showerror("错误", f"启动调度器失败: {str(e)}")
    
    def update_schedule_list(self):
        """更新调度任务列表"""
        # 清空现有列表
        for item in self.schedule_list.get_children():
            self.schedule_list.delete(item)
        
        # 获取所有调度任务
        tasks = schedule_manager.get_all_tasks()
        
        # 添加到列表中
        for task in tasks:
            status = "运行中" if task.get('running', False) else "已停止"
            self.schedule_list.insert("", tk.END, values=(task['task_id'], task['report_name'], task['schedule'], status))
    
    def add_schedule_task(self):
        """添加当前配置为调度任务"""
        # 获取当前报表配置
        config = self.get_report_config()
        
        if not config['schedule']:
            messagebox.showwarning("警告", "请先在报表配置中设置Cron表达式")
            return
        
        # 检查是否有报表名称
        if not config['name']:
            messagebox.showwarning("警告", "请先在报表配置中设置报表名称")
            return
        
        try:
            # 添加调度任务
            task_id = schedule_manager.add_task(
                report_name=config['name'],
                report_config=config,
                schedule=config['schedule'],
                email_recipients=config['email_recipients']
            )
            
            self.status_var.set(f"调度任务添加成功！任务ID：{task_id}")
            messagebox.showinfo("成功", f"调度任务已添加！\n任务ID：{task_id}")
            
            # 更新任务列表
            self.update_schedule_list()
        except Exception as e:
            self.status_var.set(f"添加调度任务失败：{str(e)}")
            messagebox.showerror("失败", f"添加调度任务失败：\n{str(e)}")
    
    def delete_schedule_task(self):
        """删除选中的调度任务"""
        selected_item = self.schedule_list.selection()
        if not selected_item:
            messagebox.showwarning("警告", "请先选择要删除的调度任务")
            return
        
        # 获取任务ID
        task_id = self.schedule_list.item(selected_item[0])['values'][0]
        
        try:
            # 删除调度任务
            schedule_manager.remove_task(task_id)
            
            self.status_var.set(f"调度任务删除成功！任务ID：{task_id}")
            messagebox.showinfo("成功", f"调度任务已删除！\n任务ID：{task_id}")
            
            # 更新任务列表
            self.update_schedule_list()
        except Exception as e:
            self.status_var.set(f"删除调度任务失败：{str(e)}")
            messagebox.showerror("失败", f"删除调度任务失败：\n{str(e)}")
    
    def stop_scheduler(self):
        """停止调度器"""
        try:
            if schedule_manager.stop_scheduler():
                self.status_var.set("调度器已停止")
                messagebox.showinfo("成功", "调度器已停止")
                
                # 更新任务列表状态
                self.update_schedule_list()
            else:
                messagebox.showwarning("警告", "调度器已经停止")
        except Exception as e:
            self.status_var.set(f"停止调度器失败：{str(e)}")
            messagebox.showerror("失败", f"停止调度器失败：\n{str(e)}")

def main():
    """主函数"""
    root = tk.Tk()
    app = ReportGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()