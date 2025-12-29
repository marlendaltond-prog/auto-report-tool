#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GUI报表配置管理组件
负责报表配置相关的UI界面和逻辑
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, Any, List, Optional
import os
import json
import logging

class ReportConfigManager:
    """报表配置管理器"""
    
    def __init__(self, parent_notebook):
        self.parent_notebook = parent_notebook
        self.report_config = {}  # 存储报表配置
        
        # 设置日志记录
        self.logger = logging.getLogger(__name__)
        
        # 创建报表配置标签页
        self.tab = ttk.Frame(self.parent_notebook)
        self.parent_notebook.add(self.tab, text="报表配置")
        
        self._create_widgets()
        
    def _create_widgets(self):
        """创建报表配置界面"""
        # 基本信息框架
        basic_info_frame = ttk.LabelFrame(self.tab, text="基本信息", padding="10")
        basic_info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 报表名称
        ttk.Label(basic_info_frame, text="报表名称：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.report_name = ttk.Entry(basic_info_frame, width=30)
        self.report_name.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 报表描述
        ttk.Label(basic_info_frame, text="报表描述：").grid(row=1, column=0, padx=5, pady=5, sticky=tk.NW)
        self.report_desc = tk.Text(basic_info_frame, width=50, height=3)
        self.report_desc.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 输出配置框架
        output_config_frame = ttk.LabelFrame(self.tab, text="输出配置", padding="10")
        output_config_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 输出格式
        ttk.Label(output_config_frame, text="输出格式：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.output_format = ttk.Combobox(output_config_frame, values=["excel", "pdf", "html", "json"], state="readonly", width=15)
        self.output_format.set("excel")
        self.output_format.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        self.output_format.bind("<<ComboboxSelected>>", self._on_output_format_change)
        
        # 输出路径
        ttk.Label(output_config_frame, text="输出路径：").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        path_frame = ttk.Frame(output_config_frame)
        path_frame.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W+tk.E)
        self.output_path = ttk.Entry(path_frame, width=40)
        self.output_path.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.browse_button = ttk.Button(path_frame, text="浏览", command=self._browse_output_path)
        self.browse_button.pack(side=tk.RIGHT, padx=(5,0))
        
        # 输出参数框架
        self.output_params_frame = ttk.LabelFrame(output_config_frame, text="输出参数", padding="10")
        self.output_params_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=10, sticky=tk.W+tk.E)
        
        # 为不同输出格式创建参数面板
        self._create_excel_params()
        self._create_pdf_params()
        self._create_html_params()
        self._create_json_params()
        
        # 初始显示Excel参数
        self._show_output_params_panel("excel")
        
        # 样式配置框架
        style_config_frame = ttk.LabelFrame(self.tab, text="样式配置", padding="10")
        style_config_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 样式主题
        ttk.Label(style_config_frame, text="样式主题：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.style_theme = ttk.Combobox(style_config_frame, values=["default", "minimal", "business", "colorful"], state="readonly", width=15)
        self.style_theme.set("default")
        self.style_theme.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 自定义CSS
        ttk.Label(style_config_frame, text="自定义CSS：").grid(row=1, column=0, padx=5, pady=5, sticky=tk.NW)
        self.custom_css = tk.Text(style_config_frame, width=50, height=5)
        self.custom_css.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 数据配置框架
        data_config_frame = ttk.LabelFrame(self.tab, text="数据配置", padding="10")
        data_config_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 数据清洗选项
        self.clean_data_var = tk.BooleanVar(value=True)
        clean_data_cb = ttk.Checkbutton(data_config_frame, text="自动清洗数据", variable=self.clean_data_var)
        clean_data_cb.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        # 数据验证选项
        self.validate_data_var = tk.BooleanVar(value=True)
        validate_data_cb = ttk.Checkbutton(data_config_frame, text="验证数据完整性", variable=self.validate_data_var)
        validate_data_cb.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 缓存数据选项
        self.cache_data_var = tk.BooleanVar(value=False)
        cache_data_cb = ttk.Checkbutton(data_config_frame, text="启用数据缓存", variable=self.cache_data_var)
        cache_data_cb.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        
        # 操作按钮框架
        button_frame = ttk.Frame(self.tab)
        button_frame.pack(fill=tk.X, padx=5, pady=10)
        
        self.save_config_button = ttk.Button(button_frame, text="保存配置", command=self._save_config)
        self.save_config_button.pack(side=tk.LEFT, padx=5)
        
        self.load_config_button = ttk.Button(button_frame, text="加载配置", command=self._load_config)
        self.load_config_button.pack(side=tk.LEFT, padx=5)
        
        self.reset_config_button = ttk.Button(button_frame, text="重置配置", command=self._reset_config)
        self.reset_config_button.pack(side=tk.LEFT, padx=5)
        
    def _create_excel_params(self):
        """创建Excel输出参数面板"""
        self.excel_output_frame = ttk.Frame(self.output_params_frame)
        
        ttk.Label(self.excel_output_frame, text="工作表名：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.excel_sheet_name = ttk.Entry(self.excel_output_frame, width=20)
        self.excel_sheet_name.insert(0, "报表数据")
        self.excel_sheet_name.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 字体大小
        ttk.Label(self.excel_output_frame, text="字体大小：").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.excel_font_size = ttk.Spinbox(self.excel_output_frame, from_=8, to=24, width=5)
        self.excel_font_size.set(11)
        self.excel_font_size.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        
    def _create_pdf_params(self):
        """创建PDF输出参数面板"""
        self.pdf_output_frame = ttk.Frame(self.output_params_frame)
        
        ttk.Label(self.pdf_output_frame, text="页面大小：").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.pdf_page_size = ttk.Combobox(self.pdf_output_frame, values=["A4", "Letter", "Legal"], state="readonly", width=15)
        self.pdf_page_size.set("A4")
        self.pdf_page_size.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 页面方向
        ttk.Label(self.pdf_output_frame, text="页面方向：").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.pdf_orientation = ttk.Combobox(self.pdf_output_frame, values=["portrait", "landscape"], state="readonly", width=15)
        self.pdf_orientation.set("portrait")
        self.pdf_orientation.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        
        # 包含图表
        self.pdf_include_charts_var = tk.BooleanVar(value=True)
        pdf_charts_cb = ttk.Checkbutton(self.pdf_output_frame, text="包含图表", variable=self.pdf_include_charts_var)
        pdf_charts_cb.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
        
    def _create_html_params(self):
        """创建HTML输出参数面板"""
        self.html_output_frame = ttk.Frame(self.output_params_frame)
        
        # 包含导航
        self.html_include_nav_var = tk.BooleanVar(value=True)
        html_nav_cb = ttk.Checkbutton(self.html_output_frame, text="包含导航菜单", variable=self.html_include_nav_var)
        html_nav_cb.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        # 包含搜索
        self.html_include_search_var = tk.BooleanVar(value=False)
        html_search_cb = ttk.Checkbutton(self.html_output_frame, text="包含搜索功能", variable=self.html_include_search_var)
        html_search_cb.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 响应式设计
        self.html_responsive_var = tk.BooleanVar(value=True)
        html_responsive_cb = ttk.Checkbutton(self.html_output_frame, text="响应式设计", variable=self.html_responsive_var)
        html_responsive_cb.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        
    def _create_json_params(self):
        """创建JSON输出参数面板"""
        self.json_output_frame = ttk.Frame(self.output_params_frame)
        
        # 缩进格式
        self.json_indent_var = tk.BooleanVar(value=True)
        json_indent_cb = ttk.Checkbutton(self.json_output_frame, text="格式化缩进", variable=self.json_indent_var)
        json_indent_cb.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        
        # 包含元数据
        self.json_include_meta_var = tk.BooleanVar(value=True)
        json_meta_cb = ttk.Checkbutton(self.json_output_frame, text="包含元数据", variable=self.json_include_meta_var)
        json_meta_cb.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 数据压缩
        self.json_compress_var = tk.BooleanVar(value=False)
        json_compress_cb = ttk.Checkbutton(self.json_output_frame, text="压缩数据", variable=self.json_compress_var)
        json_compress_cb.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        
    def _browse_output_path(self):
        """浏览选择输出路径"""
        try:
            output_format = self.output_format.get()
            if output_format == "excel":
                filetypes = [("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
                initial_file = "report.xlsx"
            elif output_format == "pdf":
                filetypes = [("PDF files", "*.pdf"), ("All files", "*.*")]
                initial_file = "report.pdf"
            elif output_format == "html":
                filetypes = [("HTML files", "*.html *.htm"), ("All files", "*.*")]
                initial_file = "report.html"
            elif output_format == "json":
                filetypes = [("JSON files", "*.json"), ("All files", "*.*")]
                initial_file = "report.json"
            else:
                filetypes = [("All files", "*.*")]
                initial_file = "output"
            
            filename = filedialog.asksaveasfilename(
                title="选择输出文件",
                filetypes=filetypes,
                initialvalue=initial_file,
                defaultextension=output_format
            )
            
            if filename:
                self.output_path.delete(0, tk.END)
                self.output_path.insert(0, filename)
                
        except Exception as e:
            self.logger.error(f"浏览输出路径时出错: {e}")
            messagebox.showerror("错误", f"浏览输出路径时出错: {e}")
            
    def _validate_config(self, config: Dict[str, Any]) -> tuple[bool, str]:
        """验证配置是否有效"""
        try:
            # 验证基本信息
            basic_info = config.get('basic_info', {})
            if not basic_info.get('name', '').strip():
                return False, "报表名称不能为空"
            
            # 验证输出配置
            output_config = config.get('output_config', {})
            output_path = output_config.get('path', '').strip()
            if not output_path:
                return False, "输出路径不能为空"
            
            # 检查路径是否有效
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                try:
                    os.makedirs(output_dir, exist_ok=True)
                except PermissionError:
                    return False, f"没有权限创建目录: {output_dir}"
                except OSError as e:
                    return False, f"创建目录失败: {e}"
            
            # 验证输出格式特定参数
            output_format = output_config.get('format', 'excel')
            if output_format == "excel":
                font_size = output_config.get('font_size')
                if font_size and (font_size < 8 or font_size > 24):
                    return False, "Excel字体大小必须在8-24之间"
            
            return True, ""
            
        except Exception as e:
            self.logger.error(f"配置验证时出错: {e}")
            return False, f"配置验证时出错: {e}"
            
    def _on_output_format_change(self, event=None):
        """输出格式改变事件处理"""
        try:
            output_format = self.output_format.get()
            self._show_output_params_panel(output_format)
        except Exception as e:
            self.logger.error(f"输出格式改变时出错: {e}")
            messagebox.showerror("错误", f"输出格式改变时出错: {e}")
        
    def _show_output_params_panel(self, output_format: str):
        """显示对应的输出参数面板"""
        try:
            # 隐藏所有输出参数面板
            for frame in [self.excel_output_frame, self.pdf_output_frame, 
                         self.html_output_frame, self.json_output_frame]:
                frame.pack_forget()
                
            # 显示对应的参数面板
            if output_format == "excel":
                self.excel_output_frame.pack(fill=tk.X)
            elif output_format == "pdf":
                self.pdf_output_frame.pack(fill=tk.X)
            elif output_format == "html":
                self.html_output_frame.pack(fill=tk.X)
            elif output_format == "json":
                self.json_output_frame.pack(fill=tk.X)
        except Exception as e:
            self.logger.error(f"显示输出参数面板时出错: {e}")
            messagebox.showerror("错误", f"显示输出参数面板时出错: {e}")
            
    def _save_config(self):
        """保存配置到文件"""
        try:
            config = self.get_report_config()
            if not config:
                messagebox.showwarning("警告", "没有配置数据可保存")
                return
                
            filename = filedialog.asksaveasfilename(
                title="保存配置",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
                defaultextension=".json"
            )
            
            if filename:
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                messagebox.showinfo("成功", "配置已保存")
                
        except Exception as e:
            self.logger.error(f"保存配置时出错: {e}")
            messagebox.showerror("错误", f"保存配置时出错: {e}")
            
    def _load_config(self):
        """从文件加载配置"""
        try:
            filename = filedialog.askopenfilename(
                title="加载配置",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            
            if filename:
                with open(filename, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                self.set_report_config(config)
                messagebox.showinfo("成功", "配置已加载")
                
        except json.JSONDecodeError as e:
            self.logger.error(f"加载配置时JSON解析错误: {e}")
            messagebox.showerror("错误", f"配置文件格式错误: {e}")
        except Exception as e:
            self.logger.error(f"加载配置时出错: {e}")
            messagebox.showerror("错误", f"加载配置时出错: {e}")
            
    def _reset_config(self):
        """重置配置"""
        try:
            result = messagebox.askyesno("确认", "确定要重置所有配置吗？")
            if result:
                self._create_widgets()  # 重新创建所有组件
                messagebox.showinfo("成功", "配置已重置")
        except Exception as e:
            self.logger.error(f"重置配置时出错: {e}")
            messagebox.showerror("错误", f"重置配置时出错: {e}")
            
    def get_report_config(self) -> Optional[Dict[str, Any]]:
        """获取报表配置"""
        try:
            config = {
                'basic_info': {
                    'name': self.report_name.get().strip(),
                    'description': self.report_desc.get("1.0", tk.END).strip()
                },
                'output_config': {
                    'format': self.output_format.get(),
                    'path': self.output_path.get().strip()
                },
                'style_config': {
                    'theme': self.style_theme.get(),
                    'custom_css': self.custom_css.get("1.0", tk.END).strip()
                },
                'data_config': {
                    'clean_data': self.clean_data_var.get(),
                    'validate_data': self.validate_data_var.get(),
                    'cache_data': self.cache_data_var.get()
                }
            }
            
            # 根据输出格式添加特定参数
            output_format = self.output_format.get()
            if output_format == "excel":
                try:
                    font_size = int(self.excel_font_size.get())
                    config['output_config'].update({
                        'sheet_name': self.excel_sheet_name.get().strip(),
                        'font_size': font_size
                    })
                except ValueError:
                    self.logger.warning("Excel字体大小无效，使用默认值")
                    config['output_config'].update({
                        'sheet_name': self.excel_sheet_name.get().strip(),
                        'font_size': 11
                    })
            elif output_format == "pdf":
                config['output_config'].update({
                    'page_size': self.pdf_page_size.get(),
                    'orientation': self.pdf_orientation.get(),
                    'include_charts': self.pdf_include_charts_var.get()
                })
            elif output_format == "html":
                config['output_config'].update({
                    'include_nav': self.html_include_nav_var.get(),
                    'include_search': self.html_include_search_var.get(),
                    'responsive': self.html_responsive_var.get()
                })
            elif output_format == "json":
                config['output_config'].update({
                    'indent': self.json_indent_var.get(),
                    'include_meta': self.json_include_meta_var.get(),
                    'compress': self.json_compress_var.get()
                })
            
            # 验证配置
            is_valid, error_msg = self._validate_config(config)
            if not is_valid:
                messagebox.showwarning("配置验证失败", error_msg)
                return None
                
            return config
            
        except Exception as e:
            self.logger.error(f"获取配置时出错: {e}")
            messagebox.showerror("错误", f"获取配置时出错: {e}")
            return None
        
    def set_report_config(self, config: Dict[str, Any]):
        """设置报表配置"""
        try:
            if not isinstance(config, dict):
                raise ValueError("配置必须是字典类型")
                
            basic_info = config.get('basic_info', {})
            self.report_name.delete(0, tk.END)
            self.report_name.insert(0, basic_info.get('name', ''))
            self.report_desc.delete("1.0", tk.END)
            self.report_desc.insert("1.0", basic_info.get('description', ''))
            
            output_config = config.get('output_config', {})
            output_format = output_config.get('format', 'excel')
            self.output_format.set(output_format)
            self.output_path.delete(0, tk.END)
            self.output_path.insert(0, output_config.get('path', ''))
            
            style_config = config.get('style_config', {})
            self.style_theme.set(style_config.get('theme', 'default'))
            self.custom_css.delete("1.0", tk.END)
            self.custom_css.insert("1.0", style_config.get('custom_css', ''))
            
            data_config = config.get('data_config', {})
            self.clean_data_var.set(data_config.get('clean_data', True))
            self.validate_data_var.set(data_config.get('validate_data', True))
            self.cache_data_var.set(data_config.get('cache_data', False))
            
            # 设置特定输出格式的参数
            if output_format == "excel":
                self.excel_sheet_name.delete(0, tk.END)
                self.excel_sheet_name.insert(0, output_config.get('sheet_name', '报表数据'))
                self.excel_font_size.delete(0, tk.END)
                font_size = output_config.get('font_size', 11)
                if isinstance(font_size, int) and 8 <= font_size <= 24:
                    self.excel_font_size.insert(0, str(font_size))
                else:
                    self.excel_font_size.insert(0, "11")
            elif output_format == "pdf":
                self.pdf_page_size.set(output_config.get('page_size', 'A4'))
                self.pdf_orientation.set(output_config.get('orientation', 'portrait'))
                self.pdf_include_charts_var.set(output_config.get('include_charts', True))
            elif output_format == "html":
                self.html_include_nav_var.set(output_config.get('include_nav', True))
                self.html_include_search_var.set(output_config.get('include_search', False))
                self.html_responsive_var.set(output_config.get('responsive', True))
            elif output_format == "json":
                self.json_indent_var.set(output_config.get('indent', True))
                self.json_include_meta_var.set(output_config.get('include_meta', True))
                self.json_compress_var.set(output_config.get('compress', False))
                
            # 更新显示的参数面板
            self._show_output_params_panel(output_format)
            
        except Exception as e:
            self.logger.error(f"设置配置时出错: {e}")
            messagebox.showerror("错误", f"设置配置时出错: {e}")
            # 设置默认配置
            self._create_widgets()