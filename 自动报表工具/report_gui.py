#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动报表工具GUI界面
重构版本 - 使用模块化组件
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sys
import os
import logging
from typing import Dict, Any

# 导入依赖管理模块
try:
    from dependencies import check_feature, is_feature_available
    has_pandas = is_feature_available('pandas')
    has_numpy = is_feature_available('numpy')
    has_openpyxl = is_feature_available('openpyxl')
    has_matplotlib = is_feature_available('matplotlib')
    has_reportlab = is_feature_available('reportlab')
    has_requests = is_feature_available('requests')
except ImportError:
    # 如果依赖管理模块不存在，定义备用功能
    def check_feature(feature):
        return True
    def is_feature_available(feature):
        return True

# 导入GUI组件
from gui_data_source import DataSourceManager
from gui_report_config import ReportConfigManager


class ReportGUI:
    """重构后的报表GUI主类"""
    
    def __init__(self, root=None):
        # 如果没有提供root，创建新的Tk实例
        if root is None:
            self.root = tk.Tk()
        else:
            self.root = root
            
        self.root.title("自动报表工具 v2.0")
        self.root.geometry("900x700")
        
        # 设置图标（如果存在）
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
            
        # 初始化日志记录
        self._setup_logging()
        
        # 检查依赖
        self._check_dependencies()
        
        # 创建GUI
        self._create_main_interface()
        
        # 存储当前配置
        self.current_config = {}
        
    def _setup_logging(self):
        """设置日志记录"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('report_generator.log', encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger('ReportGUI')
        
    def _check_dependencies(self):
        """检查依赖库"""
        required_features = ['pandas', 'numpy', 'openpyxl', 'requests']
        missing_features = []
        
        for feature in required_features:
            if not check_feature(feature):
                missing_features.append(feature)
                
        if missing_features:
            self.logger.warning(f"缺少以下依赖库: {missing_features}")
            messagebox.showwarning(
                "依赖库检查", 
                f"缺少以下依赖库:\n{', '.join(missing_features)}\n"
                f"某些功能可能无法正常工作。"
            )
            
    def _create_main_interface(self):
        """创建主界面"""
        # 创建菜单栏
        self._create_menu()
        
        # 创建工具栏
        self._create_toolbar()
        
        # 创建状态栏
        self._create_status_bar()
        
        # 创建主内容区域
        self._create_main_content()
        
    def _create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="新建配置", command=self._new_config, accelerator="Ctrl+N")
        file_menu.add_command(label="打开配置", command=self._open_config, accelerator="Ctrl+O")
        file_menu.add_command(label="保存配置", command=self._save_config, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="导入配置", command=self._import_config)
        file_menu.add_command(label="导出配置", command=self._export_config)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self._exit_app)
        
        # 运行菜单
        run_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="运行", menu=run_menu)
        run_menu.add_command(label="生成报表", command=self._generate_report, accelerator="F5")
        run_menu.add_command(label="预览报表", command=self._preview_report)
        run_menu.add_separator()
        run_menu.add_command(label="调度设置", command=self._schedule_settings)
        
        # 工具菜单
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="工具", menu=tools_menu)
        tools_menu.add_command(label="数据预览", command=self._data_preview)
        tools_menu.add_command(label="模板管理", command=self._template_manager)
        tools_menu.add_separator()
        tools_menu.add_command(label="日志查看", command=self._view_logs)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="使用帮助", command=self._show_help)
        help_menu.add_command(label="关于", command=self._show_about)
        
        # 绑定快捷键
        self.root.bind('<Control-n>', lambda e: self._new_config())
        self.root.bind('<Control-o>', lambda e: self._open_config())
        self.root.bind('<Control-s>', lambda e: self._save_config())
        self.root.bind('<F5>', lambda e: self._generate_report())
        
    def _create_toolbar(self):
        """创建工具栏"""
        toolbar = ttk.Frame(self.root)
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=2)
        
        # 新建配置按钮
        ttk.Button(toolbar, text="新建", command=self._new_config).pack(side=tk.LEFT, padx=2)
        
        # 打开配置按钮
        ttk.Button(toolbar, text="打开", command=self._open_config).pack(side=tk.LEFT, padx=2)
        
        # 保存配置按钮
        ttk.Button(toolbar, text="保存", command=self._save_config).pack(side=tk.LEFT, padx=2)
        
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 生成报表按钮
        self.generate_btn = ttk.Button(toolbar, text="生成报表", command=self._generate_report)
        self.generate_btn.pack(side=tk.LEFT, padx=2)
        
        # 预览按钮
        ttk.Button(toolbar, text="预览", command=self._preview_report).pack(side=tk.LEFT, padx=2)
        
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 进度条
        self.progress = ttk.Progressbar(toolbar, mode='determinate')
        self.progress.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)
        
    def _create_status_bar(self):
        """创建状态栏"""
        status_frame = ttk.Frame(self.root)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 状态标签
        self.status_label = ttk.Label(status_frame, text="就绪")
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        # 进度标签
        self.progress_label = ttk.Label(status_frame, text="")
        self.progress_label.pack(side=tk.RIGHT, padx=5)
        
    def _create_main_content(self):
        """创建主内容区域"""
        # 创建笔记本控件
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建数据源管理组件
        self.data_source_manager = DataSourceManager(self.notebook)
        
        # 创建报表配置管理组件
        self.report_config_manager = ReportConfigManager(self.notebook)
        
        # 添加日志标签页
        self._create_log_tab()
        
        # 绑定标签页切换事件
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        
    def _create_log_tab(self):
        """创建日志标签页"""
        log_frame = ttk.Frame(self.notebook)
        self.notebook.add(log_frame, text="运行日志")
        
        # 日志文本框
        self.log_text = tk.Text(log_frame, height=15, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 滚动条
        log_scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=log_scrollbar.set)
        
        # 清空日志按钮
        ttk.Button(log_frame, text="清空日志", command=self._clear_logs).pack(pady=5)
        
    def _on_tab_changed(self, event):
        """标签页切换事件处理"""
        current_tab = self.notebook.tab(self.notebook.select(), "text")
        self.logger.info(f"切换到标签页: {current_tab}")
        
    # 菜单和工具栏事件处理方法
    def _new_config(self):
        """新建配置"""
        self.current_config = {}
        self.data_source_manager.set_data_source_config({})
        self.report_config_manager.set_report_config({})
        self._update_status("已创建新配置")
        
    def _open_config(self):
        """打开配置"""
        file_path = filedialog.askopenfilename(
            title="打开配置文件",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                import json
                with open(file_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    
                self.current_config = config
                self.data_source_manager.set_data_source_config(config.get('data_source', {}))
                self.report_config_manager.set_report_config(config.get('report_config', {}))
                self._update_status(f"已打开配置文件: {os.path.basename(file_path)}")
                
            except Exception as e:
                messagebox.showerror("错误", f"打开配置文件失败:\n{str(e)}")
                
    def _save_config(self):
        """保存配置"""
        if not self.current_config:
            self.current_config = {}
            
        # 获取当前配置
        self.current_config.update({
            'data_source': self.data_source_manager.get_data_source_config(),
            'report_config': self.report_config_manager.get_report_config()
        })
        
        file_path = filedialog.asksaveasfilename(
            title="保存配置文件",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                import json
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(self.current_config, f, ensure_ascii=False, indent=2)
                    
                self._update_status(f"已保存配置文件: {os.path.basename(file_path)}")
                
            except Exception as e:
                messagebox.showerror("错误", f"保存配置文件失败:\n{str(e)}")
                
    def _import_config(self):
        """导入配置"""
        messagebox.showinfo("提示", "导入配置功能待实现")
        
    def _export_config(self):
        """导出配置"""
        messagebox.showinfo("提示", "导出配置功能待实现")
        
    def _exit_app(self):
        """退出应用"""
        self.root.quit()
        self.root.destroy()
        
    def _generate_report(self):
        """生成报表"""
        try:
            # 获取配置
            data_source_config = self.data_source_manager.get_data_source_config()
            report_config = self.report_config_manager.get_report_config()
            
            # 验证配置
            if not self._validate_config(data_source_config, report_config):
                return
                
            # 更新进度
            self._update_progress(10, "开始生成报表...")
            
            # 这里应该调用实际的报表生成逻辑
            self._update_progress(50, "处理数据中...")
            
            # 模拟处理时间
            import time
            time.sleep(2)
            
            self._update_progress(80, "生成输出文件...")
            time.sleep(1)
            
            self._update_progress(100, "报表生成完成!")
            
            # 显示成功消息
            messagebox.showinfo("成功", "报表生成完成!")
            
            # 重置进度
            self.root.after(2000, lambda: self._update_progress(0, ""))
            
        except Exception as e:
            self.logger.error(f"生成报表失败: {str(e)}")
            messagebox.showerror("错误", f"生成报表失败:\n{str(e)}")
            self._update_progress(0, "生成失败")
            
    def _preview_report(self):
        """预览报表"""
        messagebox.showinfo("提示", "预览功能待实现")
        
    def _schedule_settings(self):
        """调度设置"""
        messagebox.showinfo("提示", "调度设置功能待实现")
        
    def _data_preview(self):
        """数据预览"""
        messagebox.showinfo("提示", "数据预览功能待实现")
        
    def _template_manager(self):
        """模板管理"""
        messagebox.showinfo("提示", "模板管理功能待实现")
        
    def _view_logs(self):
        """查看日志"""
        self.notebook.select(2)  # 切换到日志标签页
        
    def _show_help(self):
        """显示帮助"""
        help_text = """
自动报表工具使用帮助

1. 数据源配置：
   - 选择数据源类型（Excel、CSV、SQL、API）
   - 配置数据源路径和参数
   
2. 报表配置：
   - 设置报表基本信息
   - 选择输出格式和参数
   - 配置样式和数据处理选项
   
3. 运行报表：
   - 点击"生成报表"按钮
   - 可以在运行日志中查看进度
   
4. 配置管理：
   - 使用文件菜单保存和加载配置
   - 支持导入/导出配置文件
        """
        messagebox.showinfo("使用帮助", help_text)
        
    def _show_about(self):
        """显示关于"""
        about_text = """
自动报表工具 v2.0

重构版本 - 模块化组件架构

功能特性：
- 支持多种数据源
- 灵活的报表配置
- 模块化GUI组件
- 完善的错误处理

作者：自动报表工具团队
版本：2.0.0
        """
        messagebox.showinfo("关于", about_text)
        
    def _clear_logs(self):
        """清空日志"""
        self.log_text.delete(1.0, tk.END)
        
    def _validate_config(self, data_source_config, report_config):
        """验证配置"""
        # 验证数据源配置
        if not data_source_config.get('path'):
            messagebox.showerror("配置错误", "请配置数据源路径")
            return False
            
        # 验证报表配置
        if not report_config.get('basic_info', {}).get('name'):
            messagebox.showerror("配置错误", "请设置报表名称")
            return False
            
        return True
        
    def _update_status(self, message):
        """更新状态栏"""
        self.status_label.config(text=message)
        self.root.update_idletasks()
        
    def _update_progress(self, value, message):
        """更新进度条和状态"""
        self.progress['value'] = value
        if message:
            self.progress_label.config(text=message)
        self.root.update_idletasks()
        
    def add_log_message(self, message, level="INFO"):
        """添加日志消息"""
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {level}: {message}\n"
        
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        
        # 限制日志条数
        lines = self.log_text.get("1.0", tk.END).split("\n")
        if len(lines) > 1000:
            self.log_text.delete("1.0", f"{len(lines) - 500}.0")
            
    def run(self):
        """运行GUI"""
        self.add_log_message("自动报表工具启动成功", "INFO")
        self.root.mainloop()


def main():
    """主函数"""
    try:
        root = tk.Tk()
        app = ReportGUI(root)
        app.run()
    except Exception as e:
        logging.error(f"GUI启动失败: {str(e)}")
        messagebox.showerror("启动错误", f"GUI启动失败:\n{str(e)}")


if __name__ == "__main__":
    main()