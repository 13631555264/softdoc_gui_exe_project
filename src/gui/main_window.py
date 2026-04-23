# gui/main_window.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
软著文档生成工具 - 主窗口
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import Optional, Dict, Any
import logging

# 导入自定义模块
from core.config import Config

# 设置日志
logger = logging.getLogger("softdoc_generator")

class MainWindow:
    """主窗口类"""
    
    def __init__(self):
        # 加载配置
        self.config = Config()
        
        # 创建主窗口
        self.window = tk.Tk()
        self.window.title("软著文档生成工具")
        self.window.configure(bg='#f0f0f0')
        
        # 设置窗口大小和位置
        self.window_width = self.config.get('gui.window_width', 800)
        self.window_height = self.config.get('gui.window_height', 600)
        self.setup_window_geometry()
        
        # 样式设置
        self.setup_styles()
        
        # 创建界面
        self.create_widgets()
        
        # 加载配置
        self.load_config()
        
        logger.info("GUI主窗口初始化完成")
    
    def setup_window_geometry(self):
        """设置窗口位置和大小"""
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()
        
        center_x = int(screen_width/2 - self.window_width/2)
        center_y = int(screen_height/2 - self.window_height/2)
        
        self.window.geometry(f'{self.window_width}x{self.window_height}+{center_x}+{center_y}')
        
        # 设置窗口图标（如果有）
        icon_path = Path(__file__).parent.parent / "resources" / "icons" / "app_icon.ico"
        if icon_path.exists():
            try:
                self.window.iconbitmap(str(icon_path))
            except:
                pass
    
    def setup_styles(self):
        """设置样式"""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # 自定义颜色
        self.colors = {
            'primary': '#007bff',
            'secondary': '#6c757d',
            'success': '#28a745',
            'danger': '#dc3545',
            'warning': '#ffc107',
            'info': '#17a2b8',
            'light': '#f8f9fa',
            'dark': '#343a40',
            'background': '#f0f0f0',
            'text': '#212529'
        }
        
        # 配置样式
        self.style.configure("Accent.TButton",
            background=self.colors['primary'],
            foreground="white",
            borderwidth=0,
            focuscolor="none",
            font=("Arial", 10, "bold")
        )
        
        self.style.map('Accent.TButton',
            background=[('active', self.colors['primary']),
                        ('disabled', self.colors['secondary'])],
            foreground=[('active', 'white'),
                        ('disabled','grey')]
        )
    
    def create_widgets(self):
        """创建界面组件"""
        # 创建主框架
        self.main_frame = ttk.Frame(self.window, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        
        # 标题标签
        title_label = ttk.Label(
            self.main_frame,
            text="软著文档生成工具",
            font=("Arial", 18, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=4, pady=(0,20))
        
        # 文件选择框架
        file_frame = ttk.LabelFrame(
            self.main_frame,
            text="文件选择",
            padding="10"
        )
        file_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(0,10))
        
        # 软著文件选择
        ttk.Label(file_frame, text="软著文件（PDF/图片）:").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        self.softdoc_entry = ttk.Entry(file_frame, width=50)
        self.softdoc_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(
            file_frame,
            text="浏览...",
            command=self.select_softdoc_file,
            width=10
        ).grid(row=0, column=2, padx=5)
        
        ttk.Button(
            file_frame,
            text="拖拽文件...",
            command=self.open_file_drop_dialog,
            width=10
        ).grid(row=0, column=3, padx=5)
        
        # 渠广文件选择
        ttk.Label(file_frame, text="渠广文件（txt）:").grid(row=1, column=0, sticky=tk.W, pady=5)
        
        self.qg_entry = ttk.Entry(file_frame, width=50)
        self.qg_entry.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Button(
            file_frame,
            text="浏览...",
            command=self.select_qg_file,
            width=10
        ).grid(row=1, column=2, padx=5)
        
        ttk.Button(
            file_frame,
            text="拖拽文件...",
            command=self.open_qg_drop_dialog,
            width=10
        ).grid(row=1, column=3, padx=5)
        
        # 模板目录选择
        ttk.Label(file_frame, text="模板目录:").grid(row=2, column=0, sticky=tk.W, pady=5)
        
        self.template_entry = ttk.Entry(file_frame, width=50)
        self.template_entry.grid(row=2, column=1, padx=5, pady=5)
        
        ttk.Button(
            file_frame,
            text="浏览...",
            command=self.select_template_dir,
            width=10
        ).grid(row=2, column=2, padx=5)
        
        ttk.Button(
            file_frame,
            text="拖拽文件夹...",
            command=self.open_template_drop_dialog,
            width=10
        ).grid(row=2, column=3, padx=5)
        
        # 输出目录选择
        ttk.Label(file_frame, text="输出目录:").grid(row=3, column=0, sticky=tk.W, pady=5)
        
        self.output_entry = ttk.Entry(file_frame, width=50)
        self.output_entry.grid(row=3, column=1, padx=5, pady=5)
        
        ttk.Button(
            file_frame,
            text="浏览...",
            command=self.select_output_dir,
            width=10
        ).grid(row=3, column=2, padx=5)
        
        # 高级设置框架
        setting_frame = ttk.LabelFrame(
            self.main_frame,
            text="高级设置",
            padding="10"
        )
        setting_frame.grid(row=2, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=(10,10))
        
        # OCR语言设置
        ttk.Label(setting_frame, text="OCR语言:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.ocr_combo = ttk.Combobox(
            setting_frame,
            values=['chi_sim+eng', 'chi_sim', 'eng'],
            width=20,
            state='readonly'
        )
        self.ocr_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        self.ocr_combo.set('chi_sim+eng')
        
        # 授权年限设置
        ttk.Label(setting_frame, text="授权年限:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.years_spin = ttk.Spinbox(
            setting_frame,
            from_=1,
            to=30,
            width=10
        )
        self.years_spin.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        self.years_spin.set('10')
        
        # 按钮框架
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=3, column=0, columnspan=4, pady=(10,0))
        
        # 开始按钮
        self.start_button = ttk.Button(
            button_frame,
            text="开始生成文档",
            command=self.start_processing,
            style="Accent.TButton",
            width=20
        )
        self.start_button.grid(row=0, column=0, padx=5, pady=10)
        
        # 配置按钮
        ttk.Button(
            button_frame,
            text="加载配置",
            command=self.load_config,
            width=15
        ).grid(row=0, column=1, padx=5)
        
        ttk.Button(
            button_frame,
            text="保存配置",
            command=self.save_config,
            width=15
        ).grid(row=0, column=2, padx=5)
    
    def load_config(self):
        """加载配置"""
        try:
            # 更新界面配置
            self.softdoc_entry.delete(0, tk.END)
            self.softdoc_entry.insert(0, self.config.get_last_path('softdoc'))
            
            self.qg_entry.delete(0, tk.END)
            self.qg_entry.insert(0, self.config.get_last_path('qg'))
            
            self.template_entry.delete(0, tk.END)
            self.template_entry.insert(0, self.config.get_last_path('template'))
            
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, self.config.get_last_path('output'))
            
            # 高级设置
            self.ocr_combo.set(self.config.get_processing_config('ocr_language', 'chi_sim+eng'))
            self.years_spin.delete(0, tk.END)
            self.years_spin.insert(0, str(self.config.get_processing_config('authorization_years', 10)))
            
            logger.info("配置加载到界面")
        except Exception as e:
            logger.error(f"加载配置到界面失败: {e}")
    
    def save_config(self):
        """保存配置"""
        try:
            # 从界面获取配置
            softdoc_path = self.softdoc_entry.get()
            if softdoc_path:
                self.config.set_last_path('softdoc', softdoc_path)
            
            qg_path = self.qg_entry.get()
            if qg_path:
                self.config.set_last_path('qg', qg_path)
            
            template_dir = self.template_entry.get()
            if template_dir:
                self.config.set_last_path('template', template_dir)
            
            output_dir = self.output_entry.get()
            if output_dir:
                self.config.set_last_path('output', output_dir)
            
            # 高级设置
            self.config.set_processing_config('ocr_language', self.ocr_combo.get())
            
            try:
                auth_years = int(self.years_spin.get())
                self.config.set_processing_config('authorization_years', auth_years)
            except ValueError:
                pass
            
            # 保存到文件
            if self.config.save_config():
                messagebox.showinfo("保存成功", "配置已成功保存")
                logger.info("配置保存成功")
            else:
                messagebox.showerror("保存失败", "配置保存失败")
                logger.error("配置保存失败")
        except Exception as e:
            logger.error(f"保存配置失败: {e}")
            messagebox.showerror("保存失败", f"保存配置时发生错误: {e}")
    
    def select_softdoc_file(self):
        """选择软著文件"""
        filetypes = [
            ("PDF files", "*.pdf"),
            ("Image files", "*.jpg *.jpeg *.png *.bmp *.gif *.tiff"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title="选择软著文件",
            filetypes=filetypes,
            initialdir=os.path.dirname(self.softdoc_entry.get()) if self.softdoc_entry.get() else '.'
        )
        
        if filename:
            self.softdoc_entry.delete(0, tk.END)
            self.softdoc_entry.insert(0, filename)
            
            # 自动保存配置
            self.config.set_last_path('softdoc', filename)
            
            logger.info(f"选择软著文件: {filename}")
    
    def select_qg_file(self):
        """选择渠广文件"""
        filetypes = [("Text files", "*.txt"), ("All files", "*.*")]
        
        filename = filedialog.askopenfilename(
            title="选择渠广文件",
            filetypes=filetypes,
            initialdir=os.path.dirname(self.qg_entry.get()) if self.qg_entry.get() else '.'
        )
        
        if filename:
            self.qg_entry.delete(0, tk.END)
            self.qg_entry.insert(0, filename)
            
            # 自动保存配置
            self.config.set_last_path('qg', filename)
            
            logger.info(f"选择渠广文件: {filename}")
    
    def select_template_dir(self):
        """选择模板目录"""
        directory = filedialog.askdirectory(
            title="选择模板目录",
            initialdir=self.template_entry.get() if self.template_entry.get() else '.'
        )
        
        if directory:
            self.template_entry.delete(0, tk.END)
            self.template_entry.insert(0, directory)
            
            # 自动保存配置
            self.config.set_last_path('template', directory)
            
            logger.info(f"选择模板目录: {directory}")
    
    def select_output_dir(self):
        """选择输出目录"""
        directory = filedialog.askdirectory(
            title="选择输出目录",
            initialdir=self.output_entry.get() if self.output_entry.get() else '.'
        )
        
        if directory:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, directory)
            
            # 自动保存配置
            self.config.set_last_path('output', directory)
            
            logger.info(f"选择输出目录: {directory}")
    
    def open_file_drop_dialog(self):
        """打开文件拖拽对话框"""
        # 这里可以调用自定义的文件拖拽组件
        logger.info("打开文件拖拽对话框")
        # 暂时用简单的文件选择代替
        self.select_softdoc_file()
    
    def open_qg_drop_dialog(self):
        """打开渠广文件拖拽对话框"""
        logger.info("打开渠广文件拖拽对话框")
        self.select_qg_file()
    
    def open_template_drop_dialog(self):
        """打开模板拖拽对话框"""
        logger.info("打开模板拖拽对话框")
        self.select_template_dir()
    
    def start_processing(self):
        """开始处理（在后台线程中执行）"""
        # 获取输入参数
        softdoc_path = self.softdoc_entry.get()
        qg_path = self.qg_entry.get()
        template_dir = self.template_entry.get()
        output_dir = self.output_entry.get()
        ocr_language = self.ocr_combo.get()
        
        try:
            auth_years = int(self.years_spin.get())
        except ValueError:
            auth_years = 10
        
        # 验证输入
        errors = []
        
        if not softdoc_path or not os.path.exists(softdoc_path):
            errors.append("软著文件不存在或未选择")
        
        if not qg_path or not os.path.exists(qg_path):
            errors.append("渠广文件不存在或未选择")
        
        if not template_dir or not os.path.exists(template_dir):
            errors.append("模板目录不存在或未选择")
        
        if not output_dir:
            errors.append("输出目录未选择")
        else:
            os.makedirs(output_dir, exist_ok=True)
        
        if errors:
            messagebox.showerror("输入错误", "\n".join(errors))
            return
        
        # 禁用开始按钮，防止重复点击
        self.start_button.config(state='disabled', text="处理中...")
        
        # 创建进度条窗口
        self.progress_window = tk.Toplevel(self.window)
        self.progress_window.title("处理中")
        self.progress_window.geometry("400x150")
        self.progress_window.transient(self.window)
        self.progress_window.grab_set()
        
        # 居中显示
        x = self.window.winfo_x() + (self.window.winfo_width() - 400) // 2
        y = self.window.winfo_y() + (self.window.winfo_height() - 150) // 2
        self.progress_window.geometry(f"400x150+{x}+{y}")
        
        # 进度标签
        self.progress_label = ttk.Label(self.progress_window, text="正在初始化...", font=("Arial", 10))
        self.progress_label.pack(pady=20)
        
        # 进度条
        self.progress_bar = ttk.Progressbar(self.progress_window, mode='indeterminate', length=350)
        self.progress_bar.pack(pady=10)
        self.progress_bar.start(10)
        
        # 在后台线程中执行处理
        self.processing_thread = threading.Thread(
            target=self.process_documents,
            args=(softdoc_path, qg_path, template_dir, output_dir, ocr_language, auth_years),
            daemon=True
        )
        self.processing_thread.start()
    
    def update_progress(self, message):
        """更新进度信息（在主线程中执行）"""
        def _update():
            if hasattr(self, 'progress_label') and self.progress_label:
                self.progress_label.config(text=message)
        self.window.after(0, _update)
    
    def process_documents(self, softdoc_path: str, qg_path: str, template_dir: str, 
                         output_dir: str, ocr_language: str, auth_years: int):
        """处理文档生成（在后台线程中运行）"""
        try:
            self.update_progress("正在解析渠广文件...")
            
            from core.qg_parser import QGParser
            from core.softdoc_parser import SoftDocParser
            from core.document_generator import DocumentGenerator
            
            # 更新配置
            self.config.set_processing_config('ocr_language', ocr_language)
            self.config.set_processing_config('authorization_years', auth_years)
            
            # 1. 解析渠广文件
            qg_parser = QGParser(self.config)
            game_info = qg_parser.parse_file(qg_path)
            logger.info(f"渠广解析: 游戏={game_info.get('game_name')}, 公司={game_info.get('publisher')}")
            
            # 2. 解析软著文件（这里会调用 API，需要几秒钟）
            self.update_progress("正在识别软著文件（调用API，请稍候）...")
            
            soft_parser = SoftDocParser(self.config)
            
            try:
                soft_info = soft_parser.parse_file(softdoc_path)
                logger.info(f"软著解析: 软件={soft_info.get('software_name')}, 著作权人={soft_info.get('copyright_holder')}")
            except Exception as e:
                logger.error(f"软著解析失败: {e}")
                soft_info = {
                    'software_name': game_info.get('game_name', ''),
                    'version': game_info.get('version', ''),
                    'copyright_holder': game_info.get('publisher', ''),
                    'software_type': '',
                    'registration_number': '',
                    'completion_date': '',
                    'publish_date': '',
                    'original_text': ''
                }
            
            # 3. 补充缺失的信息
            if not soft_info.get('software_name'):
                soft_info['software_name'] = game_info.get('game_name', '')
            if not soft_info.get('copyright_holder'):
                soft_info['copyright_holder'] = game_info.get('publisher', '')
            
            # 4. 生成文档
            self.update_progress("正在生成文档...")
            
            generator = DocumentGenerator(self.config)
            generator.set_template_dir(template_dir)
            generator.set_output_dir(output_dir)
            
            generated_files = generator.generate_documents(game_info, soft_info)
            
            # 5. 处理完成，关闭进度窗口
            self.window.after(0, self.on_processing_finished, True, game_info, soft_info, output_dir, generated_files)
            
        except Exception as e:
            logger.error(f"处理失败: {e}")
            import traceback
            traceback.print_exc()
            self.window.after(0, self.on_processing_finished, False, None, None, None, None, str(e))
    
    def on_processing_finished(self, success, game_info=None, soft_info=None, output_dir=None, generated_files=None, error_msg=None):
        """处理完成回调（在主线程中执行）"""
        # 关闭进度窗口
        if hasattr(self, 'progress_window') and self.progress_window:
            self.progress_window.destroy()
            self.progress_window = None
        
        # 恢复开始按钮
        self.start_button.config(state='normal', text="开始生成文档")
        
        if success:
            # 显示成功信息
            result_msg = f"文档生成完成！\n\n"
            result_msg += f"游戏名称: {game_info.get('game_name', '未知')}\n"
            result_msg += f"著作权人: {soft_info.get('copyright_holder', '未知')}\n"
            result_msg += f"生成文件数: {len(generated_files)}\n"
            result_msg += f"输出目录: {output_dir}"
            
            messagebox.showinfo("处理完成", result_msg)
            logger.info(f"处理完成，共生成{len(generated_files)}个文件")
            
            # 打开输出目录
            if generated_files and output_dir and os.path.exists(output_dir):
                os.startfile(output_dir)
        else:
            messagebox.showerror("处理失败", f"文档生成失败: {error_msg}")
    
    
    def on_processing_start(self):
        """处理开始回调"""
        messagebox.showinfo("处理开始", "已开始处理文档生成，请查看日志了解详细进度")
        logger.info("处理开始回调")
    
    def mainloop(self):
        """启动主循环"""
        self.window.mainloop()

def main():
    """主函数"""
    # 配置日志
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    window = MainWindow()
    window.mainloop()

if __name__ == "__main__":
    main()