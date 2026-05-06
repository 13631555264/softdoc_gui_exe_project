#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
软著文档生成工具 - 主窗口（整合版）
"""

import os
import sys

# 强制设置编码
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='ignore')
        sys.stderr.reconfigure(encoding='utf-8', errors='ignore')
    except:
        pass
    os.environ['PYTHONIOENCODING'] = 'utf-8'

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import Optional, Dict, Any, List
import logging
import threading
import re

# 导入拖拽支持
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

# 导入自定义模块
from core.config import Config

logger = logging.getLogger("softdoc_generator")


# ------------------------------------------------------------------ #
# 批量文件匹配器
# ------------------------------------------------------------------ #
class BatchFileMatcher:
    """根据游戏名称，将渠广目录下的 txt 与软著目录下的 pdf 自动配对"""

    def __init__(self, qg_dir: str, softdoc_dir: str):
        self.qg_dir = qg_dir
        self.softdoc_dir = softdoc_dir

    def _list_qg_files(self) -> List[str]:
        return [
            os.path.join(self.qg_dir, f)
            for f in os.listdir(self.qg_dir)
            if f.lower().endswith('.txt')
        ]

    def _list_softdoc_files(self) -> List[str]:
        img_exts = ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp')
        return [
            os.path.join(self.softdoc_dir, f)
            for f in os.listdir(self.softdoc_dir)
            if f.lower().endswith('.pdf') or os.path.splitext(f.lower())[1] in img_exts
        ]

    @staticmethod
    def extract_game_name_from_qg(filename: str) -> str:
        stem = os.path.splitext(os.path.basename(filename))[0]
        if '_' in stem:
            return stem.split('_')[0].strip()
        return stem.strip()

    def match(self, api_ocr=None) -> Dict:
        import glob as _glob, re as _re
        
        qg_files = self._list_qg_files()
        softdoc_files = self._list_softdoc_files()
        
        logger.info(f"【匹配】渠广文件 {len(qg_files)} 个，软著文件 {len(softdoc_files)} 个")
        
        softdoc_index: Dict[str, str] = {}
        for sp in softdoc_files:
            stem = os.path.splitext(os.path.basename(sp))[0]
            softdoc_index[stem] = sp
        
        matched = []
        qg_only = []
        used_softdoc = set()
        
        for qp in qg_files:
            game_name = self.extract_game_name_from_qg(qp)
            
            found_files = []
            for stem, sp in softdoc_index.items():
                if stem.startswith(game_name):
                    found_files.append(sp)
            
            if found_files:
                matched.append((qp, os.path.dirname(found_files[0]), game_name))
                used_softdoc.update(found_files)
            else:
                qg_only.append(qp)
        
        all_ocr_texts: Dict[str, str] = {}
        
        # 初始化 game_to_qg 字典（放在这里，确保在任何分支都能访问）
        game_to_qg: Dict[str, List[str]] = {}
        
        if api_ocr:
            all_softdoc_dirs: Dict[str, List[str]] = {}
            for sp in softdoc_files:
                d = os.path.dirname(sp)
                if d not in all_softdoc_dirs:
                    all_softdoc_dirs[d] = []
                all_softdoc_dirs[d].append(sp)
            
            all_game_names = set()
            for qp in qg_files:
                gn = self.extract_game_name_from_qg(qp)
                if gn:
                    all_game_names.add(gn)
            all_game_hints = '|'.join(all_game_names)
            
            for softdoc_dir, images in all_softdoc_dirs.items():
                folder_game_name, image_to_game, folder_raw_texts = self._build_softdoc_folder_game_names(
                    softdoc_dir, api_ocr, game_hint=all_game_hints
                )
                all_ocr_texts.update(folder_raw_texts)
            
            # OCR 回退匹配
            if qg_only:
                game_to_qg = {}  # 重新初始化
                for qp in qg_only:
                    gn = self.extract_game_name_from_qg(qp)
                    if gn not in game_to_qg:
                        game_to_qg[gn] = []
                    game_to_qg[gn].append(qp)
                
                for softdoc_dir, images in all_softdoc_dirs.items():
                    folder_texts = [text for img, text in all_ocr_texts.items() 
                                if os.path.dirname(img) == softdoc_dir or 
                                    os.path.dirname(img) == os.path.normpath(softdoc_dir)]
                    combined_text = '\n'.join(folder_texts) if folder_texts else ''
                    
                    # 遍历 game_to_qg 的副本，因为要修改原字典
                    for gn, qp_list in list(game_to_qg.items()):
                        if not qp_list:
                            continue
                        
                        if gn in combined_text:
                            for qp in qp_list:
                                matched.append((qp, softdoc_dir, gn))
                            game_to_qg[gn] = []
                
                # 更新 qg_only 为仍未匹配的
                qg_only = [qp for qps in game_to_qg.values() for qp in qps]
        
        softdoc_only = [sp for sp in softdoc_files if sp not in used_softdoc]
        
        return {
            'matched': matched,
            'qg_only': qg_only,
            'softdoc_only': softdoc_only,
            'all_ocr_texts': all_ocr_texts,
        }
    
    def _build_softdoc_folder_game_names(self, softdoc_dir: str, api_ocr, game_hint: str = ''):
        import glob as _glob
        softdoc_dir = os.path.normpath(softdoc_dir)
        img_exts = ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp')
        images = []
        for ext in img_exts:
            images.extend(_glob.glob(os.path.join(softdoc_dir, f'*{ext}')))
            images.extend(_glob.glob(os.path.join(softdoc_dir, f'*{ext.upper()}')))
        images = sorted(set(images))
        # 关键：统一使用正斜杠作为路径分隔符
        images = [p.replace('\\', '/') for p in images]
        
        folder_game_name = ''
        image_to_game: Dict[str, str] = {}
        all_raw_texts: Dict[str, str] = {}
        
        for img in images:
            gn, raw_text = self._ocr_image_extract_game_name(img, api_ocr, fallback_hint=game_hint)
            image_to_game[img] = gn
            if raw_text:
                all_raw_texts[img] = raw_text  # 用标准化路径作为键
            if gn and not folder_game_name:
                folder_game_name = gn
        
        return (folder_game_name, image_to_game, all_raw_texts)
        
    def _ocr_image_extract_game_name(self, image_path: str, api_ocr, fallback_hint: str = ''):
        import time
        start_time = time.time()
        
        print(f"  【OCR调用】开始识别: {os.path.basename(image_path)}")
        
        try:
            raw_text = api_ocr.recognize_image(image_path)
            elapsed = time.time() - start_time
            print(f"  【OCR调用】完成，耗时 {elapsed:.1f}秒，文本长度 {len(raw_text) if raw_text else 0}")
            
            if not raw_text or not raw_text.strip():
                print(f"  【OCR调用】返回空文本")
                return ('', '')
            
            game_name = ''
            
            # 策略0：优先匹配 fallback_hint
            if fallback_hint:
                hints = [h.strip() for h in fallback_hint.split('|') if h.strip()]
                for hint in hints:
                    if hint in raw_text:
                        print(f"  【OCR调用】匹配到 hint: {hint}")
                        return (hint, raw_text)
            
            # 策略1：从《游戏名》格式提取
            m = re.search(r'《([^》]{2,30})》', raw_text)
            if m:
                name = m.group(1).strip()
                if name not in ('游戏', '软件', '系统', '平台', '著作权授权书'):
                    print(f"  【OCR调用】从书名号提取: {name}")
                    return (name, raw_text)
            
            # 策略2：从"软件名称"字段提取
            m = re.search(r'软件名称[：:\s]*([^\n]{2,30})', raw_text)
            if m:
                print(f"  【OCR调用】从软件名称提取: {m.group(1).strip()}")
                return (m.group(1).strip(), raw_text)
            
            # 策略3：取第一行
            lines = [l.strip() for l in raw_text.split('\n') if l.strip()]
            if lines and 2 < len(lines[0]) < 30:
                print(f"  【OCR调用】从第一行提取: {lines[0]}")
                return (lines[0], raw_text)
            
            print(f"  【OCR调用】未能提取游戏名")
            return ('', raw_text)
        except Exception as e:
            print(f"  【OCR调用】异常: {e}")
            return ('', '')


# ------------------------------------------------------------------ #
# 主窗口
# ------------------------------------------------------------------ #
class MainWindow:
    """主窗口类（整合版）"""

    def __init__(self):
        self.config = Config()

        if HAS_DND:
            self.window = TkinterDnD.Tk()
        else:
            self.window = tk.Tk()

        self.window.title("软著文档生成工具")
        self.window.configure(bg='#f0f0f0')

        self.window_width = self.config.get('gui.window_width', 1000)
        self.window_height = self.config.get('gui.window_height', 1000)
        self.setup_window_geometry()

        self.processing_thread = None
        self.progress_window = None
        self.progress_label = None
        self.progress_bar = None

        self.setup_styles()
        self.create_widgets()
   

  # 在 load_config 之前，先设置模板目录的默认配置（如果为空）
        self._set_default_template_dir()
        
        self.load_config()
        
        logger.info("GUI 主窗口初始化完成（整合版）")


    def _set_default_template_dir(self):
        """设置默认模板目录"""
        script_dir = Path(__file__).parent.parent.parent  # 项目根目录
        default_template_dir = script_dir / "VIVO小游戏资质模版"
        
        # 检查当前配置中是否有模板目录
        current_template = self.config.get_last_path('template_dir')
        if not current_template and default_template_dir.exists():
            self.config.set_last_path('template_dir', str(default_template_dir))
            logger.info(f"默认模板目录已设置: {default_template_dir}")
        elif current_template:
            logger.info(f"当前模板目录: {current_template}")
        else:
            logger.warning(f"默认模板目录不存在: {default_template_dir}")

    # ---------------------------------------------------------------- #
    # 窗口基础
    # ---------------------------------------------------------------- #
    def setup_window_geometry(self):
        sw = self.window.winfo_screenwidth()
        sh = self.window.winfo_screenheight()
        cx = int(sw / 2 - self.window_width / 2)
        cy = int(sh / 2 - self.window_height / 2)
        self.window.geometry(f'{self.window_width}x{self.window_height}+{cx}+{cy}')

        icon_path = Path(__file__).parent.parent / "resources" / "icons" / "app_icon.ico"
        if icon_path.exists():
            try:
                self.window.iconbitmap(str(icon_path))
            except:
                pass

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.colors = {
            'primary': '#007bff', 'secondary': '#6c757d',
            'success': '#28a745', 'danger': '#dc3545',
            'warning': '#ffc107', 'info': '#17a2b8',
            'light': '#f8f9fa', 'dark': '#343a40',
            'background': '#f0f0f0', 'text': '#212529',
        }
        self.style.configure(
            "Accent.TButton",
            background=self.colors['primary'], foreground="white",
            borderwidth=0, focuscolor="none", font=("Arial", 10, "bold")
        )
        self.style.map(
            'Accent.TButton',
            background=[('active', self.colors['primary']), ('disabled', self.colors['secondary'])],
            foreground=[('active', 'white'), ('disabled', 'grey')]
        )

    # ---------------------------------------------------------------- #
    # 拖拽
    # ---------------------------------------------------------------- #
    def setup_drag_drop(self, entry_widget, path_type):
        if not HAS_DND:
            return

        def on_drop(event):
            files = event.data
            if not files:
                return event.action
            file_path = files.strip('{}')
            if ' ' in file_path and not os.path.exists(file_path):
                for part in file_path.split():
                    clean = part.strip('{}')
                    if os.path.exists(clean):
                        file_path = clean
                        break
            if os.path.exists(file_path):
                entry_widget.delete(0, tk.END)
                entry_widget.insert(0, file_path)
                self.config.set_last_path(path_type, file_path)
            return event.action

        def on_drag_enter(event):
            entry_widget.config(bg="#e8f4f8")
            return event.action

        def on_drag_leave(event):
            entry_widget.config(bg="white")
            return event.action

        entry_widget.drop_target_register(DND_FILES)
        entry_widget.dnd_bind('<<Drop>>', on_drop)
        entry_widget.dnd_bind('<<DragEnter>>', on_drag_enter)
        entry_widget.dnd_bind('<<DragLeave>>', on_drag_leave)

    # ---------------------------------------------------------------- #
    # 构建界面
    # ---------------------------------------------------------------- #
    def create_widgets(self):
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        
        # 标题
        ttk.Label(
            main_frame, text="软著文档生成工具",
            font=("Arial", 18, "bold")
        ).grid(row=0, column=0, pady=(0, 10))

        # 创建主内容区域
        self._build_main_content(main_frame)

    def _build_main_content(self, parent):
        # 创建可滚动的主框架
        canvas = tk.Canvas(parent, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        parent.rowconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)
        
        # === 路径配置区域 ===
        path_frame = ttk.LabelFrame(scrollable_frame, text="路径配置", padding="10")
        path_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        path_frame.columnconfigure(1, weight=1)
        
        path_configs = [
            ("软著文件夹:", 'softdoc_dir', '选择软著文件夹'),
            ("user.xlsx:", 'user_xlsx', '选择 user.xlsx（账号信息）'),
            ("game.xlsx 保存路径:", 'game_xlsx', '选择 game.xlsx 保存目录'),
            ("渠广输出文件夹:", 'qg_output_dir', '选择渠广输出文件夹'),
            ("模板目录:", 'template_dir', '选择文档模板目录'),
            ("最终文档输出目录:", 'output_dir', '选择最终文档输出目录'),
        ]
        
        self.entries = {}
        for i, (label, key, tip) in enumerate(path_configs):
            ttk.Label(path_frame, text=label, width=18, anchor='e').grid(
                row=i, column=0, sticky=tk.E, pady=4, padx=(0, 5))
            ent = ttk.Entry(path_frame, width=60)
            ent.grid(row=i, column=1, padx=5, pady=4, sticky=(tk.W, tk.E))
            self.setup_drag_drop(ent, key)
            self.entries[key] = ent
            
            def _make_browse(k=key, t=tip):
                def _browse():
                    if k == 'user_xlsx':
                        p = tk.filedialog.askopenfilename(
                            title=t, 
                            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
                        )
                    elif k == 'game_xlsx':
                        p = tk.filedialog.askdirectory(title=t)
                        if p:
                            p = os.path.join(p, "game.xlsx")
                    else:
                        p = tk.filedialog.askdirectory(title=t)
                    
                    if p:
                        self.entries[k].delete(0, tk.END)
                        self.entries[k].insert(0, p)
                        self.config.set_last_path(k, p)
                return _browse
            ttk.Button(path_frame, text="浏览...", command=_make_browse(), width=8).grid(
                row=i, column=2, padx=5)
        
        # === 一键生成按钮（放在广告类型之前，显眼位置）===
        onekey_frame = ttk.Frame(scrollable_frame)
        onekey_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 10))
        
        self.onekey_btn = ttk.Button(
            onekey_frame, 
            text="🚀 一键生成（完整流程：扫描软著 → 生成渠广 → 生成文档）",
            command=self.onekey_generate,
            style="Accent.TButton", 
            width=80
        )
        self.onekey_btn.pack(pady=10)
        
        # === 广告类型配置区域 ===
        ads_frame = ttk.LabelFrame(scrollable_frame, text="广告位类型（生成渠广时使用）", padding="10")
        ads_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.ads_vars = {}
        self.ads_num = {}
        self.ads_text = {}
        
        ads_defs = [
            ('open',      True, '1', 'H5-开屏',     0, 0),
            ('banner',    True, '1', 'H5-banner',   0, 3),
            ('reward',    True, '1', 'H5-激励视频',  1, 0),
            ('ori',       True, '1', 'H5-结算模版',  1, 3),
            ('ori_other', True, '1', 'H5-其它模版',  2, 3),
        ]
        ad_label_map = {
            'open': '开屏', 'banner': 'banner',
            'reward': '激励', 'ori': '原生(结算)', 'ori_other': '原生(其它)'
        }
        
        for ad_key, default_on, default_num, default_text, row, col in ads_defs:
            var = tk.BooleanVar(value=default_on)
            self.ads_vars[ad_key] = var
            cb = ttk.Checkbutton(ads_frame, text=ad_label_map[ad_key], variable=var)
            cb.grid(row=row, column=col, sticky=tk.W, padx=(5, 0), pady=2)
            
            num_ent = ttk.Entry(ads_frame, width=5)
            num_ent.insert(0, default_num)
            num_ent.grid(row=row, column=col+1, sticky=tk.W, padx=2)
            self.ads_num[ad_key] = num_ent
            
            txt_ent = ttk.Entry(ads_frame, width=12)
            txt_ent.insert(0, default_text)
            txt_ent.grid(row=row, column=col+2, sticky=tk.W, padx=2)
            self.ads_text[ad_key] = txt_ent
        
        # === 高级设置区域 ===
        setting_frame = ttk.LabelFrame(scrollable_frame, text="高级设置", padding="10")
        setting_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(setting_frame, text="OCR 语言:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.ocr_combo = ttk.Combobox(setting_frame, values=['chi_sim+eng', 'chi_sim', 'eng'], width=20, state='readonly')
        self.ocr_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        self.ocr_combo.set('chi_sim+eng')
        
        ttk.Label(setting_frame, text="授权年限:").grid(row=0, column=2, sticky=tk.W, pady=5, padx=(20, 0))
        self.years_spin = ttk.Spinbox(setting_frame, from_=1, to=30, width=10)
        self.years_spin.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        self.years_spin.set('10')
        
        # === 匹配预览区域 ===
        preview_frame = ttk.LabelFrame(scrollable_frame, text="匹配预览", padding="8")
        preview_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        self.preview_text = tk.Text(
            preview_frame, height=8, width=80, state='disabled',
            font=("Consolas", 9), bg="#fafafa", relief=tk.FLAT
        )
        scrollbar_preview = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
        self.preview_text.configure(yscrollcommand=scrollbar_preview.set)
        self.preview_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_preview.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # === 功能按钮区域（小按钮）===
        btn_frame = ttk.Frame(scrollable_frame)
        btn_frame.grid(row=5, column=0, pady=(5, 10))
        
        self.scan_btn = ttk.Button(
            btn_frame, text="① 扫描软著 → 生成 game.xlsx",
            command=self.scan_softdoc, width=28
        )
        self.scan_btn.grid(row=0, column=0, padx=5)
        
        self.gen_qg_btn = ttk.Button(
            btn_frame, text="② 生成渠广 txt 文件",
            command=self.generate_qg, width=22
        )
        self.gen_qg_btn.grid(row=0, column=1, padx=5)
        
        # self.preview_btn = ttk.Button(
        #     btn_frame, text="预览匹配", command=self.preview_match, width=14
        # )
        # self.preview_btn.grid(row=0, column=2, padx=5)
        
        self.batch_btn = ttk.Button(
            btn_frame, text="批量生成文档",
            command=self.batch_generate_documents,
            width=18
        )
        self.batch_btn.grid(row=0, column=3, padx=5)
        
        # === 日志区域 ===
        log_frame = ttk.LabelFrame(scrollable_frame, text="日志", padding="5")
        log_frame.grid(row=6, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(
            log_frame, height=12, state='disabled',
            font=("Consolas", 9), bg="#fafafa", relief=tk.FLAT
        )
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
    # ---------------------------------------------------------------- #
    # 日志方法
    # ---------------------------------------------------------------- #
    def log_append(self, msg: str):
        def _do():
            self.log_text.config(state='normal')
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state='disabled')
        self.window.after(0, _do)

    def set_preview(self, text: str):
        def _do():
            self.preview_text.config(state='normal')
            self.preview_text.delete('1.0', tk.END)
            self.preview_text.insert(tk.END, text)
            self.preview_text.config(state='disabled')
        self.window.after(0, _do)

    # ---------------------------------------------------------------- #
    # 配置存取
    # ---------------------------------------------------------------- #
    def load_config(self):
        try:
            for key, entry in self.entries.items():
                val = self.config.get_last_path(key)
                if val:
                    entry.delete(0, tk.END)
                    entry.insert(0, val)
            
            self.ocr_combo.set(self.config.get_processing_config('ocr_language', 'chi_sim+eng'))
            self.years_spin.delete(0, tk.END)
            self.years_spin.insert(0, str(self.config.get_processing_config('authorization_years', 10)))
        except Exception as e:
            logger.error(f"加载配置失败: {e}")

    def get_ads_config(self) -> dict:
        ads = {}
        for ad_key, var in self.ads_vars.items():
            if var.get():
                try:
                    num = int(self.ads_num[ad_key].get())
                except ValueError:
                    num = 1
                ads[ad_key] = {
                    'number': num,
                    'text': self.ads_text[ad_key].get().strip()
                }
        return ads

    # ---------------------------------------------------------------- #
    # 步骤1：扫描软著 → 生成 game.xlsx
    # ---------------------------------------------------------------- #
    def scan_softdoc(self):
        softdoc_dir = self.entries['softdoc_dir'].get().strip()
        game_xlsx_input = self.entries['game_xlsx'].get().strip()

        if not softdoc_dir or not os.path.isdir(softdoc_dir):
            messagebox.showerror("错误", "请先选择软著文件夹")
            return
        if not game_xlsx_input:
            messagebox.showerror("错误", "请先选择 game.xlsx 保存路径")
            return

        game_xlsx = os.path.normpath(game_xlsx_input)
        if os.path.isdir(game_xlsx):
            game_xlsx = os.path.join(game_xlsx, "game.xlsx")
        elif not game_xlsx.lower().endswith('.xlsx'):
            game_xlsx += ".xlsx"
        
        target_dir = os.path.dirname(game_xlsx)
        if target_dir:
            os.makedirs(target_dir, exist_ok=True)

        self.log_append("=" * 60)
        self.log_append(f"开始扫描软著文件夹: {softdoc_dir}")
        self.log_append(f"game.xlsx 将保存至: {game_xlsx}")

        self.scan_btn.config(state='disabled', text="扫描中...")

        def _worker():
            try:
                from core.api_ocr import VolcEngineOCR
                from core.vivo_workflow import VivoWorkflow, generate_game_xlsx
                api_key = self.config.get('advanced.volc_api_key', '374bc8a8-b92c-4e1c-a839-6f6d51f61b8c')
                api_ocr = VolcEngineOCR(api_key=api_key)
                wf = VivoWorkflow(config=self.config)
                infos = wf.scan_softdoc_folder(softdoc_dir, api_ocr=api_ocr, log_cb=self.log_append)
                if not infos:
                    self.log_append("未提取到任何软著信息，请检查文件夹内容")
                    return
                ok = generate_game_xlsx(infos, game_xlsx)
                if ok:
                    self.log_append(f"\n✓ game.xlsx 已生成: {game_xlsx}")
                    messagebox.showinfo("完成", f"game.xlsx 已生成！\n路径：{game_xlsx}")
                else:
                    self.log_append("✗ game.xlsx 生成失败")
            except Exception as e:
                self.log_append(f"✗ 扫描失败: {e}")
                import traceback
                traceback.print_exc()
            finally:
                self.window.after(0, lambda: self.scan_btn.config(state='normal', text="① 扫描软著 → 生成 game.xlsx"))

        threading.Thread(target=_worker, daemon=True).start()

    # ---------------------------------------------------------------- #
    # 步骤2：生成渠广 txt 文件
    # ---------------------------------------------------------------- #
    def generate_qg(self):
        game_xlsx = self.entries['game_xlsx'].get().strip()
        user_xlsx = self.entries['user_xlsx'].get().strip()
        qg_out_dir = self.entries['qg_output_dir'].get().strip()

        errors = []
        if not game_xlsx or not os.path.exists(game_xlsx):
            errors.append("game.xlsx 不存在，请先执行扫描")
        if not user_xlsx or not os.path.exists(user_xlsx):
            errors.append("user.xlsx 不存在，请先选择账号文件")
        if not qg_out_dir:
            errors.append("请先选择渠广输出文件夹")
        if errors:
            messagebox.showerror("参数错误", "\n".join(errors))
            return

        # 处理路径
        if os.path.isdir(game_xlsx):
            game_xlsx = os.path.join(game_xlsx, "game.xlsx")
        if os.path.isdir(user_xlsx):
            user_xlsx = os.path.join(user_xlsx, "user.xlsx")

        ads_config = self.get_ads_config()
        if not ads_config:
            messagebox.showwarning("提示", "未选择任何广告类型，请至少勾选一种")
            return

        self.log_append("=" * 60)
        self.log_append("开始生成渠广文件...")
        self.gen_qg_btn.config(state='disabled', text="生成中...")

        def _worker():
            try:
                from core.vivo_workflow import VivoWorkflow
                wf = VivoWorkflow(config=self.config)
                files = wf.generate_qg_files(
                    game_xlsx_path=game_xlsx,
                    user_xlsx_path=user_xlsx,
                    output_dir=qg_out_dir,
                    ads_config=ads_config,
                    log_cb=self.log_append
                )
                self.window.after(0, lambda: self._on_qg_done(files, qg_out_dir))
            except Exception as e:
                self.log_append(f"✗ 生成失败: {e}")
                import traceback
                traceback.print_exc()
                self.window.after(0, lambda: self.gen_qg_btn.config(state='normal', text="② 生成渠广 txt 文件"))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_qg_done(self, files: list, qg_out_dir: str):
        self.gen_qg_btn.config(state='normal', text="② 生成渠广 txt 文件")
        self.log_append(f"\n✓ 渠广生成完成，共 {len(files)} 个文件")
        if files:
            self.log_append("已生成：")
            for f in files:
                self.log_append(f"  {os.path.basename(f)}")
        messagebox.showinfo("完成", f"渠广文件生成完成！\n共 {len(files)} 个\n输出目录：{qg_out_dir}")

    # ---------------------------------------------------------------- #
    # 预览匹配
    # ---------------------------------------------------------------- #
    def preview_match(self):
        qg_dir = self.entries['qg_output_dir'].get().strip()
        softdoc_dir = self.entries['softdoc_dir'].get().strip()

        if not qg_dir or not os.path.isdir(qg_dir):
            messagebox.showerror("错误", "渠广输出文件夹不存在或未选择")
            return
        if not softdoc_dir or not os.path.isdir(softdoc_dir):
            messagebox.showerror("错误", "软著文件夹不存在或未选择")
            return

        self.set_preview("正在OCR识别图片并匹配，请稍候...")
        self.preview_btn.config(state='disabled')

        def _do_preview():
            try:
                from core.api_ocr import VolcEngineOCR
                api_key = self.config.get('advanced.volc_api_key', '374bc8a8-b92c-4e1c-a839-6f6d51f61b8c')
                api_ocr = VolcEngineOCR(api_key=api_key)
                matcher = BatchFileMatcher(qg_dir, softdoc_dir)
                result = matcher.match(api_ocr=api_ocr)
            except Exception as e:
                result = {'matched': [], 'qg_only': [], 'softdoc_only': []}
                print(f"预览匹配异常: {e}")

            lines = []
            lines.append(f"=== 匹配结果 ===")
            lines.append(f"成功匹配：{len(result['matched'])} 组")
            lines.append(f"渠广无软著：{len(result['qg_only'])} 个")
            lines.append(f"软著无渠广：{len(result['softdoc_only'])} 个")
            lines.append("")

            if result['matched']:
                lines.append("--- 已匹配（将生成文档）---")
                for qp, sp, gn in result['matched']:
                    lines.append(f"  [{gn}]")
                    lines.append(f"    渠广: {os.path.basename(qp)}")
                    lines.append(f"    软著: {os.path.basename(sp)}")

            if result['qg_only']:
                lines.append("")
                lines.append("--- 渠广有、软著未匹配（将跳过）---")
                for qp in result['qg_only']:
                    gn = BatchFileMatcher.extract_game_name_from_qg(qp)
                    lines.append(f"  [{gn}] {os.path.basename(qp)}")

            self.set_preview('\n'.join(lines))
            self.window.after(0, lambda: self.preview_btn.config(state='normal'))

        threading.Thread(target=_do_preview, daemon=True).start()

    # ---------------------------------------------------------------- #
    # 批量生成文档
    # ---------------------------------------------------------------- #
    def batch_generate_documents(self):
        qg_dir = self.entries['qg_output_dir'].get().strip()
        softdoc_dir = self.entries['softdoc_dir'].get().strip()
        template_dir = self.entries['template_dir'].get().strip()
        output_dir = self.entries['output_dir'].get().strip()

        errors = []
        if not qg_dir or not os.path.isdir(qg_dir):
            errors.append("渠广输出文件夹不存在或未选择")
        if not softdoc_dir or not os.path.isdir(softdoc_dir):
            errors.append("软著文件夹不存在或未选择")
        if not template_dir or not os.path.isdir(template_dir):
            errors.append("模板目录不存在或未选择")
        if not output_dir:
            errors.append("输出目录未选择")
        if errors:
            messagebox.showerror("输入错误", "\n".join(errors))
            return

        self.batch_btn.config(state='disabled', text="生成中...")
        
        def _worker():
            try:
                from core.api_ocr import VolcEngineOCR
                api_key = self.config.get('advanced.volc_api_key', '374bc8a8-b92c-4e1c-a839-6f6d51f61b8c')
                api_ocr = VolcEngineOCR(api_key=api_key)
                self._do_batch_generate(qg_dir, softdoc_dir, template_dir, output_dir, api_ocr)
            except Exception as e:
                self.log_append(f"批量生成失败: {e}")
                import traceback
                traceback.print_exc()
            finally:
                self.window.after(0, lambda: self.batch_btn.config(state='normal', text="批量生成文档"))

        threading.Thread(target=_worker, daemon=True).start()

    def _do_batch_generate(self, qg_dir, softdoc_dir, template_dir, output_dir, api_ocr):
        from core.qg_parser import QGParser
        from core.softdoc_parser import SoftDocParser
        from core.document_generator import DocumentGenerator
        
        self.log_append("正在匹配渠广文件和软著文件夹...")
        matcher = BatchFileMatcher(qg_dir, softdoc_dir)
        match_result = matcher.match(api_ocr=api_ocr)
        
        if not match_result['matched']:
            self.log_append("❌ 未找到任何匹配的渠广+软著文件对")
            return
        
        self.log_append(f"✅ 匹配成功 {len(match_result['matched'])} 组")
        
        # 打印缓存信息
        cached_ocr = match_result.get('all_ocr_texts', {})
        self.log_append(f"📦 OCR 缓存数量: {len(cached_ocr)}")
        if cached_ocr:
            self.log_append(f"   缓存键示例: {list(cached_ocr.keys())[:3]}")
        
        os.makedirs(output_dir, exist_ok=True)
        
        qg_parser = QGParser(self.config)
        
        self.log_append("创建 SoftDocParser（传递缓存）...")
        soft_parser = SoftDocParser(
            self.config, 
            external_api_ocr=api_ocr, 
            cached_ocr_texts=cached_ocr  # 传递缓存
        )
        
        generator = DocumentGenerator(self.config)
        generator.set_template_dir(template_dir)
        generator.set_output_dir(output_dir)
        
        total = len(match_result['matched'])
        success_list = []
        fail_list = []
        
        for idx, (qp, sp, game_name) in enumerate(match_result['matched'], 1):
            self.log_append(f"\n[{idx}/{total}] 处理: {game_name}")
            self.log_append(f"  渠广文件: {qp}")
            self.log_append(f"  软著文件夹: {sp}")
            
            try:
                self.log_append("  📖 解析渠广文件...")
                game_info = qg_parser.parse_file(qp)
                self.log_append(f"    游戏名: {game_info.get('game_name')}")
                self.log_append(f"    包名: {game_info.get('package_name')}")
                
                # 提取该文件夹对应的缓存
                folder_cached = {}
                for img_path, ocr_text in cached_ocr.items():
                    img_dir = os.path.dirname(img_path).replace('\\', '/')
                    sp_norm = sp.replace('\\', '/')
                    if img_dir == sp_norm:
                        folder_cached[img_path] = ocr_text
                
                self.log_append(f"  📦 该文件夹缓存数量: {len(folder_cached)}")
                self.log_append(f"  🔍 解析软著文件（使用缓存）...")
                
                soft_info = soft_parser.parse_from_folder(sp, game_name, cached_ocr_texts=folder_cached)
                
                self.log_append(f"    软件名: {soft_info.get('software_name')}")
                self.log_append(f"    著作权人: {soft_info.get('copyright_holder')}")
                
                if not soft_info.get('software_name'):
                    soft_info['software_name'] = game_info.get('game_name', '')
                if not soft_info.get('copyright_holder'):
                    soft_info['copyright_holder'] = game_info.get('publisher', '')
                
                # 收集软著文件
                import glob
                img_exts = ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp')
                softdoc_files = []
                game_clean = game_name.replace(' ', '').replace('_', '')
                
                for ext in ['pdf'] + list(img_exts):
                    for f in glob.glob(os.path.join(sp, f'*.{ext}')):
                        fname = os.path.basename(f).replace(' ', '').replace('_', '')
                        if game_clean in fname:
                            softdoc_files.append(f)
                
                softdoc_files = sorted(set(softdoc_files))
                self.log_append(f"  📄 软著文件数量: {len(softdoc_files)}")
                
                self.log_append(f"  📝 生成文档...")
                files = generator.generate_documents(game_info, soft_info, softdoc_files=softdoc_files)
                success_list.append((game_name, files))
                self.log_append(f"  ✅ 成功，生成 {len(files)} 个文件")
                
            except Exception as e:
                fail_list.append((game_name, str(e)))
                self.log_append(f"  ❌ 失败: {e}")
                import traceback
                traceback.print_exc()
        
        self.log_append("\n" + "=" * 50)
        self.log_append(f"批量生成完成！成功: {len(success_list)}, 失败: {len(fail_list)}")
        self.log_append(f"输出目录: {output_dir}")
        
        if fail_list:
            self.log_append("\n失败列表:")
            for name, err in fail_list:
                self.log_append(f"  - {name}: {err}")
        
        messagebox.showinfo("完成", f"批量生成完成！\n成功：{len(success_list)} 个\n失败：{len(fail_list)} 个")

    # ---------------------------------------------------------------- #
    # 一键生成
    # ---------------------------------------------------------------- #
    def onekey_generate(self):
        # 检查必要配置
        softdoc_dir = self.entries['softdoc_dir'].get().strip()
        user_xlsx = self.entries['user_xlsx'].get().strip()
        game_xlsx_input = self.entries['game_xlsx'].get().strip()
        qg_out_dir = self.entries['qg_output_dir'].get().strip()
        template_dir = self.entries['template_dir'].get().strip()
        output_dir = self.entries['output_dir'].get().strip()

        errors = []
        if not softdoc_dir or not os.path.isdir(softdoc_dir):
            errors.append("软著文件夹不存在或未选择")
        if not user_xlsx or not os.path.exists(user_xlsx):
            errors.append("user.xlsx 不存在，请选择账号文件")
        if not game_xlsx_input:
            errors.append("请选择 game.xlsx 保存路径")
        if not qg_out_dir:
            errors.append("请选择渠广输出文件夹")
        if not template_dir or not os.path.isdir(template_dir):
            errors.append("模板目录不存在或未选择")
        if not output_dir:
            errors.append("请选择最终文档输出目录")
        
        if errors:
            messagebox.showerror("参数错误", "\n".join(errors))
            return

        # 处理 game.xlsx 路径
        game_xlsx = os.path.normpath(game_xlsx_input)
        if os.path.isdir(game_xlsx):
            game_xlsx = os.path.join(game_xlsx, "game.xlsx")
        elif not game_xlsx.lower().endswith('.xlsx'):
            game_xlsx += ".xlsx"
        
        target_dir = os.path.dirname(game_xlsx)
        if target_dir:
            os.makedirs(target_dir, exist_ok=True)

        ads_config = self.get_ads_config()
        if not ads_config:
            messagebox.showwarning("提示", "未选择任何广告类型，请至少勾选一种")
            return

        if not messagebox.askyesno("确认", 
            f"即将执行以下步骤：\n\n"
            f"1. 扫描软著文件夹 → 生成 {game_xlsx}\n"
            f"2. 读取 game.xlsx + user.xlsx → 生成渠广文件到 {qg_out_dir}\n"
            f"3. 匹配渠广+软著 → 生成最终文档到 {output_dir}\n\n"
            f"是否继续？"):
            return

        self.onekey_btn.config(state='disabled', text="执行中...")
        self.scan_btn.config(state='disabled')
        self.gen_qg_btn.config(state='disabled')
        self.batch_btn.config(state='disabled')
        
        self.log_append("=" * 60)
        self.log_append("🚀 开始一键生成流程")
        self.log_append("=" * 60)

        def _worker():
            try:
                from core.api_ocr import VolcEngineOCR
                from core.vivo_workflow import VivoWorkflow, generate_game_xlsx
                
                api_key = self.config.get('advanced.volc_api_key', '374bc8a8-b92c-4e1c-a839-6f6d51f61b8c')
                api_ocr = VolcEngineOCR(api_key=api_key)
                wf = VivoWorkflow(config=self.config)
                
                # 步骤1：扫描软著 → 生成 game.xlsx
                self.window.after(0, lambda: self.log_append("\n【步骤 1/3】扫描软著文件夹..."))
                infos = wf.scan_softdoc_folder(softdoc_dir, api_ocr=api_ocr, log_cb=self.log_append)
                
                if not infos:
                    self.window.after(0, lambda: self.log_append("❌ 未提取到任何软著信息，流程终止"))
                    self.window.after(0, lambda: self._on_onekey_failed())
                    return
                
                ok = generate_game_xlsx(infos, game_xlsx)
                if not ok:
                    self.window.after(0, lambda: self.log_append("❌ game.xlsx 生成失败，流程终止"))
                    self.window.after(0, lambda: self._on_onekey_failed())
                    return
                
                self.window.after(0, lambda: self.log_append(f"✅ game.xlsx 已生成: {game_xlsx}"))
                
                # 步骤2：生成渠广 txt 文件
                self.window.after(0, lambda: self.log_append("\n【步骤 2/3】生成渠广文件..."))
                files = wf.generate_qg_files(
                    game_xlsx_path=game_xlsx,
                    user_xlsx_path=user_xlsx,
                    output_dir=qg_out_dir,
                    ads_config=ads_config,
                    log_cb=self.log_append
                )
                
                if not files:
                    self.window.after(0, lambda: self.log_append("❌ 渠广文件生成失败，流程终止"))
                    self.window.after(0, lambda: self._on_onekey_failed())
                    return
                
                self.window.after(0, lambda: self.log_append(f"✅ 渠广文件已生成，共 {len(files)} 个"))
                
                # 步骤3：批量生成文档
                self.window.after(0, lambda: self.log_append("\n【步骤 3/3】批量生成文档..."))
                self._do_batch_generate(qg_out_dir, softdoc_dir, template_dir, output_dir, api_ocr)
                
                self.window.after(0, lambda: self._on_onekey_finished())
                
            except Exception as e:
                self.window.after(0, lambda: self.log_append(f"❌ 一键生成失败: {e}"))
                import traceback
                traceback.print_exc()
                self.window.after(0, lambda: self._on_onekey_failed())

        threading.Thread(target=_worker, daemon=True).start()

    def _on_onekey_finished(self):
        self.onekey_btn.config(state='normal', text="🚀 一键生成（完整流程）")
        self.scan_btn.config(state='normal', text="① 扫描软著 → 生成 game.xlsx")
        self.gen_qg_btn.config(state='normal', text="② 生成渠广 txt 文件")
        self.batch_btn.config(state='normal', text="批量生成文档")
        self.log_append("\n🎉 一键生成完成！")

    def _on_onekey_failed(self):
        self.onekey_btn.config(state='normal', text="🚀 一键生成（完整流程）")
        self.scan_btn.config(state='normal', text="① 扫描软著 → 生成 game.xlsx")
        self.gen_qg_btn.config(state='normal', text="② 生成渠广 txt 文件")
        self.batch_btn.config(state='normal', text="批量生成文档")

    # ---------------------------------------------------------------- #
    # 主循环
    # ---------------------------------------------------------------- #
    def mainloop(self):
        self.window.mainloop()


def main():
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    window = MainWindow()
    window.mainloop()


if __name__ == "__main__":
    main()