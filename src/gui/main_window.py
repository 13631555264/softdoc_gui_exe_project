#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
软著文档生成工具 - 主窗口（批量模式）
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
    """
    根据游戏名称，将渠广目录下的 txt 与软著目录下的 pdf 自动配对。

    渠广文件名规范：  {游戏名}_vivo小游戏渠广.txt  （_前面是游戏名）
    软著文件名规范：  {游戏名}...pdf              （文件名以游戏名开头）
    """

    def __init__(self, qg_dir: str, softdoc_dir: str):
        self.qg_dir = qg_dir
        self.softdoc_dir = softdoc_dir

    # ---- 收集文件列表 ----
    def _list_qg_files(self) -> List[str]:
        return [
            os.path.join(self.qg_dir, f)
            for f in os.listdir(self.qg_dir)
            if f.lower().endswith('.txt')
        ]

    def _list_softdoc_files(self) -> List[str]:
        """返回软著文件夹下所有 PDF 和图片文件（用于匹配）"""
        img_exts = ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp')
        return [
            os.path.join(self.softdoc_dir, f)
            for f in os.listdir(self.softdoc_dir)
            if f.lower().endswith('.pdf') or os.path.splitext(f.lower())[1] in img_exts
        ]

    # ---- 从渠广文件名提取游戏名 ----
    @staticmethod
    def extract_game_name_from_qg(filename: str) -> str:
        """
        从渠广文件名提取游戏名。
        规则：取文件名（不含扩展名）中第一个 _ 前面的部分。
        若没有 _，则取整个文件名。
        """
        stem = os.path.splitext(os.path.basename(filename))[0]
        if '_' in stem:
            return stem.split('_')[0].strip()
        return stem.strip()

    # ---- 匹配 ----
    def match(self, api_ocr=None) -> Dict:
        """
        返回：
        {
            'matched':      [(qg_path, softdoc_dir, game_name), ...],
            'qg_only':      [qg_path, ...],           # 渠广有，软著没匹配上
            'softdoc_only': [softdoc_dir, ...],        # 软著有，渠广没匹配上
            'all_ocr_texts': {图片路径: OCR文本, ...}, # 所有 OCR 文本（供后续复用）
        }
        注意：matched / softdoc_only 中的第二项是文件夹路径（str），不是文件路径。
        parse_from_folder() 会自动扫描该文件夹下的 PDF 和所有图片。

        若 api_ocr 传入，则收集所有软著图片的 OCR 结果，供后续处理复用。
        """
        import os, glob as _glob, re as _re
        
        # ===== 调试日志：收集文件列表 =====
        qg_files = self._list_qg_files()
        softdoc_files = self._list_softdoc_files()
        
        logger.info("=" * 60)
        logger.info("【匹配调试】渠广目录: " + self.qg_dir)
        logger.info(f"【匹配调试】找到渠广文件 ({len(qg_files)} 个):")
        for qf in qg_files:
            logger.info(f"    {os.path.basename(qf)}")
        
        logger.info(f"【匹配调试】找到软著文件 ({len(softdoc_files)} 个):")
        for sf in softdoc_files:
            logger.info(f"    {os.path.basename(sf)}")
        
        # 建立软著文件名索引（basename，不含扩展名）
        softdoc_index: Dict[str, str] = {}  # basename_stem -> full_path
        for sp in softdoc_files:
            stem = os.path.splitext(os.path.basename(sp))[0]
            softdoc_index[stem] = sp
        
        logger.info(f"【匹配调试】软著文件名 stems: {list(softdoc_index.keys())}")
        
        matched = []
        qg_only = []
        used_softdoc = set()

        for qp in qg_files:
            game_name = self.extract_game_name_from_qg(qp)
            logger.info(f"【匹配调试】处理渠广文件: {os.path.basename(qp)} -> 游戏名: '{game_name}'")
            
            # 在软著文件名中寻找以游戏名开头的文件
            found = None
            for stem, sp in softdoc_index.items():
                logger.info(f"    对比: stem='{stem}' startswith '{game_name}' ? {stem.startswith(game_name)}")
                if stem.startswith(game_name):
                    found = sp
                    logger.info(f"    >>> 匹配成功! softdoc: {os.path.basename(sp)}")
                    break
            if found:
                # 找到的是文件，取其所在文件夹
                matched.append((qp, os.path.dirname(found), game_name))
                used_softdoc.add(found)
            else:
                logger.info(f"    >>> 未匹配")
                qg_only.append(qp)

        # ---------- OCR 收集（所有软著文件夹） ----------
        # 无论文件名匹配是否成功，只要有 api_ocr，就收集所有软著图片的 OCR 结果
        # 这样后续处理时可以复用，避免重复 OCR
        all_ocr_texts: Dict[str, str] = {}
        all_softdoc_dirs: Dict[str, List[str]] = {}  # dir -> [image_paths]
        
        if api_ocr:
            # 收集所有软著文件夹
            for sp in softdoc_files:
                d = os.path.dirname(sp)
                if d not in all_softdoc_dirs:
                    all_softdoc_dirs[d] = []
                all_softdoc_dirs[d].append(sp)
            
            print(f"【DEBUG】收集 OCR 文本: {len(all_softdoc_dirs)} 个文件夹")
            
            # 构建所有渠广游戏名提示
            all_game_names = set()
            for qp in qg_files:
                gn = self.extract_game_name_from_qg(qp)
                if gn:
                    all_game_names.add(gn)
            all_game_hints = '|'.join(all_game_names)
            
            for softdoc_dir, images in all_softdoc_dirs.items():
                print(f"【DEBUG】OCR 扫描文件夹: {softdoc_dir}")
                folder_game_name, image_to_game, folder_raw_texts = _build_softdoc_folder_game_names(
                    softdoc_dir, api_ocr, game_hint=all_game_hints
                )
                all_ocr_texts.update(folder_raw_texts)
                print(f"【DEBUG】OCR 结果: 游戏名='{folder_game_name}', 文本数={len(folder_raw_texts)}")
        
        # ---------- OCR 回退匹配 ----------
        # 文件名匹配失败时：从已收集的 OCR 文本中进行第二轮匹配
        if api_ocr and qg_only:
            print(f"【DEBUG】开始 OCR 回退匹配，共 {len(qg_only)} 个渠广文件需要匹配")
            logger.info(f"【匹配调试】文件名匹配后剩余 {len(qg_only)} 个渠广文件未匹配，尝试 OCR 回退...")

            # 构建游戏名 → QG文件 的反向索引
            game_to_qg: Dict[str, List[str]] = {}
            for qp in qg_only:
                gn = self.extract_game_name_from_qg(qp)
                if gn not in game_to_qg:
                    game_to_qg[gn] = []
                game_to_qg[gn].append(qp)

            # 对每个软著文件夹，从已收集的 OCR 文本中进行回退匹配
            for softdoc_dir, images in all_softdoc_dirs.items():
                # 从该文件夹的图片中获取 OCR 文本
                folder_texts = [text for img, text in all_ocr_texts.items() 
                           if os.path.dirname(img) == softdoc_dir or 
                              os.path.dirname(img) == os.path.normpath(softdoc_dir)]
                combined_text = '\n'.join(folder_texts) if folder_texts else ''
                
                # 尝试匹配 qg_only 中任一游戏名
                folder_game_name = ''
                for gn, qp_list in list(game_to_qg.items()):
                    if not qp_list:
                        continue
                    # 在 OCR 文本中搜索游戏名
                    if gn in combined_text:
                        folder_game_name = gn
                        print(f"【DEBUG】>>> OCR 回退匹配成功! '{gn}' in '{softdoc_dir}'")
                        logger.info(f"    >>> OCR 回退匹配成功! 游戏名 '{gn}'")
                        for qp in qp_list:
                            matched.append((qp, softdoc_dir, gn))
                            used_softdoc.update(images)
                        game_to_qg[gn] = []  # 该游戏已匹配
                        break
                
                if not folder_game_name:
                    print(f"【DEBUG】>>> OCR 回退匹配失败: 文件夹 '{softdoc_dir}'")

            # 剩余仍未匹配的 QG 文件
            qg_only = [qp for qps in game_to_qg.values() for qp in qps]
        elif qg_only and not api_ocr:
            print(f"【DEBUG】跳过 OCR 回退：qg_only={len(qg_only)} 但 api_ocr 为空")
            logger.warning("【匹配调试】跳过了 OCR 回退，因为 api_ocr 未传入")

        softdoc_only = [sp for sp in softdoc_files if sp not in used_softdoc]
        
        # ===== 调试日志：最终结果 =====
        logger.info("【匹配调试】===== 匹配结果汇总 =====")
        logger.info(f"  匹配成功: {len(matched)} 个")
        for m in matched:
            logger.info(f"    {os.path.basename(m[0])} <-> {os.path.basename(m[1])}")
        logger.info(f"  渠广无匹配: {len(qg_only)} 个")
        for q in qg_only:
            logger.info(f"    {os.path.basename(q)}")
        logger.info(f"  软著无匹配: {len(softdoc_only)} 个")
        logger.info("=" * 60)

        return {
            'matched': matched,
            'qg_only': qg_only,
            'softdoc_only': softdoc_only,
            'all_ocr_texts': all_ocr_texts,  # 复用 OCR 结果，避免重复调用
        }


# ------------------------------------------------------------------ #
# OCR 提取：获取图片文本并提取游戏名
# ------------------------------------------------------------------ #
import re as _re

def _ocr_image_extract_game_name(image_path: str, api_ocr, fallback_hint: str = '') -> tuple:
    """
    对单张图片做 OCR，提取完整文本并从中提取软著游戏名称。
    返回：(提取到的游戏名, 原始识别文本)
    """
    try:
        # 统一只做一次 OCR，获取完整文本
        raw_text = api_ocr.recognize_image(image_path)
        
        # 打印原始识别结果，方便调试
        print(f"\n{'='*60}")
        print(f"【OCR原始】图片: {os.path.basename(image_path)}")
        print(f"【OCR原始】内容:\n{raw_text[:1000]}{'...[截断]' if len(raw_text) > 1000 else ''}")
        print(f"{'='*60}\n")
        
        if not raw_text or not raw_text.strip():
            return ('', '')
        
        game_name = ''

        # ---------- 策略0：优先匹配 fallback_hint（渠广游戏名） ----------
        if fallback_hint:
            hints = [h.strip() for h in fallback_hint.split('|') if h.strip()]
            for hint in hints:
                # 精确匹配
                if hint in raw_text:
                    print(f"【OCR提取】策略0渠广精确匹配: '{hint}'")
                    return (hint, raw_text)
                # 模糊匹配：忽略空格/标点
                clean_raw = _re.sub(r'[\s\d\.,，。、]', '', raw_text)
                clean_hint = _re.sub(r'[\s\d\.,，。、]', '', hint)
                if clean_hint and clean_hint in clean_raw:
                    print(f"【OCR提取】策略0渠广模糊匹配: '{hint}'")
                    return (hint, raw_text)

        # ---------- 策略1：从《游戏名》格式提取 ----------
        m = _re.search(r'《([^》]{2,30})》', raw_text)
        if m:
            name = m.group(1).strip()
            if name not in ('游戏', '软件', '系统', '平台', '著作权授权书'):
                print(f"【OCR提取】策略1《》格式: '{name}'")
                return (name, raw_text)

        # ---------- 策略2：从"软件名称"字段提取 ----------
        m = _re.search(r'软件名称[：:\s]*([^\n]{2,30})', raw_text)
        if m:
            name = m.group(1).strip()
            print(f"【OCR提取】策略2软件名称字段: '{name}'")
            return (name, raw_text)

        # ---------- 策略3：取第一行 ----------
        lines = [l.strip() for l in raw_text.split('\n') if l.strip()]
        if lines:
            first = lines[0]
            if 2 < len(first) < 30 and not first.startswith('软著'):
                print(f"【OCR提取】策略3第一行: '{first}'")
                return (first, raw_text)

        # ---------- 策略4：取最短的非通用行 ----------
        for line in lines:
            line = line.strip()
            if len(line) < 2 or len(line) > 20:
                continue
            invalid_keywords = ('软著', '著作权', '登记证书', '保护条例', '授权书', '计算机')
            if any(kw in line for kw in invalid_keywords):
                continue
            if _re.match(r'^[a-zA-Z0-9\s]+$', line):
                continue
            print(f"【OCR提取】策略4最短行: '{line}'")
            return (line, raw_text)

        print(f"【OCR提取】未能提取游戏名，原始文本长度={len(raw_text)}")
        return ('', raw_text)
    except Exception as e:
        print(f"【OCR提取】异常: {e}")
        import traceback
        traceback.print_exc()
        return ('', '')


def _build_softdoc_folder_game_names(softdoc_dir: str, api_ocr, game_hint: str = '') -> tuple:
    """
    扫描软著文件夹内所有图片，OCR 提取游戏名。
    返回：(folder_game_name, image_to_game, all_raw_texts)
    """
    import glob as _glob
    softdoc_dir = os.path.normpath(softdoc_dir)  # 标准化路径
    img_exts = ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp')
    images = []
    for ext in img_exts:
        images.extend(_glob.glob(os.path.join(softdoc_dir, f'*{ext}')))
        images.extend(_glob.glob(os.path.join(softdoc_dir, f'*{ext.upper()}')))
    images = sorted(set(images))
    # 标准化所有路径（统一用正斜杠），确保缓存命中
    images = [p.replace('\\', '/') for p in images]

    folder_game_name = ''
    image_to_game: Dict[str, str] = {}
    all_raw_texts: Dict[str, str] = {}

    for img in images:
        gn, raw_text = _ocr_image_extract_game_name(img, api_ocr, fallback_hint=game_hint)
        image_to_game[img] = gn
        all_raw_texts[img] = raw_text
        if gn and not folder_game_name:
            folder_game_name = gn

    return (folder_game_name, image_to_game, all_raw_texts)


# ------------------------------------------------------------------ #
# 主窗口
# ------------------------------------------------------------------ #
class MainWindow:
    """主窗口类（批量模式）"""

    def __init__(self):
        self.config = Config()

        if HAS_DND:
            self.window = TkinterDnD.Tk()
        else:
            self.window = tk.Tk()

        self.window.title("软著文档生成工具（批量）")
        self.window.configure(bg='#f0f0f0')

        self.window_width = self.config.get('gui.window_width', 860)
        self.window_height = self.config.get('gui.window_height', 620)
        self.setup_window_geometry()

        self.processing_thread = None
        self.progress_window = None
        self.progress_label = None
        self.progress_bar = None

        self.setup_styles()
        self.create_widgets()
        self.load_config()

        logger.info("GUI 主窗口初始化完成（批量模式）")

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
        self.main_frame = ttk.Frame(self.window, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)

        # 标题
        ttk.Label(
            self.main_frame, text="软著文档生成工具（批量）",
            font=("Arial", 18, "bold")
        ).grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # ---- 文件夹选择框架 ----
        dir_frame = ttk.LabelFrame(self.main_frame, text="文件夹选择（支持拖拽）", padding="10")
        dir_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        dir_frame.columnconfigure(1, weight=1)

        rows_cfg = [
            ("渠广文件夹（含 txt）:", 'qg_dir',      'qg_dir',      self._select_qg_dir),
            ("软著文件夹（含 pdf）:", 'softdoc_dir', 'softdoc_dir', self._select_softdoc_dir),
            ("模板目录:",            'template',    'template',    self.select_template_dir),
            ("输出目录:",            'output',      'output',      self.select_output_dir),
        ]

        self._entries = {}
        for i, (label_text, attr_name, path_type, cmd) in enumerate(rows_cfg):
            ttk.Label(dir_frame, text=label_text).grid(row=i, column=0, sticky=tk.W, pady=5, padx=(0, 5))
            entry = ttk.Entry(dir_frame, width=55)
            entry.grid(row=i, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
            self.setup_drag_drop(entry, path_type)
            ttk.Button(dir_frame, text="浏览...", command=cmd, width=10).grid(row=i, column=2, padx=5)
            self._entries[attr_name] = entry

        # 提示
        ttk.Label(
            dir_frame,
            text="提示：渠广文件名格式为「游戏名_xxx.txt」，软著文件名以游戏名开头即可自动匹配",
            foreground="gray", wraplength=620, justify=tk.LEFT
        ).grid(row=len(rows_cfg), column=0, columnspan=3, pady=(5, 0), sticky=tk.W)

        # ---- 匹配预览框架 ----
        preview_frame = ttk.LabelFrame(self.main_frame, text="文件匹配预览", padding="8")
        preview_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        self.main_frame.rowconfigure(2, weight=1)

        self.preview_text = tk.Text(
            preview_frame, height=8, width=80, state='disabled',
            font=("Consolas", 9), bg="#fafafa", relief=tk.FLAT
        )
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
        self.preview_text.configure(yscrollcommand=scrollbar.set)
        self.preview_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        # ---- 高级设置 ----
        setting_frame = ttk.LabelFrame(self.main_frame, text="高级设置", padding="10")
        setting_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        ttk.Label(setting_frame, text="OCR 语言:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.ocr_combo = ttk.Combobox(setting_frame, values=['chi_sim+eng', 'chi_sim', 'eng'], width=20, state='readonly')
        self.ocr_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        self.ocr_combo.set('chi_sim+eng')

        ttk.Label(setting_frame, text="授权年限:").grid(row=0, column=2, sticky=tk.W, pady=5, padx=(20, 0))
        self.years_spin = ttk.Spinbox(setting_frame, from_=1, to=30, width=10)
        self.years_spin.grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        self.years_spin.set('10')

        # ---- 按钮区 ----
        btn_frame = ttk.Frame(self.main_frame)
        btn_frame.grid(row=4, column=0, columnspan=3, pady=(5, 0))

        self.preview_button = ttk.Button(btn_frame, text="预览匹配", command=self._preview_match, width=14)
        self.preview_button.grid(row=0, column=0, padx=5)
        self.start_button = ttk.Button(
            btn_frame, text="批量生成文档", command=self.start_processing,
            style="Accent.TButton", width=18
        )
        self.start_button.grid(row=0, column=1, padx=5, pady=10)
        ttk.Button(btn_frame, text="保存配置", command=self.save_config, width=12).grid(row=0, column=2, padx=5)

    # ---------------------------------------------------------------- #
    # 目录浏览
    # ---------------------------------------------------------------- #
    def _select_dir(self, entry_key, title, config_key):
        cur = self._entries[entry_key].get()
        directory = filedialog.askdirectory(title=title, initialdir=cur if cur else '.')
        if directory:
            self._entries[entry_key].delete(0, tk.END)
            self._entries[entry_key].insert(0, directory)
            self.config.set_last_path(config_key, directory)

    def _select_qg_dir(self):
        self._select_dir('qg_dir', '选择渠广文件夹', 'qg_dir')

    def _select_softdoc_dir(self):
        self._select_dir('softdoc_dir', '选择软著文件夹', 'softdoc_dir')

    def select_template_dir(self):
        self._select_dir('template', '选择模板目录', 'template')

    def select_output_dir(self):
        self._select_dir('output', '选择输出目录', 'output')

    # ---------------------------------------------------------------- #
    # 配置存取
    # ---------------------------------------------------------------- #
    def load_config(self):
        try:
            self._entries['qg_dir'].delete(0, tk.END)
            self._entries['qg_dir'].insert(0, self.config.get_last_path('qg_dir'))

            self._entries['softdoc_dir'].delete(0, tk.END)
            self._entries['softdoc_dir'].insert(0, self.config.get_last_path('softdoc_dir'))

            self._entries['template'].delete(0, tk.END)
            self._entries['template'].insert(0, self.config.get_last_path('template'))

            self._entries['output'].delete(0, tk.END)
            self._entries['output'].insert(0, self.config.get_last_path('output'))

            self.ocr_combo.set(self.config.get_processing_config('ocr_language', 'chi_sim+eng'))
            self.years_spin.delete(0, tk.END)
            self.years_spin.insert(0, str(self.config.get_processing_config('authorization_years', 10)))
        except Exception as e:
            logger.error(f"加载配置失败: {e}")

    def save_config(self):
        try:
            for key in ('qg_dir', 'softdoc_dir', 'template', 'output'):
                val = self._entries[key].get()
                if val:
                    self.config.set_last_path(key, val)
            self.config.set_processing_config('ocr_language', self.ocr_combo.get())
            try:
                self.config.set_processing_config('authorization_years', int(self.years_spin.get()))
            except ValueError:
                pass
            if self.config.save_config():
                messagebox.showinfo("保存成功", "配置已保存")
            else:
                messagebox.showerror("保存失败", "配置保存失败")
        except Exception as e:
            messagebox.showerror("保存失败", str(e))

    # ---------------------------------------------------------------- #
    # 匹配预览
    # ---------------------------------------------------------------- #
    def _preview_match(self):
        qg_dir = self._entries['qg_dir'].get().strip()
        softdoc_dir = self._entries['softdoc_dir'].get().strip()

        errors = []
        if not qg_dir or not os.path.isdir(qg_dir):
            errors.append("渠广文件夹不存在或未选择")
        if not softdoc_dir or not os.path.isdir(softdoc_dir):
            errors.append("软著文件夹不存在或未选择")
        if errors:
            messagebox.showerror("路径错误", "\n".join(errors))
            return

        # OCR + 匹配放到子线程里执行，避免 GUI 无响应
        self._set_preview("正在OCR识别图片并匹配，请稍候...")
        self.preview_button.config(state='disabled')

        def _do_preview():
            print("【DEBUG】_do_preview 开始执行")
            try:
                from core.api_ocr import VolcEngineOCR
                api_key = self.config.get('advanced.volc_api_key', 'a0081937-958a-44d4-8144-f713d09ada03')
                print(f"【DEBUG】API Key: {api_key[:10]}...")
                api_ocr = VolcEngineOCR(api_key=api_key)
                print(f"【DEBUG】VolcEngineOCR 实例创建成功")
                matcher = BatchFileMatcher(qg_dir, softdoc_dir)
                print(f"【DEBUG】开始调用 matcher.match()")
                result = matcher.match(api_ocr=api_ocr)
                print(f"【DEBUG】matcher.match() 完成，结果: matched={len(result.get('matched', []))}")
            except Exception as e:
                print(f"【DEBUG】_do_preview 异常: {e}")
                import traceback
                traceback.print_exc()
                result = {'matched': [], 'qg_only': [], 'softdoc_only': []}

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

            if result['softdoc_only']:
                lines.append("")
                lines.append("--- 软著有、渠广未匹配（不处理）---")
                for sp in result['softdoc_only']:
                    lines.append(f"  {os.path.basename(sp)}")

            # 匹配结果回写到 GUI
            text = '\n'.join(lines)
            self.window.after(0, lambda: self._show_preview_result(text))

        threading.Thread(target=_do_preview, daemon=True).start()

    def _show_preview_result(self, text: str):
        self._set_preview(text)
        self.preview_button.config(state='normal')

    def _set_preview(self, text: str):
        self.preview_text.config(state='normal')
        self.preview_text.delete('1.0', tk.END)
        self.preview_text.insert(tk.END, text)
        self.preview_text.config(state='disabled')

    # ---------------------------------------------------------------- #
    # 批量处理入口
    # ---------------------------------------------------------------- #
    def start_processing(self):
        qg_dir = self._entries['qg_dir'].get().strip()
        softdoc_dir = self._entries['softdoc_dir'].get().strip()
        template_dir = self._entries['template'].get().strip()
        output_dir = self._entries['output'].get().strip()
        ocr_language = self.ocr_combo.get()
        try:
            auth_years = int(self.years_spin.get())
        except ValueError:
            auth_years = 10

        errors = []
        if not qg_dir or not os.path.isdir(qg_dir):
            errors.append("渠广文件夹不存在或未选择")
        if not softdoc_dir or not os.path.isdir(softdoc_dir):
            errors.append("软著文件夹不存在或未选择")
        if not template_dir or not os.path.isdir(template_dir):
            errors.append("模板目录不存在或未选择")
        if not output_dir:
            errors.append("输出目录未选择")
        if errors:
            messagebox.showerror("输入错误", "\n".join(errors))
            return

        # 禁用按钮，提示用户正在匹配
        self.start_button.config(state='disabled', text="匹配中...")

        def _do_match():
            # 子线程里做 OCR + 匹配，避免卡 GUI
            print("【DEBUG】_do_match 开始执行")
            try:
                from core.api_ocr import VolcEngineOCR
                api_key = self.config.get('advanced.volc_api_key', 'a0081937-958a-44d4-8144-f713d09ada03')
                print(f"【DEBUG】API Key: {api_key[:10]}...")
                api_ocr = VolcEngineOCR(api_key=api_key)
                print(f"【DEBUG】VolcEngineOCR 实例创建成功: {type(api_ocr)}")
                matcher = BatchFileMatcher(qg_dir, softdoc_dir)
                print(f"【DEBUG】开始调用 matcher.match()")
                result = matcher.match(api_ocr=api_ocr)
                print(f"【DEBUG】matcher.match() 完成，结果: matched={len(result.get('matched', []))}")
            except Exception as e:
                print(f"【DEBUG】_do_match 异常: {e}")
                import traceback
                traceback.print_exc()
                result = {'matched': [], 'qg_only': [], 'softdoc_only': []}
            # match 结果回到主线程，决定后续动作
            self.window.after(0, lambda r=result: self._on_match_done(
                r, template_dir, output_dir, ocr_language, auth_years, api_ocr
            ))

        threading.Thread(target=_do_match, daemon=True).start()

    def _on_match_done(self, match_result, template_dir, output_dir,
                       ocr_language, auth_years, api_ocr):
        """match 结果回到主线程：无匹配则弹警告，有匹配则开始处理"""
        if not match_result['matched']:
            self.start_button.config(state='normal', text="批量生成文档")
            
            # ===== 在预览区显示详细的匹配信息，帮助调试 =====
            lines = []
            lines.append("=== 匹配结果（调试信息）===")
            lines.append(f"渠广目录: {self._entries['qg_dir'].get()}")
            lines.append(f"软著目录: {self._entries['softdoc_dir'].get()}")
            lines.append("")
            lines.append(f"渠广文件总数: {len(match_result['qg_only'])}")
            for qp in match_result['qg_only']:
                gn = BatchFileMatcher.extract_game_name_from_qg(qp)
                lines.append(f"  未匹配: {os.path.basename(qp)} (提取游戏名: '{gn}')")
            
            lines.append("")
            lines.append(f"软著文件总数: {len(match_result['softdoc_only'])}")
            for sp in match_result['softdoc_only']:
                lines.append(f"  未匹配: {os.path.basename(sp)}")
            
            lines.append("")
            lines.append(">>> 请查看上方日志或控制台，了解详细匹配过程")
            lines.append(">>> 匹配规则：软著文件名必须以渠广提取的游戏名开头")
            
            self._set_preview('\n'.join(lines))
            
            # 同时打印到控制台
            import sys
            detail = '\n'.join(lines)
            print(detail)
            if hasattr(sys, 'stderr'):
                sys.stderr.write(detail + '\n')
            
            messagebox.showwarning("无可处理文件",
                "未找到任何匹配的渠广+软著文件对，请查看上方预览区的调试信息。")
            return

        os.makedirs(output_dir, exist_ok=True)

        # 禁用按钮，打开进度窗口
        self.start_button.config(state='disabled', text="处理中...")

        self.progress_window = tk.Toplevel(self.window)
        self.progress_window.title("批量处理中")
        self.progress_window.geometry("440x170")
        self.progress_window.transient(self.window)
        self.progress_window.grab_set()
        x = self.window.winfo_x() + (self.window.winfo_width() - 440) // 2
        y = self.window.winfo_y() + (self.window.winfo_height() - 170) // 2
        self.progress_window.geometry(f"440x170+{x}+{y}")

        self.progress_label = ttk.Label(self.progress_window, text="正在初始化...", font=("Arial", 10))
        self.progress_label.pack(pady=(20, 5))

        self.progress_sub_label = ttk.Label(self.progress_window, text="", font=("Arial", 9), foreground="gray")
        self.progress_sub_label.pack()

        self.progress_bar = ttk.Progressbar(self.progress_window, mode='determinate', length=380)
        self.progress_bar.pack(pady=10)

        self.processing_thread = threading.Thread(
            target=self._batch_process,
            args=(match_result['matched'], template_dir, output_dir, ocr_language, auth_years, api_ocr,
                  match_result.get('all_ocr_texts', {})),
            daemon=True
        )
        self.processing_thread.start()

    # ---------------------------------------------------------------- #
    # 批量处理后台线程
    # ---------------------------------------------------------------- #
    def _batch_process(self, matched: list, template_dir: str, output_dir: str,
                       ocr_language: str, auth_years: int, api_ocr=None,
                       cached_ocr_texts: Dict[str, str] = None):
        import time
        from core.qg_parser import QGParser
        from core.softdoc_parser import SoftDocParser
        from core.document_generator import DocumentGenerator

        step_timer = time.time()
        print(f"\n{'='*60}")
        print(f"【BATCH DEBUG】批处理开始，共 {len(matched)} 个匹配项")
        print(f"【BATCH DEBUG】cached_ocr_texts 传入: {len(cached_ocr_texts) if cached_ocr_texts else 0} 个")
        print(f"{'='*60}\n")

        self.config.set_processing_config('ocr_language', ocr_language)
        self.config.set_processing_config('authorization_years', auth_years)

        total = len(matched)
        success_list = []
        fail_list = []

        qg_parser = QGParser(self.config)
        soft_parser = SoftDocParser(self.config, external_api_ocr=api_ocr, cached_ocr_texts=cached_ocr_texts)
        generator = DocumentGenerator(self.config)
        generator.set_template_dir(template_dir)
        generator.set_output_dir(output_dir)

        for idx, (qp, sp, game_name) in enumerate(matched, 1):
            item_timer = time.time()
            print(f"\n--- [{idx}/{total}] 开始处理: {game_name} ---")

            self._update_progress_batch(
                f"正在处理 ({idx}/{total})：{game_name}",
                f"渠广: {os.path.basename(qp)}",
                int((idx - 1) / total * 100)
            )

            try:
                # ===== 1. 解析渠广 =====
                step_timer = time.time()
                print(f"【{idx}】[1/4] 解析渠广: {os.path.basename(qp)}")
                game_info = qg_parser.parse_file(qp)
                print(f"【{idx}】[1/4] 渠广解析完成，耗时: {time.time()-step_timer:.2f}s")

                # ===== 2. 筛选属于当前文件夹的 OCR 缓存 =====
                step_timer = time.time()
                folder_cached = {}
                if cached_ocr_texts:
                    import os as _os
                    # 标准化路径用于比较（统一用正斜杠）
                    sp_norm = sp.replace('\\', '/')
                    for img_path, ocr_text in cached_ocr_texts.items():
                        img_dir = _os.path.dirname(img_path).replace('\\', '/')
                        if img_dir == sp_norm:
                            folder_cached[img_path] = ocr_text
                    print(f"【{idx}】[2/4] OCR缓存筛选: {len(folder_cached)}/{len(cached_ocr_texts)} 个文件命中")
                else:
                    print(f"【{idx}】[2/4] OCR缓存为空，跳过筛选")
                print(f"【{idx}】[2/4] 筛选完成，耗时: {time.time()-step_timer:.2f}s")

                # ===== 3. 解析软著 =====
                step_timer = time.time()
                print(f"【{idx}】[3/4] 开始解析软著文件夹: {sp}")
                try:
                    soft_info = soft_parser.parse_from_folder(sp, game_name, cached_ocr_texts=folder_cached)
                    print(f"【{idx}】[3/4] 软著解析完成，耗时: {time.time()-step_timer:.2f}s")
                except Exception as e:
                    print(f"【{idx}】[3/4] 软著解析异常: {e}，使用默认值")
                    logger.warning(f"软著解析失败，使用默认值: {e}")
                    soft_info = {
                        'software_name': game_info.get('game_name', ''),
                        'version': game_info.get('version', ''),
                        'copyright_holder': game_info.get('publisher', ''),
                        'software_type': '', 'registration_number': '',
                        'completion_date': '', 'publish_date': '', 'original_text': ''
                    }

                # 补充缺失
                if not soft_info.get('software_name'):
                    soft_info['software_name'] = game_info.get('game_name', '')
                if not soft_info.get('copyright_holder'):
                    soft_info['copyright_holder'] = game_info.get('publisher', '')

                # ===== 4. 生成文档 =====
                step_timer = time.time()
                print(f"【{idx}】[4/4] 开始生成文档...")
                files = generator.generate_documents(game_info, soft_info)
                print(f"【{idx}】[4/4] 文档生成完成，耗时: {time.time()-step_timer:.2f}s")

                success_list.append((game_name, files))
                logger.info(f"[{idx}/{total}] 成功: {game_name}, 生成 {len(files)} 个文件")
                print(f"--- [{idx}/{total}] 完成，总耗时: {time.time()-item_timer:.2f}s ---\n")

            except Exception as e:
                import traceback
                fail_list.append((game_name, str(e)))
                logger.error(f"[{idx}/{total}] 失败: {game_name} — {e}")
                traceback.print_exc()
                print(f"--- [{idx}/{total}] 失败，耗时: {time.time()-item_timer:.2f}s ---\n")

        # 全部完成
        self.window.after(
            0, self._on_batch_finished,
            success_list, fail_list, output_dir, total
        )

    def _update_progress_batch(self, main_text: str, sub_text: str, pct: int):
        def _do():
            if self.progress_label:
                self.progress_label.config(text=main_text)
            if hasattr(self, 'progress_sub_label') and self.progress_sub_label:
                self.progress_sub_label.config(text=sub_text)
            if self.progress_bar:
                self.progress_bar['value'] = pct
        self.window.after(0, _do)

    def update_progress(self, message):
        def _update():
            if self.progress_label:
                self.progress_label.config(text=message)
        self.window.after(0, _update)

    # ---------------------------------------------------------------- #
    # 批量完成回调
    # ---------------------------------------------------------------- #
    def _on_batch_finished(self, success_list, fail_list, output_dir, total):
        if self.progress_window:
            self.progress_window.destroy()
            self.progress_window = None

        self.start_button.config(state='normal', text="批量生成文档")

        # 更新预览区
        lines = [f"=== 处理完成 ({total} 组) ===", ""]
        lines.append(f"成功：{len(success_list)} 个")
        for name, files in success_list:
            lines.append(f"  [OK] {name}  ({len(files)} 个文件)")

        if fail_list:
            lines.append("")
            lines.append(f"失败：{len(fail_list)} 个")
            for name, err in fail_list:
                lines.append(f"  [!!] {name}")
                lines.append(f"       原因: {err}")

        self._set_preview('\n'.join(lines))

        # 弹出汇总
        msg = f"批量处理完成！\n\n成功：{len(success_list)} 个\n失败：{len(fail_list)} 个\n\n输出目录：{output_dir}"
        if fail_list:
            msg += f"\n\n失败列表：\n" + "\n".join(f"  {n}" for n, _ in fail_list)
        messagebox.showinfo("完成", msg)

        if os.path.exists(output_dir):
            os.startfile(output_dir)

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
