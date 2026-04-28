# core/softdoc_parser.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
软著文件解析器
功能：解析软著 PDF 或图片文件
- PDF：使用 pdfplumber 直接提取文字
- 图片：使用火山引擎 API 识别文字
"""

import os
import re
import logging
from typing import Dict, Optional

# PDF 解析库
import pdfplumber

# 图片处理
from PIL import Image

# API OCR
from .api_ocr import VolcEngineOCR, extract_soft_info_from_text

from .config import Config

logger = logging.getLogger("softdoc_generator")


class SoftDocParser:
    """软著文件解析器 - 支持 PDF 和图片"""
    
    def __init__(self, config: Optional[Config] = None, external_api_ocr=None, cached_ocr_texts: Dict = None):
        self.config = config or Config()
        # 优先复用外部传入的 OCR 实例，避免重复创建
        if external_api_ocr is not None:
            self.api_ocr = external_api_ocr
        else:
            api_key = self.config.get('advanced.volc_api_key', 'a0081937-958a-44d4-8144-f713d09ada03')
            self.api_ocr = VolcEngineOCR(api_key=api_key)
        # 复用 match 阶段已 OCR 的文本，避免重复调用 API
        self.cached_ocr_texts: Dict[str, str] = cached_ocr_texts or {}
    
    def parse_file(self, file_path: str) -> Dict:
        """解析软著文件（支持 PDF 和图片）"""
        if not os.path.exists(file_path):
            logger.error(f"软著文件不存在: {file_path}")
            raise FileNotFoundError(f"软著文件不存在: {file_path}")
        
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext == '.pdf':
            # PDF：直接提取文字
            text = self._parse_pdf(file_path)
        elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp']:
            # 图片：使用 API 识别
            text = self._parse_image_with_api(file_path)
        else:
            raise ValueError(f"不支持的文件类型: {ext}，请使用 PDF 或图片格式")
        
        if not text or not text.strip():
            logger.warning("未能提取到文字内容")
            return self._get_empty_result()
        
        # 提取信息
        result = self._extract_soft_info(text)
        result['original_text'] = text
        
        return result
    
    def _parse_pdf(self, pdf_path: str) -> str:
        """解析 PDF 文件（直接提取文本）"""
        logger.info(f"正在解析 PDF: {pdf_path}")
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                all_text = []
                for page_num, page in enumerate(pdf.pages, 1):
                    page_text = page.extract_text()
                    if page_text and page_text.strip():
                        all_text.append(page_text)
                        logger.info(f"第 {page_num} 页提取 {len(page_text)} 字符")
                    else:
                        logger.warning(f"第 {page_num} 页无文字内容，可能是扫描件")
                
                full_text = '\n'.join(all_text)
                logger.info(f"PDF 解析完成，共 {len(pdf.pages)} 页，提取 {len(full_text)} 字符")
                
                if not full_text.strip():
                    logger.warning("PDF 未提取到文字，可能是扫描件，请使用图片格式")
                
                return full_text
                
        except Exception as e:
            logger.error(f"PDF 解析失败: {e}")
            raise
    
    def _parse_image_with_api(self, image_path: str) -> str:
        """使用火山引擎 API 识别图片中的文字（优先复用缓存）"""
        # 检查是否有已缓存的 OCR 结果
        cached = self.cached_ocr_texts.get(image_path)
        if cached is not None:
            logger.info(f"复用 OCR 缓存: {image_path}")
            return cached
        
        logger.info(f"正在使用 API 识别图片: {image_path}")
        
        try:
            # 调用 API 识别
            text = self.api_ocr.recognize_image(image_path)
            
            if text and text.strip():
                logger.info(f"API 识别成功，获取 {len(text)} 字符")
                return text
            else:
                logger.warning("API 识别返回空内容")
                return ""
                
        except Exception as e:
            logger.error(f"API 识别失败: {e}")
            raise
    
    def _extract_soft_info(self, text: str) -> Dict:
        """从文本中提取软著信息"""
        result = {
            'software_name': '',
            'version': '',
            'copyright_holder': '',
            'software_type': '',
            'registration_number': '',
            'completion_date': '',
            'publish_date': '',
            'original_text': text
        }
        
        if not text:
            return result
        
        # 提取软件名称
        patterns = [
            r'软件名称[：:]\s*([^\n]+)',
            r'软件全称[：:]\s*([^\n]+)',
            r'产品名称[：:]\s*([^\n]+)',
            r'《([^》]+)》',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                result['software_name'] = match.group(1).strip()
                break
        
        # 如果没有找到，尝试取第一行
        if not result['software_name']:
            lines = [l.strip() for l in text.split('\n') if l.strip()]
            if lines:
                result['software_name'] = lines[0][:100]
        
        # 提取版本号
        patterns = [
            r'版本号[：:]\s*([Vv]?\d+\.\d+(?:\.\d+)?)',
            r'版本[：:]\s*([Vv]?\d+\.\d+(?:\.\d+)?)',
            r'[Vv](\d+\.\d+(?:\.\d+)?)',
        ]
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                result['version'] = match.group(1).strip()
                break
        
        # 提取著作权人
        # PDF 提取的文字可能：
        #   ① 字间有空格："企 业 名 称:   深圳xxx有限公司"
        #   ② 所有字段挤在一行，无换行："...深圳xxx有限公司证书号: xxx首次发表日期: ..."
        # 策略：先对全文压缩中文字间空格得到 text_compact，
        #       再匹配著作权人字段，最终截断到"公司/科技/有限"等关键词结尾，避免把后续字段带进去。
        text_compact = re.sub(r'(?<=[^\x00-\x7F])\s+(?=[^\x00-\x7F])', '', text)

        # 后续字段名（用于截断）
        _field_stop = r'(?=证书号|认证号|登记号|首次发表|权利获取|权利范围|发表日期|完成日期|开发完成|\n|$)'

        def _clean_holder(raw: str) -> str:
            """清理著作权人原始文本：去前缀、截断到公司名结尾"""
            s = raw.strip()
            # 去除"企 业 名 称:"等前缀（支持字间空格）
            s = re.sub(r'^[\u4e00-\u9fff\s]*企[\s]*业[\s]*名[\s]*称[\s]*[：:\s]*', '', s).strip()
            # 截断到第一个下一字段名之前（PDF单行情况）
            m = re.search(r'证书号|认证号|登记号|首次发表|权利获取|权利范围|发表日期|完成日期|开发完成', s)
            if m:
                s = s[:m.start()].strip()
            # 去除末尾标点
            s = re.sub(r'[。；，、：:、\s]+$', '', s).strip()
            return s

        patterns = [
            r'著作权人[：:]\s*(.+?)' + _field_stop,
            r'申请人[：:]\s*(.+?)' + _field_stop,
            r'权利人[：:]\s*(.+?)' + _field_stop,
            r'版权人[：:]\s*(.+?)' + _field_stop,
            r'([^\n]{2,40}(?:有限公司|股份公司|科技公司|网络科技))',
        ]
        for pattern in patterns:
            for search_text in (text_compact, text):
                match = re.search(pattern, search_text)
                if match:
                    holder = _clean_holder(match.group(1))
                    if 2 < len(holder) < 60:
                        result['copyright_holder'] = holder
                        break
            if result['copyright_holder']:
                break

        # 提取软著登记号
        # 真实格式（PDF单行，无换行）：
        #   "...认证号: 2026SA0056162"  ← 优先取这个
        #   "软著认000463696号"          ← 证书号，不取
        # 注意：[^\n]+ 在单行 PDF 里会吃掉整行，改用截断到下一字段名
        _num_stop = r'(?=\s*(?:首次发表|权利获取|权利范围|发表日期|完成日期|证书号|\n|$))'
        patterns = [
            (r'认证号[\s:：]+(\S+)' + _num_stop, False),          # 认证号: 2026SA0056162

            (r'登记号[：:]\s*(\S+)' + _num_stop, False),
            (r'(2\d{3}S[A-Z]\d+)', False),                        # 如 2026SA0056162 / 2024SR0099999
        ]
        for pattern, is_cert in patterns:
            match = re.search(pattern, text_compact) or re.search(pattern, text)
            if match:
                result['registration_number'] = match.group(1).strip()
                break
        
        logger.info(f"解析结果: 软件={result['software_name']}, 著作权人={result['copyright_holder']}")
        return result
    
    def _get_empty_result(self) -> Dict:
        """返回空结果"""
        return {
            'software_name': '',
            'version': '',
            'copyright_holder': '',
            'software_type': '',
            'registration_number': '',
            'completion_date': '',
            'publish_date': '',
            'original_text': ''
        }
    
    def parse_from_folder(self, folder_path: str, game_name_hint: str = '', 
                          cached_ocr_texts: Dict[str, str] = None) -> Dict:
        """
        解析软著文件夹，自动识别其中的 PDF 和图片文件并提取信息。
        优先用 PDF 文字；若无文字则扫描文件夹内所有图片进行 OCR 识别。
        多张图片时按文件名排序后逐张识别并合并文字。
        
        Args:
            folder_path: 软著文件夹路径
            game_name_hint: 游戏名提示
            cached_ocr_texts: 预缓存的 OCR 文本字典 {图片路径: OCR文本}，
                             如果传入，则优先使用缓存避免重复 OCR
        """
        import glob
        import os

        folder_path = os.path.abspath(folder_path)
        # 标准化路径格式（统一用正斜杠），避免 K:\xxx 和 K:/xxx 不匹配
        folder_path_norm = folder_path.replace('\\', '/')
        
        print(f"\n【PARSER DEBUG】parse_from_folder: {folder_path}")
        print(f"【PARSER DEBUG】标准化路径: {folder_path_norm}")
        print(f"【PARSER DEBUG】self.cached_ocr_texts 数量: {len(self.cached_ocr_texts) if self.cached_ocr_texts else 0}")
        print(f"【PARSER DEBUG】传入的 cached_ocr_texts 数量: {len(cached_ocr_texts) if cached_ocr_texts else 0}")
        logger.info(f"解析软著文件夹: {folder_path}")

        all_text_parts = []

        # 1. 扫描文件夹内所有支持的文件
        supported_exts = ('.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp')
        all_files = []
        for ext in supported_exts:
            all_files.extend(glob.glob(os.path.join(folder_path, f'*{ext}')))
            all_files.extend(glob.glob(os.path.join(folder_path, f'*{ext.upper()}')))
        all_files = sorted(set(all_files))  # 去重+排序

        if not all_files:
            logger.warning(f"文件夹内未找到软著文件: {folder_path}")
            return self._get_empty_result()

        print(f"【PARSER DEBUG】扫描到 {len(all_files)} 个文件:")
        for f in all_files:
            print(f"    {os.path.basename(f)}")
        logger.info(f"找到 {len(all_files)} 个文件: {[os.path.basename(f) for f in all_files]}")

        # 2. 如果有游戏名提示，按游戏名筛选 PDF 文件
        matched_files = []
        if game_name_hint:
            print(f"【PARSER DEBUG】使用游戏名筛选: '{game_name_hint}'")
            for f in all_files:
                fname = os.path.basename(f)
                # 文件名中包含游戏名（忽略空格、下划线）
                fname_clean = fname.replace(' ', '').replace('_', '')
                hint_clean = game_name_hint.replace(' ', '').replace('_', '')
                if hint_clean in fname_clean:
                    matched_files.append(f)
                    print(f"【PARSER DEBUG】  匹配: {fname}")
            if matched_files:
                print(f"【PARSER DEBUG】通过文件名匹配筛选出 {len(matched_files)} 个文件")
            else:
                print(f"【PARSER DEBUG】文件名匹配失败，使用全部文件作为后备")
                matched_files = all_files
        else:
            matched_files = all_files

        # 3. 处理匹配到的文件
        for fpath in matched_files:
            ext = os.path.splitext(fpath)[1].lower()
            try:
                # 标准化路径用于缓存查询
                fpath_norm = os.path.normpath(fpath).replace('\\', '/')
                
                if ext == '.pdf':
                    text = self._parse_pdf(fpath)
                    if text and text.strip():
                        all_text_parts.append(text)
                        logger.info(f"  PDF 文字提取成功: {len(text)} 字符")
                    else:
                        logger.info(f"  PDF 无文字内容，将尝试 OCR 图片...")
                else:
                    # 图片：优先使用预缓存的 OCR 结果
                    print(f"【PARSER DEBUG】处理图片原始: {fpath}")
                    print(f"【PARSER DEBUG】标准化路径: {fpath_norm}")
                    
                    # 打印缓存的 keys 用于调试
                    if cached_ocr_texts:
                        print(f"【PARSER DEBUG】传入缓存 keys: {list(cached_ocr_texts.keys())}")
                    if self.cached_ocr_texts:
                        print(f"【PARSER DEBUG】self缓存 keys: {list(self.cached_ocr_texts.keys())}")
                    
                    # 尝试在缓存中查找（尝试多种路径格式）
                    found_cache = None
                    # 1. 优先用传入的 cached_ocr_texts
                    if cached_ocr_texts:
                        if fpath_norm in cached_ocr_texts:
                            found_cache = cached_ocr_texts[fpath_norm]
                            print(f"【PARSER DEBUG】>>> 传入缓存命中! (fpath_norm)")
                        elif fpath in cached_ocr_texts:
                            found_cache = cached_ocr_texts[fpath]
                            print(f"【PARSER DEBUG】>>> 传入缓存命中! (fpath原始)")
                    # 2. 其次用 self.cached_ocr_texts
                    if found_cache is None and self.cached_ocr_texts:
                        if fpath_norm in self.cached_ocr_texts:
                            found_cache = self.cached_ocr_texts[fpath_norm]
                            print(f"【PARSER DEBUG】>>> self缓存命中! (fpath_norm)")
                        elif fpath in self.cached_ocr_texts:
                            found_cache = self.cached_ocr_texts[fpath]
                            print(f"【PARSER DEBUG】>>> self缓存命中! (fpath原始)")
                    
                    if found_cache is not None:
                        text = found_cache
                        logger.info(f"  图片使用缓存 OCR: {os.path.basename(fpath)}, {len(text)} 字符")
                    else:
                        print(f"【PARSER DEBUG】>>> 缓存未命中，调用 API!")
                        logger.info(f"  图片 OCR 识别: {os.path.basename(fpath)}")
                        text = self._parse_image_with_api(fpath)
                    if text and text.strip():
                        all_text_parts.append(text)
                        logger.info(f"  内容提取成功: {len(text)} 字符")
            except Exception as e:
                logger.warning(f"  处理失败 {os.path.basename(fpath)}: {e}")
                continue

        if not all_text_parts:
            logger.warning("文件夹内所有文件均未能提取到文字")
            return self._get_empty_result()

        # 合并所有文字
        combined_text = '\n'.join(all_text_parts)
        logger.info(f"合计提取 {len(combined_text)} 字符")

        result = self._extract_soft_info(combined_text)
        result['original_text'] = combined_text

        # 用游戏名hint补充空字段
        if not result.get('software_name') and game_name_hint:
            result['software_name'] = game_name_hint

        return result

    def validate_result(self, result: Dict) -> bool:
        """验证解析结果"""
        if result.get('software_name') and result.get('copyright_holder'):
            return True
        logger.warning(f"软著信息不完整: 软件名={result.get('software_name')}, 著作权人={result.get('copyright_holder')}")
        return False


def main():
    """测试函数"""
    import sys
    logging.basicConfig(level=logging.INFO)
    
    parser = SoftDocParser()
    
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        try:
            result = parser.parse_file(file_path)
            print("\n" + "=" * 50)
            print("解析结果:")
            print("=" * 50)
            print(f"软件名称: {result['software_name']}")
            print(f"版本号: {result['version']}")
            print(f"著作权人: {result['copyright_holder']}")
            print(f"登记号: {result['registration_number']}")
        except Exception as e:
            print(f"解析失败: {e}")
    else:
        print("请提供文件路径: python softdoc_parser.py <文件路径>")


if __name__ == "__main__":
    main()