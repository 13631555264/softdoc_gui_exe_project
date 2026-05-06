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
            api_key = self.config.get('advanced.volc_api_key', '374bc8a8-b92c-4e1c-a839-6f6d51f61b8c')
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
        
        print(text)
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
        
        # 先清理文本：去除多余空格，但保留中文字符间的空格（因为 OCR 可能产生）
        # 对于提取关键信息，我们同时使用原始文本和清理后的文本
        
        # 1. 提取软件名称
        # 模式1: "软 件 名 称 : xxx" 格式（OCR常见）
        patterns_software = [
            r'软\s*件\s*名\s*称\s*[：:]\s*([^\n\[\]]+)',
            r'软件全称[：:]\s*([^\n]+)',
            r'APP\s*软件名称[：:]\s*([^\n]+)',
            r'《([^》]+)》',
        ]
        for pattern in patterns_software:
            match = re.search(pattern, text)
            if match:
                name = match.group(1).strip()
                # 过滤无效名称
                invalid_names = ['计算机软件保护条例', '计算机软件', '软件保护条例']
                if name not in invalid_names and len(name) < 100:
                    result['software_name'] = name
                    break
        
        # 2. 提取著作权人（关键修复）
        patterns_copyright = [
            # 标准格式："著 作 权 人 : xxx"
            r'著\s*作\s*权\s*人\s*[：:]\s*([^\n]+)',
            # "著作权人：" 无空格
            r'著作权人[：:]\s*([^\n]+)',
            # "企 业 名 称: xxx"
            r'企\s*业\s*名\s*称\s*[：:]\s*([^\n]+)',
            # "企业名称：" 无空格
            r'企业名称[：:]\s*([^\n]+)',
        ]
        for pattern in patterns_copyright:
            match = re.search(pattern, text)
            if match:
                holder = match.group(1).strip()
                # 清理可能的后缀和空格
                holder = re.sub(r'[\s]+', '', holder)  # 移除所有空格
                # 验证是否为有效的公司名（包含有限公司等关键词）
                if len(holder) >= 4 and ('有限' in holder or '公司' in holder or '工作室' in holder):
                    result['copyright_holder'] = holder
                    logger.info(f"提取到著作权人: {holder}")
                    break
        
        # 如果上面没找到，尝试在文本中搜索公司名模式
        if not result['copyright_holder']:
            # 查找包含"有限公司"的行
            lines = text.split('\n')
            for line in lines:
                if '有限公司' in line or '科技公司' in line or '网络科技' in line:
                    # 清理行内容，提取公司名
                    # 移除常见前缀
                    cleaned = re.sub(r'^.*?[：:]\s*', '', line)
                    cleaned = re.sub(r'[\s]+', '', cleaned)
                    if len(cleaned) >= 6:
                        result['copyright_holder'] = cleaned
                        logger.info(f"从行提取到著作权人: {cleaned}")
                        break
        
        # 3. 提取登记号/认证号
        patterns_reg = [
            r'认\s*证\s*号\s*[：:]\s*([A-Z0-9]+)',
            r'认证号[：:]\s*([A-Z0-9]+)',
            r'登记号[：:]\s*([A-Z0-9]+)',
            r'证书号[：:]\s*([A-Z0-9]+)',
            r'(202[4-9]SA\d{7})',  # 2026SA0027426 格式
            r'(202[4-9]SR\d{7})',
        ]
        for pattern in patterns_reg:
            match = re.search(pattern, text)
            if match:
                result['registration_number'] = match.group(1).strip()
                break
        
        # 4. 提取版本号
        patterns_version = [
            r'版\s*本\s*号\s*[：:]\s*([Vv]?\d+\.\d+(?:\.\d+)?)',
            r'版本[：:]\s*([Vv]?\d+\.\d+(?:\.\d+)?)',
            r'APP\s*版\s*本\s*号\s*[：:]\s*([Vv]?\d+\.\d+(?:\.\d+)?)',
            r'([Vv]\d+\.\d+)',
        ]
        for pattern in patterns_version:
            match = re.search(pattern, text)
            if match:
                result['version'] = match.group(1).strip()
                break
        
        # 5. 提取完成日期
        patterns_date = [
            r'开\s*发\s*完\s*成\s*日\s*期\s*[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)',
            r'完成日期[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)',
            r'(\d{4}-\d{1,2}-\d{1,2})',
        ]
        for pattern in patterns_date:
            match = re.search(pattern, text)
            if match:
                result['completion_date'] = match.group(1).strip()
                break
        
        logger.info(f"解析结果: 软件={result['software_name']}, 著作权人={result['copyright_holder']}, 登记号={result['registration_number']}")
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
            
    # core/softdoc_parser.py - 添加详细日志的 parse_from_folder 方法

    def parse_from_folder(self, folder_path: str, game_name_hint: str = '', 
                        cached_ocr_texts: Dict[str, str] = None) -> Dict:
        """
        解析软著文件夹，自动识别其中的 PDF 和图片文件并提取信息。
        """
        import glob
        import os

        folder_path = os.path.abspath(folder_path)
        folder_path_norm = folder_path.replace('\\', '/')
        
        print(f"\n{'='*60}")
        print(f"【PARSER DEBUG】parse_from_folder 调用")
        print(f"  文件夹: {folder_path}")
        print(f"  游戏名提示: '{game_name_hint}'")
        print(f"  传入的 cached_ocr_texts 数量: {len(cached_ocr_texts) if cached_ocr_texts else 0}")
        print(f"  传入的 cached_ocr_texts 键: {list(cached_ocr_texts.keys()) if cached_ocr_texts else []}")
        print(f"  self.cached_ocr_texts 数量: {len(self.cached_ocr_texts) if self.cached_ocr_texts else 0}")
        logger.info(f"解析软著文件夹: {folder_path}")

        all_text_parts = []

        # 扫描文件夹内所有支持的文件
        supported_exts = ('.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp')
        all_files = []
        for ext in supported_exts:
            all_files.extend(glob.glob(os.path.join(folder_path, f'*{ext}')))
            all_files.extend(glob.glob(os.path.join(folder_path, f'*{ext.upper()}')))
        all_files = sorted(set(all_files))

        print(f"  扫描到 {len(all_files)} 个文件: {[os.path.basename(f) for f in all_files]}")

        if not all_files:
            logger.warning(f"文件夹内未找到软著文件: {folder_path}")
            return self._get_empty_result()

        # 筛选匹配游戏名的文件
        matched_files = []
        if game_name_hint:
            hint_clean = game_name_hint.replace(' ', '').replace('_', '')
            print(f"  游戏名清理后: '{hint_clean}'")
            
            for f in all_files:
                fname = os.path.basename(f)
                fname_clean = fname.replace(' ', '').replace('_', '')
                if hint_clean in fname_clean:
                    matched_files.append(f)
                    print(f"  ✅ 文件名匹配: {fname}")
            
            # 从 OCR 缓存中查找
            if not matched_files and cached_ocr_texts:
                print(f"  🔍 尝试从 OCR 缓存匹配...")
                for img_path in cached_ocr_texts.keys():
                    img_dir = os.path.dirname(img_path).replace('\\', '/')
                    if img_dir == folder_path_norm:
                        matched_files.append(img_path)
                        print(f"  ✅ OCR缓存匹配: {os.path.basename(img_path)}")
            
            if not matched_files:
                logger.warning(f"文件夹 {folder_path} 未找到游戏 '{game_name_hint}' 对应的文件")
                return self._get_empty_result()
        else:
            matched_files = all_files

        print(f"  匹配到 {len(matched_files)} 个文件")

        # 处理 PDF 文件（直接提取文字，不需要 OCR）
        pdf_files = [f for f in matched_files if os.path.splitext(f)[1].lower() == '.pdf']
        print(f"  PDF 文件: {len(pdf_files)} 个")
        
        for fpath in pdf_files:
            try:
                text = self._parse_pdf(fpath)
                if text and text.strip():
                    all_text_parts.append(text)
                    logger.info(f"  PDF 文字提取成功: {len(text)} 字符")
            except Exception as e:
                logger.warning(f"  PDF 处理失败: {e}")

        # 处理图片文件 - 检查缓存命中情况
        image_files = [f for f in matched_files if f not in pdf_files]
        print(f"  图片文件: {len(image_files)} 个")
        
        cache_hit_count = 0
        cache_miss_count = 0
        
        for fpath in image_files:
            fpath_norm = os.path.normpath(fpath).replace('\\', '/')
            text = None
            
            # 优先从传入的缓存获取
            if cached_ocr_texts:
                text = cached_ocr_texts.get(fpath_norm) or cached_ocr_texts.get(fpath)
                if text:
                    cache_hit_count += 1
                    print(f"  ✅ 缓存命中: {os.path.basename(fpath)} ({len(text)} 字符)")
                else:
                    print(f"  ❌ 缓存未命中: {os.path.basename(fpath)}")
                    print(f"      标准化路径: {fpath_norm}")
                    print(f"      缓存键列表: {list(cached_ocr_texts.keys())[:5]}...")  # 只显示前5个
            
            # 如果传入缓存没有，尝试从 self.cached_ocr_texts 获取
            if not text and self.cached_ocr_texts:
                text = self.cached_ocr_texts.get(fpath_norm) or self.cached_ocr_texts.get(fpath)
                if text:
                    cache_hit_count += 1
                    print(f"  ✅ self缓存命中: {os.path.basename(fpath)}")
            
            # 缓存未命中，调用 API
            if not text:
                cache_miss_count += 1
                print(f"  🔴 缓存未命中，将调用 API OCR: {os.path.basename(fpath)}")
                try:
                    text = self._parse_image_with_api(fpath)
                    if text and text.strip():
                        # 同时更新两个缓存
                        if cached_ocr_texts is not None:
                            cached_ocr_texts[fpath_norm] = text
                        if self.cached_ocr_texts is not None:
                            self.cached_ocr_texts[fpath_norm] = text
                        print(f"  ✅ API OCR 成功: {len(text)} 字符")
                except Exception as e:
                    logger.error(f"  API OCR 失败: {e}")
                    text = ""
            
            if text and text.strip():
                all_text_parts.append(text)

        print(f"\n  📊 缓存统计: 命中={cache_hit_count}, 未命中={cache_miss_count}, 总计={len(image_files)}")

        if not all_text_parts:
            logger.warning("文件夹内所有文件均未能提取到文字")
            return self._get_empty_result()

        combined_text = '\n'.join(all_text_parts)
        print(f"  合并文本长度: {len(combined_text)} 字符")
        logger.info(f"合计提取 {len(combined_text)} 字符")

        result = self._extract_soft_info(combined_text)
        result['original_text'] = combined_text

        if not result.get('software_name') and game_name_hint:
            result['software_name'] = game_name_hint

        print(f"{'='*60}\n")
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