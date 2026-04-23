# core/softdoc_parser.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
软著文件解析器
功能：使用火山引擎 API 解析软著 PDF 或图片文件
"""

import os
import re
import logging
from typing import Dict, Optional
from .config import Config
from .api_ocr import VolcEngineOCR, extract_soft_info_from_text

logger = logging.getLogger("softdoc_generator")

class SoftDocParser:
    """软著文件解析器 - 使用火山引擎 API"""
    
    def __init__(self, config: Optional[Config] = None):
        self.config = config or Config()
        # 初始化 API OCR
        api_key = self.config.get('advanced.volc_api_key', 'a0081937-958a-44d4-8144-f713d09ada03')
        self.ocr = VolcEngineOCR(api_key=api_key)
    
    def parse_file(self, file_path: str) -> Dict:
        """解析软著文件
        
        Args:
            file_path: 软著文件路径（PDF或图片）
            
        Returns:
            Dict: 解析结果
        """
        if not os.path.exists(file_path):
            logger.error(f"软著文件不存在: {file_path}")
            raise FileNotFoundError(f"软著文件不存在: {file_path}")
        
        logger.info(f"正在使用火山引擎 API 识别: {file_path}")
        
        try:
            # 调用 API 识别
            text = self.ocr.recognize_file(file_path)
            
            if not text:
                logger.warning("API 识别返回空内容")
                return self._get_empty_result()
            
            logger.info(f"API 识别成功，获取 {len(text)} 字符")
            
            # 提取信息
            result = extract_soft_info_from_text(text)
            result['original_text'] = text
            
            # 记录解析结果
            logger.info(f"软著文件解析完成:")
            logger.info(f"  软件名称: {result['software_name']}")
            logger.info(f"  版本号: {result['version']}")
            logger.info(f"  著作权人: {result['copyright_holder']}")
            logger.info(f"  登记号: {result['registration_number']}")
            
            return result
            
        except Exception as e:
            logger.error(f"API 识别失败: {e}")
            return self._get_empty_result()
    
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
    
    def validate_result(self, result: Dict) -> bool:
        """验证解析结果"""
        if result.get('software_name') and result.get('copyright_holder'):
            return True
        logger.warning(f"软著信息不完整: 软件名={result.get('software_name')}, 著作权人={result.get('copyright_holder')}")
        return False


def main():
    """测试函数"""
    import sys
    from pathlib import Path
    
    logging.basicConfig(level=logging.INFO)
    
    parser = SoftDocParser()
    
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        current_dir = Path(__file__).parent.parent
        files = list(current_dir.glob("*.pdf")) + list(current_dir.glob("*.jpg")) + list(current_dir.glob("*.png"))
        if files:
            file_path = str(files[0])
        else:
            print("请提供文件路径: python softdoc_parser.py <文件路径>")
            return
    
    result = parser.parse_file(file_path)
    
    print("\n解析结果:")
    print(f"  软件名称: {result['software_name']}")
    print(f"  版本号: {result['version']}")
    print(f"  著作权人: {result['copyright_holder']}")
    print(f"  登记号: {result['registration_number']}")
    print(f"  完成日期: {result['completion_date']}")


if __name__ == "__main__":
    main()