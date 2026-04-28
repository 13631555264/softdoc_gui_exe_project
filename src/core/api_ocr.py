# core/api_ocr.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
火山引擎 API OCR 识别模块
使用豆包大模型识别图片/PDF中的文字
"""

import os
import base64
import requests
import logging
from pathlib import Path
from typing import Dict, Optional, List
import tempfile
import fitz  # PyMuPDF，用于PDF转图片

logger = logging.getLogger("softdoc_generator")

class VolcEngineOCR:
    """火山引擎 OCR 识别器"""
    
    def __init__(self, api_key: str = None, model: str = "doubao-seed-2-0-lite-260215"):
        self.api_url = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"
        # 请在这里填入您的 API Key
        self.api_key = api_key or "374bc8a8-b92c-4e1c-a839-6f6d51f61b8c"
        self.model = model
        self.max_tokens = 65535
    
    def read_image_as_base64(self, image_path: str) -> str:
        """将图片转换为 base64 格式"""
        image_path = os.path.normpath(image_path)
        
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"图片文件不存在: {image_path}")
        
        # 获取 MIME 类型
        ext = os.path.splitext(image_path)[1].lower()
        mime_types = {
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.png': 'image/png',
            '.gif': 'image/gif',
            '.bmp': 'image/bmp',
            '.webp': 'image/webp'
        }
        mime_type = mime_types.get(ext, 'image/jpeg')
        
        # 读取并编码
        with open(image_path, 'rb') as f:
            base64_data = base64.b64encode(f.read()).decode('utf-8')
        
        return f"data:{mime_type};base64,{base64_data}"
    
    def pdf_to_images(self, pdf_path: str, dpi: int = 150) -> List[str]:
        """将 PDF 转换为图片列表"""
        image_paths = []
        
        try:
            doc = fitz.open(pdf_path)
            zoom = dpi / 72  # 1点=1/72英寸
            matrix = fitz.Matrix(zoom, zoom)
            
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                pix = page.get_pixmap(matrix=matrix)
                
                # 保存为临时图片
                temp_img = tempfile.NamedTemporaryFile(suffix=f'.png', delete=False)
                pix.save(temp_img.name)
                image_paths.append(temp_img.name)
                temp_img.close()
            
            doc.close()
            logger.info(f"PDF 转换完成，共 {len(image_paths)} 页")
            
        except Exception as e:
            logger.error(f"PDF 转换失败: {e}")
            raise
        
        return image_paths
    
    def recognize_image(self, image_path: str, question: str = None) -> str:
        """识别单张图片中的文字"""
        # 默认 prompt：针对软著证书优化，提高识别准确性
        if question is None:
            question = """请仔细识别这张图片中的所有文字内容，这是一张中国计算机软件著作权登记证书。

重要要求：
1. **精确识别**：严格按照图片中的文字输出，不要修改、添加或推断任何文字
2. **注意相似字**：常见易混淆的字包括：要/耍、来/米、大/太、人/入等，务必对照图片仔细识别
3. **保持原样**：保留原文中的标点符号、换行和格式
4. **逐字核对**：对于公司名称、人名等重要信息，逐字核对确保准确
5. **特殊符号**：注意中文顿号（、）和逗号（，）的区别，注意书名号《》和引号""的区别

请直接输出识别结果，不要添加解释说明。"""
        """识别单张图片中的文字"""
        try:
            image_data_url = self.read_image_as_base64(image_path)
            
            payload = {
                "model": self.model,
                "max_completion_tokens": self.max_tokens,
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": image_data_url
                                }
                            },
                            {
                                "type": "text",
                                "text": question
                            }
                        ]
                    }
                ],
                "reasoning_effort": "medium"
            }
            
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}"
            }
            
            # 重试机制：最多尝试 3 次，超时时间增加到 180 秒
            last_error = None
            for attempt in range(1, 4):
                try:
                    response = requests.post(
                        self.api_url,
                        headers=headers,
                        json=payload,
                        timeout=180
                    )
                    
                    if response.status_code == 200:
                        result = response.json()
                        content = result['choices'][0]['message']['content']
                        logger.info(f"API 识别成功，返回 {len(content)} 字符")
                        return content
                    else:
                        logger.warning(f"API 请求失败 (尝试 {attempt}/3): {response.status_code}")
                        last_error = f"HTTP {response.status_code}"
                        
                except requests.exceptions.Timeout:
                    logger.warning(f"API 请求超时 (尝试 {attempt}/3): 180秒内未收到响应")
                    last_error = "请求超时"
                except Exception as e:
                    logger.warning(f"API 请求异常 (尝试 {attempt}/3): {e}")
                    last_error = str(e)
                
                if attempt < 3:
                    import time
                    time.sleep(2)  # 等待 2 秒后重试
            
            logger.error(f"图片识别失败: {last_error}")
            return ""
        except Exception as e:
            logger.error(f"图片识别失败: {e}")
            return ""
    
    def recognize_pdf(self, pdf_path: str) -> str:
        """识别 PDF 文件中的所有文字"""
        logger.info(f"开始识别 PDF: {pdf_path}")
        
        # 将 PDF 转换为图片
        image_paths = self.pdf_to_images(pdf_path)
        
        all_text = []
        for i, img_path in enumerate(image_paths, 1):
            logger.info(f"识别第 {i}/{len(image_paths)} 页...")
            text = self.recognize_image(img_path, f"请提取第{i}页的所有文字内容。")
            all_text.append(f"=== 第 {i} 页 ===\n{text}")
            
            # 删除临时图片
            try:
                os.unlink(img_path)
            except:
                pass
        
        return '\n\n'.join(all_text)
    
    def recognize_file(self, file_path: str) -> str:
        """识别文件（自动判断类型）"""
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext == '.pdf':
            return self.recognize_pdf(file_path)
        elif ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp', '.tiff']:
            return self.recognize_image(file_path)
        else:
            raise ValueError(f"不支持的文件类型: {ext}")


def extract_soft_info_from_text(text: str) -> Dict:
    """从识别文本中提取软著信息"""
    import re
    
    info = {
        'software_name': '',
        'version': '',
        'copyright_holder': '',
        'software_type': '',
        'registration_number': '',
        'completion_date': '',
        'publish_date': '',
        'original_text': text
    }
    
    # 提取软件名称
    patterns = [
        r'软件名称[：:]\s*([^\n]+)',
        r'软件全称[：:]\s*([^\n]+)',
        r'产品名称[：:]\s*([^\n]+)',
        r'名称[：:]\s*([^\n]+)',
        r'《([^》]+)》',
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            info['software_name'] = match.group(1).strip()
            break
    
    # 如果没找到，尝试取第一行
    if not info['software_name']:
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        if lines:
            info['software_name'] = lines[0][:100]
    
    # 提取版本号
    patterns = [
        r'版本号[：:]\s*([Vv]?\d+\.\d+(?:\.\d+)?)',
        r'版本[：:]\s*([Vv]?\d+\.\d+(?:\.\d+)?)',
        r'[Vv](\d+\.\d+(?:\.\d+)?)',
        r'(\d+\.\d+\.\d+)',
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            info['version'] = match.group(1).strip()
            break
    
    # 提取著作权人
    patterns = [
        r'著作权人[：:]\s*([^\n]+)',
        r'著作权人[：:]\s*([^\n]+?(?:有限公司|股份公司|工作室|中心))',
        r'申请人[：:]\s*([^\n]+)',
        r'权利人[：:]\s*([^\n]+)',
        r'版权人[：:]\s*([^\n]+)',
        r'([^\n]{2,30}(?:有限公司|股份公司|科技公司|网络科技))',
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            holder = match.group(1).strip()
            if len(holder) < 100:
                info['copyright_holder'] = holder
                break
    
    # 提取登记号
    patterns = [
        r'登记号[：:]\s*([^\n]+)',
        r'软著登字第\s*(\d+)',
        r'([A-Z0-9]{10,20})',
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            info['registration_number'] = match.group(1).strip()
            break
    
    # 提取日期
    patterns = [
        r'完成日期[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)',
        r'开发完成[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)',
        r'(\d{4}-\d{1,2}-\d{1,2})',
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            info['completion_date'] = match.group(1).strip()
            break
    
    return info


def main():
    """测试函数"""
    import sys
    from pathlib import Path
    
    # 配置日志
    logging.basicConfig(level=logging.INFO)
    
    # 创建 OCR 实例
    ocr = VolcEngineOCR()
    
    # 测试文件路径（请修改为您的软著文件路径）
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        # 在项目目录中查找
        current_dir = Path(__file__).parent.parent
        files = list(current_dir.glob("*.pdf")) + list(current_dir.glob("*.jpg")) + list(current_dir.glob("*.png"))
        if files:
            file_path = str(files[0])
            print(f"自动选择文件: {file_path}")
        else:
            print("请提供文件路径: python api_ocr.py <文件路径>")
            return
    
    # 识别文件
    print(f"\n正在识别: {file_path}")
    text = ocr.recognize_file(file_path)
    
    print("\n" + "=" * 60)
    print("识别结果:")
    print("=" * 60)
    print(text)
    
    # 提取信息
    info = extract_soft_info_from_text(text)
    
    print("\n" + "=" * 60)
    print("提取的软著信息:")
    print("=" * 60)
    print(f"软件名称: {info['software_name']}")
    print(f"版本号: {info['version']}")
    print(f"著作权人: {info['copyright_holder']}")
    print(f"登记号: {info['registration_number']}")
    print(f"完成日期: {info['completion_date']}")


if __name__ == "__main__":
    main()