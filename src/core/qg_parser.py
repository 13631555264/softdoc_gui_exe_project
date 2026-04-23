# core/qg_parser.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
渠广文件解析器
功能：解析渠广txt文件，提取游戏信息（GBK编码）
"""

import os
import re
import logging
from typing import Dict, Optional
from .config import Config

# 设置日志
logger = logging.getLogger("softdoc_generator")

class QGParser:
    """渠广文件解析器"""
    
    def __init__(self, config: Optional[Config] = None):
        self.config = config or Config()
        self.encoding = self.config.get('processing.qg_file_encoding', 'GBK')
    
    def parse_file(self, file_path: str) -> Dict:
        """解析渠广文件
        
        Args:
            file_path: 渠广文件路径
            
        Returns:
            Dict: 解析结果，包含游戏名称、包名、上架主体等信息
        """
        if not os.path.exists(file_path):
            logger.error(f"渠广文件不存在: {file_path}")
            raise FileNotFoundError(f"渠广文件不存在: {file_path}")
        
        try:
            with open(file_path, 'r', encoding=self.encoding) as f:
                content = f.read()
        except UnicodeDecodeError:
            # 尝试其他编码
            encodings = ['GBK', 'GB2312', 'GB18030', 'UTF-8', 'UTF-16']
            for encoding in encodings:
                try:
                    with open(file_path, 'r', encoding=encoding) as f:
                        content = f.read()
                    self.encoding = encoding
                    logger.info(f"使用编码: {encoding}")
                    break
                except UnicodeDecodeError:
                    continue
            else:
                logger.error(f"无法解码渠广文件，尝试过的编码: {encodings}")
                raise
        
        return self.parse_content(content)
    
    def parse_content(self, content: str) -> Dict:
        """解析渠广内容
        
        Args:
            content: 渠广文件内容
            
        Returns:
            Dict: 解析结果
        """
        result = {
            'game_name': '',
            'package_name': '',
            'publisher': '',
            'developer': '',
            'version': '',
            'category': '',
            'description': '',
            'raw_content': content
        }
        
        lines = content.strip().split('\n')
        
        # 解析游戏名称（通常是第一行）
        if lines:
            result['game_name'] = lines[0].strip()
        
        # 解析包名
        for line in lines:
            if '包名' in line or 'package' in line.lower():
                parts = line.split('：') if '：' in line else line.split(':')
                if len(parts) > 1:
                    result['package_name'] = parts[1].strip()
                    break
        
        # 解析上架主体（公司名称）
        company_keywords = ['公司', '主体', '上架', '出版单位', '运营单位']
        for line in lines:
            if any(keyword in line for keyword in company_keywords):
                # 提取公司名称
                for keyword in company_keywords:
                    if keyword in line:
                        parts = line.split(keyword)
                        if len(parts) > 1:
                            company_part = parts[1].strip('：: ')
                            if company_part:
                                result['publisher'] = company_part
                                break
                if result['publisher']:
                    break
        
        # 如果没有找到明确的公司名称，尝试查找包含"公司"的行
        if not result['publisher']:
            for line in lines:
                if '公司' in line and '：' not in line and ':' not in line:
                    result['publisher'] = line.strip()
                    break
        
        # 解析开发者
        dev_keywords = ['开发', '研发', 'developer', '制作']
        for line in lines:
            if any(keyword in line for keyword in dev_keywords):
                for keyword in dev_keywords:
                    if keyword in line:
                        parts = line.split(keyword)
                        if len(parts) > 1:
                            dev_part = parts[1].strip('：: ')
                            if dev_part:
                                result['developer'] = dev_part
                                break
                if result['developer']:
                    break
        
        # 如果没有找到开发者，使用上架主体作为开发者
        if not result['developer'] and result['publisher']:
            result['developer'] = result['publisher']
        
        # 解析版本号
        version_patterns = [
            r'版本[：:]?\s*([0-9]+\.[0-9]+\.[0-9]+)',
            r'version[：:]?\s*([0-9]+\.[0-9]+\.[0-9]+)',
            r'v([0-9]+\.[0-9]+\.[0-9]+)',
        ]
        
        for pattern in version_patterns:
            match = re.search(pattern, content, re.IGNORECASE)
            if match:
                result['version'] = match.group(1)
                break
        
        # 解析分类
        category_keywords = ['分类', '类别', 'category', 'type']
        for line in lines:
            if any(keyword in line.lower() for keyword in category_keywords):
                for keyword in category_keywords:
                    if keyword in line.lower():
                        parts = line.lower().split(keyword)
                        if len(parts) > 1:
                            category_part = parts[1].strip('：: ')
                            if category_part:
                                result['category'] = category_part
                                break
                if result['category']:
                    break
        
        # 解析描述（通常是最后几行）
        description_lines = []
        for line in reversed(lines):
            line = line.strip()
            if line and not any(keyword in line for keyword in [
                '包名', '版本', '分类', '公司', '主体', '开发', '研发'
            ]):
                description_lines.insert(0, line)
            elif description_lines:
                break
        
        if description_lines:
            result['description'] = '\n'.join(description_lines)
        
        # 记录解析结果
        logger.info(f"渠广文件解析完成:")
        logger.info(f"  游戏名称: {result['game_name']}")
        logger.info(f"  包名: {result['package_name']}")
        logger.info(f"  上架主体: {result['publisher']}")
        logger.info(f"  开发者: {result['developer']}")
        logger.info(f"  版本: {result['version']}")
        logger.info(f"  分类: {result['category']}")
        
        return result
    
    def validate_result(self, result: Dict) -> bool:
        """验证解析结果
        
        Args:
            result: 解析结果
            
        Returns:
            bool: 是否有效
        """
        required_fields = ['game_name', 'package_name', 'publisher']
        
        for field in required_fields:
            if not result.get(field):
                logger.warning(f"渠广文件缺少必要字段: {field}")
                return False
        
        # 验证包名格式
        package_name = result.get('package_name', '')
        if package_name and not re.match(r'^[a-z][a-z0-9_]*(\.[a-z][a-z0-9_]*)+$', package_name):
            logger.warning(f"包名格式可能不正确: {package_name}")
        
        return True

def main():
    """测试函数"""
    # 创建测试实例
    parser = QGParser()
    
    # 测试数据
    test_content = """生存挑战模拟
包名：com.hxwl.sctzmn.vivominigame
版本：1.0.0
分类：休闲益智
深圳市鸿鑫网络科技有限公司
这是一款生存挑战类游戏，玩家需要在各种环境下生存"""
    
    try:
        result = parser.parse_content(test_content)
        
        print("渠广文件解析测试:")
        print(f"游戏名称: {result['game_name']}")
        print(f"包名: {result['package_name']}")
        print(f"上架主体: {result['publisher']}")
        print(f"开发者: {result['developer']}")
        print(f"版本: {result['version']}")
        print(f"分类: {result['category']}")
        print(f"描述: {result['description']}")
        
        if parser.validate_result(result):
            print("✅ 解析结果有效")
        else:
            print("❌ 解析结果无效")
            
    except Exception as e:
        print(f"解析失败: {e}")

if __name__ == "__main__":
    main()