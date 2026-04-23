# core/document_generator.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档生成器
功能：复制模板文档并替换其中的文字内容，保持原格式
"""

import os
import re
import shutil
import logging
from datetime import datetime, timedelta
from typing import Dict, List, Optional
from docx import Document
from .config import Config

# 设置日志
logger = logging.getLogger("softdoc_generator")

class DocumentGenerator:
    """文档生成器 - 基于模板复制和文字替换"""
    
    def __init__(self, config: Optional[Config] = None):
        self.config = config or Config()
        self.template_dir = ""
        self.output_dir = ""
    
    def set_template_dir(self, template_dir: str):
        """设置模板目录"""
        self.template_dir = template_dir
        logger.info(f"设置模板目录: {template_dir}")
    
    def set_output_dir(self, output_dir: str):
        """设置输出目录"""
        self.output_dir = output_dir
        os.makedirs(output_dir, exist_ok=True)
        logger.info(f"设置输出目录: {output_dir}")
    
    def generate_documents(self, game_info: Dict, soft_info: Dict) -> List[str]:
        """生成所有文档"""
        generated_files = []
        
        # 创建游戏专属目录
        game_name = self._sanitize_filename(game_info.get('game_name', '未知游戏'))
        game_output_dir = os.path.join(self.output_dir, game_name)
        os.makedirs(game_output_dir, exist_ok=True)
        
        # 准备替换数据
        replace_data = self._prepare_replace_data(game_info, soft_info)
        
        # 模板文件列表
        templates = [
            ("单机承诺函.docx", "单机承诺函"),
            ("免责承诺函.docx", "免责承诺函"),
        ]
        
        # 生成基础文档
        for template_name, doc_type in templates:
            output_file = self._copy_and_replace_template(
                template_name, doc_type, game_output_dir, replace_data, game_info
            )
            if output_file:
                generated_files.append(output_file)
        
        # 判断是否需要生成授权书
        if self._need_authorization(game_info, soft_info):
            auth_file = self._copy_and_replace_template(
                "授权书.docx", "授权书", game_output_dir, replace_data, game_info, soft_info
            )
            if auth_file:
                generated_files.append(auth_file)
        
        # 生成处理日志
        log_file = self._generate_process_log(game_info, soft_info, game_output_dir, generated_files)
        if log_file:
            generated_files.append(log_file)
        
        logger.info(f"文档生成完成，共生成{len(generated_files)}个文件")
        return generated_files
    
    def _prepare_replace_data(self, game_info: Dict, soft_info: Dict) -> Dict:
        """准备替换数据"""
        auth_years = self.config.get('processing.authorization_years', 10)
        now = datetime.now()
        
        # 计算授权结束日期
        auth_end_date = now + timedelta(days=auth_years * 365)
        
        # 平台名称（可从配置读取或默认）
        platform_name = self.config.get('processing.platform_name', '广东天宸网络科技有限公司')
        
        return {
            '{平台名称}': platform_name,
            '{游戏名称}': game_info.get('game_name', ''),
            '{包名}': game_info.get('package_name', ''),
            '{公司名称}': game_info.get('publisher', ''),
            '{开发者}': game_info.get('developer', ''),
            '{版本}': game_info.get('version', ''),
            '{软件名称}': soft_info.get('software_name', ''),
            '{版本号}': soft_info.get('version', ''),
            '{著作权人}': soft_info.get('copyright_holder', ''),
            '{软著登记号}': soft_info.get('registration_number', ''),
            '{授权方}': soft_info.get('copyright_holder', game_info.get('developer', '')),
            '{被授权方}': game_info.get('publisher', ''),
            '{授权年限}': str(auth_years),
            '{授权开始日期}': now.strftime('%Y年%m月%d日'),
            '{授权结束日期}': auth_end_date.strftime('%Y年%m月%d日'),
            '{当前日期}': now.strftime('%Y年%m月%d日'),
            '{当前年份}': now.strftime('%Y'),
            '{当前月份}': now.strftime('%m'),
            '{当前日}': now.strftime('%d'),
            '{版号}': self.config.get('processing.game_version_number', '0'),
        }
        
    def _copy_and_replace_template(self, template_name: str, doc_type: str, 
                                    output_dir: str, replace_data: Dict,
                                    game_info: Dict, soft_info: Dict = None) -> Optional[str]:
        """复制模板并替换文字"""
        template_path = os.path.join(self.template_dir, template_name)
        
        if not os.path.exists(template_path):
            logger.warning(f"模板不存在: {template_path}，跳过 {doc_type}")
            return None
        
        # 确定输出文件名
        game_name = game_info.get('game_name', '游戏')
        company = game_info.get('publisher', '公司')
        
        if doc_type == "单机承诺函":
            filename = f"{game_name}-单机-{company}.docx"
        elif doc_type == "免责承诺函":
            filename = f"{game_name}-免责-{company}.docx"
        elif doc_type == "授权书":
            licensor = (soft_info or {}).get('copyright_holder', '授权方')
            filename = f"{game_name}-授权书-{company}-{licensor}.docx"
        else:
            filename = f"{game_name}-{doc_type}-{company}.docx"
        
        filename = self._sanitize_filename(filename)
        output_path = os.path.join(output_dir, filename)
        
        # 调试：打印替换数据
        logger.info(f"准备替换的数据 ({doc_type}):")
        for key, value in replace_data.items():
            logger.info(f"  {key} -> {value}")
        
        try:
            # 复制模板文件
            shutil.copy2(template_path, output_path)
            logger.info(f"复制模板: {template_name}")
            
            # 打开并替换文字
            doc = Document(output_path)
            
            # 调试：打印原始文本
            for para in doc.paragraphs[:5]:  # 只打印前5段
                logger.info(f"模板段落原文: {para.text[:100]}")
            
            replace_count = self._replace_all_text(doc, replace_data)
            doc.save(output_path)
            
            logger.info(f"完成 {doc_type}: 替换了 {replace_count} 处，输出: {filename}")
            return output_path
            
        except Exception as e:
            logger.error(f"处理模板 {template_name} 失败: {e}")
            return None
    
    def _replace_all_text(self, doc, replace_data: Dict) -> int:
        """替换文档中的所有文本"""
        replace_count = 0
        
        # 替换段落
        for paragraph in doc.paragraphs:
            replace_count += self._replace_text_in_paragraph(paragraph, replace_data)
        
        # 替换表格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        replace_count += self._replace_text_in_paragraph(paragraph, replace_data)
        
        # 替换页眉页脚
        if hasattr(doc, 'sections'):
            for section in doc.sections:
                if section.header:
                    for paragraph in section.header.paragraphs:
                        replace_count += self._replace_text_in_paragraph(paragraph, replace_data)
                if section.footer:
                    for paragraph in section.footer.paragraphs:
                        replace_count += self._replace_text_in_paragraph(paragraph, replace_data)
        
        return replace_count
    
    def _replace_text_in_paragraph(self, paragraph, replace_data: Dict) -> int:
        """替换段落中的文本"""
        replace_count = 0
        
        for run in paragraph.runs:
            original_text = run.text
            new_text = original_text
            
            for old_str, new_str in replace_data.items():
                if old_str in new_text:
                    new_text = new_text.replace(old_str, new_str)
                    replace_count += original_text.count(old_str)
            
            if new_text != original_text:
                run.text = new_text
        
        return replace_count
    
    def _need_authorization(self, game_info: Dict, soft_info: Dict) -> bool:
        """判断是否需要生成授权书"""
        publisher = game_info.get('publisher', '')
        copyright_holder = soft_info.get('copyright_holder', '')
        
        # 如果上架主体和著作权人不同，需要授权书
        if publisher and copyright_holder and publisher != copyright_holder:
            logger.info(f"上架主体与著作权人不一致，需要生成授权书")
            return True
        
        return False
    
    def _generate_process_log(self, game_info: Dict, soft_info: Dict, output_dir: str, generated_files: List[str]) -> Optional[str]:
        """生成处理日志"""
        try:
            log_path = os.path.join(output_dir, "处理日志.txt")
            
            with open(log_path, 'w', encoding='utf-8') as f:
                f.write("=" * 60 + "\n")
                f.write("软著文档生成工具 - 处理日志\n")
                f.write("=" * 60 + "\n\n")
                
                f.write(f"处理时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                
                f.write("游戏信息:\n")
                f.write(f"  游戏名称: {game_info.get('game_name', '')}\n")
                f.write(f"  包名: {game_info.get('package_name', '')}\n")
                f.write(f"  上架主体: {game_info.get('publisher', '')}\n\n")
                
                f.write("软著信息:\n")
                f.write(f"  软件名称: {soft_info.get('software_name', '')}\n")
                f.write(f"  著作权人: {soft_info.get('copyright_holder', '')}\n")
                f.write(f"  登记号: {soft_info.get('registration_number', '')}\n\n")
                
                f.write("生成的文件:\n")
                for i, file_path in enumerate(generated_files, 1):
                    f.write(f"  {i}. {os.path.basename(file_path)}\n")
                
                f.write("\n" + "=" * 60 + "\n")
            
            return log_path
        except Exception as e:
            logger.error(f"生成日志失败: {e}")
            return None
    
    def _sanitize_filename(self, filename: str) -> str:
        """清理文件名"""
        illegal_chars = r'<>:"/\|?*'
        for char in illegal_chars:
            filename = filename.replace(char, '_')
        return filename.strip()


def main():
    """测试函数"""
    generator = DocumentGenerator()
    
    game_info = {
        'game_name': '生存挑战模拟',
        'package_name': 'com.hxwl.sctzmn.vivominigame',
        'publisher': '深圳市鸿鑫网络科技有限公司',
        'developer': '深圳市顺思畅想科技有限公司',
    }
    
    soft_info = {
        'software_name': '生存挑战模拟游戏软件',
        'copyright_holder': '深圳市顺思畅想科技有限公司',
        'registration_number': '2021SRE028650',
    }
    
    generator.set_template_dir("./templates")
    generator.set_output_dir("./test_output")
    files = generator.generate_documents(game_info, soft_info)
    
    print(f"生成了 {len(files)} 个文件")

if __name__ == "__main__":
    main()