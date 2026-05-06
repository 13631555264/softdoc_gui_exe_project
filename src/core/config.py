# core/config.py
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置管理模块
"""

import os
import json
from pathlib import Path
from typing import Any, Dict, Optional
import logging

# 设置日志
logger = logging.getLogger("softdoc_generator")

class Config:
    """配置管理类"""
    
    def __init__(self, config_file: Optional[str] = None):
        self.config_file = config_file or self.get_default_config_path()
        self.config: Dict[str, Any] = self.load_config()
    
    @staticmethod
    def get_default_config_path() -> str:
        """获取默认配置文件路径"""
        config_dir = Path.home() / ".softdoc_generator"
        config_dir.mkdir(parents=True, exist_ok=True)
        return str(config_dir / "config.json")
    
    def get_default_config(self) -> Dict[str, Any]:
        """获取默认配置"""
        return {
            'version': '1.0.0',
            'paths': {
                'last_softdoc_path': '',
                'last_qg_path': '',
                'last_template_dir': '',
                'last_output_dir': '',
            },
            'processing': {
                'qg_file_encoding': 'GBK',
                'authorization_years': 10,
                'ocr_language': 'chi_sim+eng',
                'default_template_zip': '',
            },
            'gui': {
                'window_width': 800,
                'window_height': 1000,
                'theme': 'default',
                'font_size': 12,
                'auto_save_config': True,
                'show_log_panel': True,
            },
            'advanced': {
                'tesseract_path': '',
                'python_path': '',
                'debug_mode': False,
                'log_level': 'INFO',
                'temp_dir': '',
            }
        }
    
    def load_config(self) -> Dict[str, Any]:
        """加载配置"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded_config = json.load(f)
                
                # 合并默认配置和加载的配置
                default_config = self.get_default_config()
                merged_config = self._merge_configs(default_config, loaded_config)
                
                logger.info(f"配置加载成功: {self.config_file}")
                return merged_config
                
            except Exception as e:
                logger.error(f"加载配置文件失败: {e}")
                return self.get_default_config()
        else:
            logger.info("配置文件不存在，使用默认配置")
            return self.get_default_config()
    
    def _merge_configs(self, default: Dict, loaded: Dict) -> Dict:
        """深度合并两个配置字典"""
        merged = default.copy()
        
        for key, value in loaded.items():
            if key in merged and isinstance(merged[key], dict) and isinstance(value, dict):
                merged[key] = self._merge_configs(merged[key], value)
            else:
                merged[key] = value
        
        return merged
    
    def save_config(self) -> bool:
        """保存配置"""
        try:
            config_dir = os.path.dirname(self.config_file)
            if config_dir:
                os.makedirs(config_dir, exist_ok=True)
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            
            logger.info(f"配置保存成功: {self.config_file}")
            return True
            
        except Exception as e:
            logger.error(f"保存配置文件失败: {e}")
            return False
    
    def get(self, key: str, default: Any = None) -> Any:
        """获取配置值
        
        Args:
            key: 配置键，支持点号分隔，如 'paths.last_softdoc_path'
            default: 默认值
            
        Returns:
            Any: 配置值
        """
        try:
            keys = key.split('.')
            value = self.config
            
            for k in keys:
                value = value[k]
            
            return value
        except (KeyError, TypeError):
            return default
    
    def set(self, key: str, value: Any) -> bool:
        """设置配置值
        
        Args:
            key: 配置键，支持点号分隔
            value: 配置值
            
        Returns:
            bool: 是否成功
        """
        try:
            keys = key.split('.')
            config_part = self.config
            
            # 遍历到倒数第二个键
            for k in keys[:-1]:
                if k not in config_part:
                    config_part[k] = {}
                config_part = config_part[k]
            
            # 设置最后一个键的值
            config_part[keys[-1]] = value
            
            # 自动保存
            if self.get('gui.auto_save_config', True):
                self.save_config()
            
            return True
            
        except Exception as e:
            logger.error(f"设置配置失败: {key} = {value}, 错误: {e}")
            return False
    
    def update(self, updates: Dict[str, Any]) -> bool:
        """批量更新配置
        
        Args:
            updates: 更新字典
            
        Returns:
            bool: 是否成功
        """
        try:
            for key, value in updates.items():
                self.set(key, value)
            return True
        except Exception as e:
            logger.error(f"批量更新配置失败: {e}")
            return False
    
    def reset_to_default(self) -> bool:
        """重置为默认配置"""
        self.config = self.get_default_config()
        return self.save_config()
    
    def get_all(self) -> Dict[str, Any]:
        """获取所有配置"""
        return self.config.copy()
    
    def get_file_paths(self) -> Dict[str, str]:
        """获取文件路径配置"""
        return self.get('paths', {})
    
    def set_last_path(self, path_type: str, path: str) -> bool:
        """设置最后使用的路径
        
        Args:
            path_type: 路径类型 ('softdoc', 'qg', 'template', 'output')
            path: 路径值
            
        Returns:
            bool: 是否成功
        """
        path_key = f'paths.last_{path_type}_path'
        return self.set(path_key, path)
    
    def get_last_path(self, path_type: str) -> str:
        """获取最后使用的路径
        
        Args:
            path_type: 路径类型
            
        Returns:
            str: 路径值
        """
        path_key = f'paths.last_{path_type}_path'
        return self.get(path_key, '')
    
    def set_processing_config(self, key: str, value: Any) -> bool:
        """设置处理配置
        
        Args:
            key: 配置键
            value: 配置值
            
        Returns:
            bool: 是否成功
        """
        config_key = f'processing.{key}'
        return self.set(config_key, value)
    
    def get_processing_config(self, key: str, default: Any = None) -> Any:
        """获取处理配置
        
        Args:
            key: 配置键
            default: 默认值
            
        Returns:
            Any: 配置值
        """
        config_key = f'processing.{key}'
        return self.get(config_key, default)
    
    def set_gui_config(self, key: str, value: Any) -> bool:
        """设置GUI配置
        
        Args:
            key: 配置键
            value: 配置值
            
        Returns:
            bool: 是否成功
        """
        config_key = f'gui.{key}'
        return self.set(config_key, value)
    
    def get_gui_config(self, key: str, default: Any = None) -> Any:
        """获取GUI配置
        
        Args:
            key: 配置键
            default: 默认值
            
        Returns:
            Any: 配置值
        """
        config_key = f'gui.{key}'
        return self.get(config_key, default)
    
    def set_advanced_config(self, key: str, value: Any) -> bool:
        """设置高级配置
        
        Args:
            key: 配置键
            value: 配置值
            
        Returns:
            bool: 是否成功
        """
        config_key = f'advanced.{key}'
        return self.set(config_key, value)
    
    def get_advanced_config(self, key: str, default: Any = None) -> Any:
        """获取高级配置
        
        Args:
            key: 配置键
            default: 默认值
            
        Returns:
            Any: 配置值
        """
        config_key = f'advanced.{key}'
        return self.get(config_key, default)
    
    def validate(self) -> Dict[str, str]:
        """验证配置
        
        Returns:
            Dict[str, str]: 验证错误信息
        """
        errors = {}
        
        # 检查必要的路径配置
        required_paths = ['last_softdoc_path', 'last_qg_path', 'last_template_dir', 'last_output_dir']
        for path_key in required_paths:
            path = self.get(f'paths.{path_key}', '')
            if path and not os.path.exists(path):
                errors[path_key] = f"路径不存在: {path}"
        
        # 检查处理配置
        auth_years = self.get('processing.authorization_years', 10)
        if not isinstance(auth_years, int) or auth_years <= 0:
            errors['authorization_years'] = f"授权年限必须为正整数: {auth_years}"
        
        ocr_lang = self.get('processing.ocr_language', '')
        if not ocr_lang:
            errors['ocr_language'] = "OCR语言不能为空"
        
        # 检查GUI配置
        window_width = self.get('gui.window_width', 800)
        window_height = self.get('gui.window_height', 600)
        if window_width < 400 or window_height < 300:
            errors['window_size'] = f"窗口大小过小: {window_width}x{window_height}"
        
        return errors

def main():
    """测试函数"""
    print("配置管理模块测试")
    print("-" * 40)
    
    # 创建配置实例
    config = Config()
    
    print(f"配置文件路径: {config.config_file}")
    print(f"配置版本: {config.get('version')}")
    print()
    
    # 测试设置和获取
    print("测试设置和获取:")
    config.set('paths.last_softdoc_path', '/path/to/softdoc.pdf')
    softdoc_path = config.get('paths.last_softdoc_path')
    print(f"  软著路径: {softdoc_path}")
    
    config.set('processing.authorization_years', 15)
    auth_years = config.get('processing.authorization_years')
    print(f"  授权年限: {auth_years}")
    
    # 测试批量更新
    print("\n测试批量更新:")
    updates = {
        'gui.window_width': 1024,
        'gui.window_height': 768,
        'processing.qg_file_encoding': 'UTF-8'
    }
    config.update(updates)
    
    print(f"  窗口大小: {config.get('gui.window_width')}x{config.get('gui.window_height')}")
    print(f"  文件编码: {config.get('processing.qg_file_encoding')}")
    
    # 测试验证
    print("\n测试配置验证:")
    errors = config.validate()
    if errors:
        print("  验证错误:")
        for key, error in errors.items():
            print(f"    {key}: {error}")
    else:
        print("  配置验证通过")
    
    # 保存配置
    if config.save_config():
        print("\n✅ 配置保存成功")
    else:
        print("\n❌ 配置保存失败")
    
    print("\n✅ 配置管理模块测试完成")

if __name__ == "__main__":
    main()