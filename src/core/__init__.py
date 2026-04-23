# core/__init__.py
"""
软著文档生成工具 - 核心模块
"""

__version__ = "1.0.0"
__author__ = "软著文档生成工具团队"

from .qg_parser import QGParser
from .softdoc_parser import SoftDocParser
from .document_generator import DocumentGenerator
from .config import Config

__all__ = [
    'QGParser',
    'SoftDocParser', 
    'DocumentGenerator',
    'Config',
]