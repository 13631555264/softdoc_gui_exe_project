#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
软著文档生成工具 - 主程序入口
"""

import os
import sys
from pathlib import Path

# 添加当前目录到Python路径
current_dir = Path(__file__).parent.parent
sys.path.insert(0, str(current_dir / "src"))

def check_environment():
    """检查运行环境"""
    print("🔍 检查运行环境...")
    
    # 检查Python版本
    if sys.version_info < (3, 7):
        print("❌ 需要 Python 3.7 或更高版本")
        return False
    
    print(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")
    
    # 检查tkinter
    try:
        import tkinter
        print("✅ tkinter 可用")
    except ImportError:
        print("❌ tkinter 不可用")
        return False
    
    # 检查核心模块
    try:
        import core
        print("✅ 核心模块可用")
    except ImportError as e:
        print(f"❌ 核心模块导入失败: {e}")
        return False
    
    # 检查GUI模块
    try:
        import gui
        print("✅ GUI模块可用")
    except ImportError as e:
        print(f"❌ GUI模块导入失败: {e}")
        return False
    
    return True

def main():
    """主函数"""
    print("=" * 70)
    print("软著文档生成工具")
    print("=" * 70)
    print()
    
    # 检查环境
    if not check_environment():
        print("\n❌ 环境检查失败，请确保:")
        print("   1. Python版本为3.7+")
        print("   2. tkinter已安装")
        print("   3. 所有依赖包已安装")
        input("\n按回车键退出...")
        return 1
    
    print()
    print("🚀 启动GUI界面...")
    print("-" * 30)
    
    try:
        # 导入并运行GUI主程序
        from gui.main_window import MainWindow
        
        # 创建主窗口
        app = MainWindow()
        
        # 启动主循环
        app.mainloop()
        
        print("\n✅ 程序正常退出")
        return 0
        
    except ImportError as e:
        print(f"❌ 导入GUI模块失败: {e}")
        print("\n请运行以下命令安装依赖:")
        print("    pip install -r requirements.txt")
        input("\n按回车键退出...")
        return 1
        
    except Exception as e:
        print(f"❌ 启动过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        input("\n按回车键退出...")
        return 1

if __name__ == "__main__":
    sys.exit(main())