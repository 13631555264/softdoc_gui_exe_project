#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
带自动安装的运行脚本
功能：自动检查并安装所有依赖，然后运行主程序
"""

import os
import sys
import subprocess
from pathlib import Path

def print_banner():
    """打印横幅"""
    print("=" * 70)
    print("软著文档生成工具 - 自动化安装版")
    print("=" * 70)
    print()

def check_and_install_tesseract():
    """检查并安装Tesseract OCR"""
    print("🔍 检查 Tesseract OCR...")
    
    try:
        # 尝试运行tesseract命令
        result = subprocess.run(
            ['tesseract', '--version'],
            capture_output=True,
            text=True,
            timeout=10
        )
        
        if result.returncode == 0:
            version = result.stdout.split('\n')[0]
            print(f"✅ Tesseract OCR 已安装: {version}")
            return True
    except:
        pass
    
    print("❌ Tesseract OCR 未安装")
    print()
    print("需要安装 Tesseract OCR 才能使用OCR功能")
    print("请按照以下步骤操作:")
    print()
    print("方法1: 自动安装")
    print("   1. 下载 Tesseract 安装包")
    print("   2. 放入项目根目录的 installers/ 文件夹")
    print("   3. 运行 python src/tesseract_installer.py")
    print()
    print("方法2: 手动安装")
    print("   1. 访问: https://github.com/UB-Mannheim/tesseract/wiki")
    print("   2. 下载最新版安装程序")
    print("   3. 安装时勾选'添加到PATH'")
    print("   4. 安装中文语言包")
    print()
    
    print("是否继续运行？ (y/n): ", end="")
    try:
        choice = input().strip().lower()
        return choice == 'y'
    except:
        # 非交互式环境，继续运行
        return True

def check_and_install_python_deps():
    """检查并安装Python依赖"""
    print("\n🔍 检查 Python 依赖包...")
    
    # 检查requirements.txt
    requirements_file = Path(__file__).parent / "requirements.txt"
    if not requirements_file.exists():
        print("❌ requirements.txt 文件不存在")
        return False
    
    # 尝试安装依赖
    try:
        print("正在安装Python依赖包...")
        
        command = [
            sys.executable, "-m", "pip", "install", "-r", str(requirements_file),
            "-i", "https://pypi.tuna.tsinghua.edu.cn/simple",
            "--trusted-host", "pypi.tuna.tsinghua.edu.cn"
        ]
        
        process = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=300
        )
        
        if process.returncode == 0:
            print("✅ Python依赖包安装成功")
            return True
        else:
            print(f"❌ Python依赖包安装失败: {process.stderr}")
            print("请手动运行: pip install -r requirements.txt")
            return False
            
    except subprocess.TimeoutExpired:
        print("⚠️  依赖安装超时，请手动安装")
        return False
    except Exception as e:
        print(f"❌ 依赖安装失败: {e}")
        return False

def check_python_version():
    """检查Python版本"""
    print("🔍 检查 Python 版本...")
    
    if sys.version_info < (3, 7):
        print(f"❌ Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro} 版本过低")
        print("   需要 Python 3.7 或更高版本")
        return False
    
    print(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro} 符合要求")
    return True

def run_dependency_checker():
    """运行依赖检查器"""
    print("\n🔧 运行依赖检查器...")
    
    try:
        sys.path.insert(0, str(Path(__file__).parent / "src"))
        from dependency_checker import DependencyChecker
        
        checker = DependencyChecker()
        return checker.run()
    except ImportError:
        print("❌ 无法导入依赖检查器")
        return False
    except Exception as e:
        print(f"❌ 依赖检查失败: {e}")
        return False

def run_tesseract_installer():
    """运行Tesseract安装器"""
    print("\n🔧 运行 Tesseract 安装器...")
    
    try:
        sys.path.insert(0, str(Path(__file__).parent / "src"))
        from tesseract_installer import TesseractInstaller
        
        installer = TesseractInstaller()
        return installer.run()
    except ImportError:
        print("❌ 无法导入Tesseract安装器")
        return False
    except Exception as e:
        print(f"❌ Tesseract安装失败: {e}")
        return False

def launch_main_program():
    """启动主程序"""
    print("\n🚀 启动软著文档生成工具...")
    print("-" * 50)
    
    try:
        # 导入主程序
        sys.path.insert(0, str(Path(__file__).parent / "src"))
        import main
        
        # 运行主程序
        return main.main()
    except ImportError as e:
        print(f"❌ 无法导入主程序: {e}")
        return 1
    except Exception as e:
        print(f"❌ 启动失败: {e}")
        return 1

def main():
    """主函数"""
    print_banner()
    
    print("本脚本将自动检查并安装所有依赖，然后运行主程序")
    print()
    
    # 检查Python版本
    if not check_python_version():
        input("\n按回车键退出...")
        return 1
    
    # 运行依赖检查器
    if not run_dependency_checker():
        print("\n⚠️  依赖检查失败，继续运行可能导致错误")
    
    # 检查Tesseract
    if not check_and_install_tesseract():
        print("\n⚠️  Tesseract检查失败，OCR功能可能不可用")
    
    # 检查Python依赖
    if not check_and_install_python_deps():
        print("\n⚠️  Python依赖检查失败，继续运行可能导致错误")
    
    print("\n" + "=" * 70)
    print("环境检查完成，准备启动主程序")
    print("=" * 70)
    print()
    
    # 启动主程序
    return launch_main_program()

if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\n程序被用户中断")
        sys.exit(1)