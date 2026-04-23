#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
依赖检查与自动安装器
功能：检查Python依赖并自动安装
"""

import os
import sys
import subprocess
import importlib
from pathlib import Path

class DependencyChecker:
    """依赖检查器"""
    
    def __init__(self):
        self.current_dir = Path(__file__).parent.parent
        self.requirements_file = self.current_dir / "requirements.txt"
        
        # 核心依赖包列表
        self.core_packages = [
           
            ("numpy", "numpy", "数值计算")
        ]
        
        # 可选包（exe构建需要）
        self.optional_packages = [
            ("pyinstaller", "pyinstaller", "exe打包工具"),
            ("requests", "requests", "HTTP请求库")
        ]
    
    def check_package(self, import_name, package_name, description):
        """检查单个包是否安装"""
        try:
            if import_name == "tkinterdnd2":
                import tkinterdnd2
            elif import_name == "PIL":
                from PIL import Image
            else:
                importlib.import_module(import_name)
            
            version = self.get_package_version(package_name)
            version_str = f" (版本: {version})" if version else ""
            
            return True, f"✅ {description}{version_str}"
            
        except ImportError:
            return False, f"❌ {description} 未安装"
        
        except Exception as e:
            return False, f"⚠️  {description} 检查失败: {e}"
    
    def get_package_version(self, package_name):
        """获取包的版本"""
        try:
            result = subprocess.run(
                ['pip', 'show', package_name],
                capture_output=True,
                text=True,
                timeout=10
            )
            
            if result.returncode == 0:
                for line in result.stdout.split('\n'):
                    if line.startswith('Version:'):
                        return line.split(':')[1].strip()
        except:
            pass
        
        return None
    
    def install_package(self, package_name):
        """安装单个包"""
        try:
            print(f"正在安装 {package_name}...")
            
            # 使用国内镜像源加速
            command = [
                'pip', 'install', package_name,
                '-i', 'https://pypi.tuna.tsinghua.edu.cn/simple',
                '--trusted-host', 'pypi.tuna.tsinghua.edu.cn'
            ]
            
            process = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=300  # 5分钟超时
            )
            
            if process.returncode == 0:
                print(f"✅ {package_name} 安装成功")
                return True
            else:
                print(f"❌ {package_name} 安装失败: {process.stderr}")
                return False
                
        except subprocess.TimeoutExpired:
            print(f"⚠️  {package_name} 安装超时")
            return False
            
        except Exception as e:
            print(f"❌ {package_name} 安装失败: {e}")
            return False
    
    def install_from_requirements(self):
        """从requirements.txt安装依赖"""
        if not self.requirements_file.exists():
            print("❌ requirements.txt 文件不存在")
            return False
        
        try:
            print("📦 正在从 requirements.txt 安装依赖...")
            
            command = [
                'pip', 'install', '-r', str(self.requirements_file),
                '-i', 'https://pypi.tuna.tsinghua.edu.cn/simple',
                '--trusted-host', 'pypy.tuna.tsinghua.edu.cn'
            ]
            
            process = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=600  # 10分钟超时
            )
            
            if process.returncode == 0:
                print("✅ 所有依赖安装成功")
                return True
            else:
                print(f"❌ 依赖安装失败: {process.stderr}")
                return False
                
        except subprocess.TimeoutExpired:
            print("⚠️  依赖安装超时，可能网络较慢")
            return False
            
        except Exception as e:
            print(f"❌ 依赖安装失败: {e}")
            return False
    
    def check_all_packages(self):
        """检查所有包"""
        print("📊 检查Python依赖包...")
        print("-" * 50)
        
        results = []
        missing_packages = []
        
        total = len(self.core_packages) + len(self.optional_packages)
        passed = 0
        
        # 检查核心包
        print("核心包检查:")
        for import_name, package_name, description in self.core_packages:
            installed, message = self.check_package(import_name, package_name, description)
            results.append((package_name, installed, message))
            
            if installed:
                passed += 1
                print(f"  {message}")
            else:
                missing_packages.append(package_name)
                print(f"  {message}")
        
        # 检查可选包
        print("\n可选包检查:")
        for import_name, package_name, description in self.optional_packages:
            installed, message = self.check_package(import_name, package_name, description)
            results.append((package_name, installed, message))
            
            if installed:
                passed += 1
                print(f"  {message}")
            else:
                # 可选包不强制安装
                print(f"  {message}")
        
        print("\n" + "=" * 50)
        print(f"依赖检查结果: {passed}/{total}")
        
        return results, missing_packages
    
    def fix_missing_packages(self, missing_packages):
        """修复缺失的包"""
        if not missing_packages:
            print("✅ 所有依赖包已安装")
            return True
        
        print(f"\n⚠️  缺少 {len(missing_packages)} 个依赖包: {', '.join(missing_packages)}")
        print("是否自动安装？ (y/n): ", end="")
        
        try:
            choice = input().strip().lower()
        except:
            # 非交互式环境，自动安装
            choice = "y"
        
        if choice != 'y':
            print("安装已取消")
            return False
        
        print("\n开始安装缺失的依赖包...")
        print("-" * 50)
        
        success = True
        for package in missing_packages:
            if not self.install_package(package):
                success = False
        
        if success:
            print("\n✅ 所有缺失依赖已安装")
        else:
            print("\n⚠️  部分依赖安装失败")
        
        return success
    
    def check_python_version(self):
        """检查Python版本"""
        print("🐍 检查Python环境...")
        
        version = sys.version_info
        
        if version.major == 3 and version.minor >= 7:
            print(f"✅ Python {version.major}.{version.minor}.{version.micro} 符合要求")
            return True
        else:
            print(f"❌ Python {version.major}.{version.minor}.{version.micro} 版本过低")
            print("   需要 Python 3.7 或更高版本")
            return False
    
    def create_setup_script(self):
        """创建设置脚本"""
        print("\n📝 创建环境设置脚本...")
        
        # 创建Windows批处理文件
        batch_file = self.current_dir / "setup_environment.bat"
        
        batch_content = '''@echo off
chcp 65001 > nul
echo ==========================================
echo 软著文档生成工具 - 环境设置脚本
echo ==========================================

echo.
echo 步骤 1: 检查Python版本...
python --version
if errorlevel 1 (
    echo ❌ Python未安装或不在PATH中
    echo    请安装Python 3.7+并添加到PATH
    pause
    exit /b 1
)

echo.
echo 步骤 2: 升级pip...
python -m pip install --upgrade pip

echo.
echo 步骤 3: 安装Python依赖包...
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn

echo.
echo 步骤 4: 检查Tesseract OCR...
where tesseract >nul 2>&1
if errorlevel 1 (
    echo ⚠️  Tesseract未安装
    echo    请下载并安装Tesseract OCR
    echo    下载地址: https://github.com/UB-Mannheim/tesseract/wiki
    echo    安装后请将路径添加到PATH
)


echo.
echo 步骤 5: 验证安装...
echo.
echo 检查依赖包...
python -c "
try:
    import tkinterdnd2
    print('✅ tkinterdnd2 已安装')
except:
    print('❌ tkinterdnd2 未安装')
"
echo.

echo.
echo ==========================================
echo 环境设置完成！
echo ==========================================
echo.
echo 使用方法：
echo   1. 运行软著文档生成工具：python src\\main.py
echo   2. 或运行自动安装版本：python run_with_install.py
echo.
pause
'''
        
        try:
            with open(batch_file, 'w', encoding='gbk') as f:
                f.write(batch_content)
            
            # 创建Python版本
            py_file = self.current_dir / "setup_environment.py"
            
            py_content = '''#!/usr/bin/env python3

import sys
import subprocess

def run_command(cmd, description):
    """运行命令并显示结果"""
    print(f"{description}...", end="")
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300
        )
        
        if result.returncode == 0:
            print(" ✅")
            if result.stdout.strip():
                print(f"  输出: {result.stdout.strip()}")
            return True
        else:
            print(" ❌")
            if result.stderr.strip():
                print(f"  错误: {result.stderr.strip()}")
            return False
            
    except subprocess.TimeoutExpired:
        print(" ⏱️ 超时")
        return False
        
    except Exception as e:
        print(f" ❌ {e}")
        return False

def main():
    print("软著文档生成工具 - 环境设置")
    print("=" * 50)
    
    # 检查Python版本
    if sys.version_info < (3, 7):
        print("❌ 需要 Python 3.7 或更高版本")
        return 1
    
    print(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")
    
    # 升级pip
    run_command([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], "升级pip")
    
    # 安装依赖
    print("\n安装依赖包...")
    success = run_command([
        sys.executable, "-m", "pip", "install", "-r", "requirements.txt",
        "-i", "https://pypi.tuna.tsinghua.edu.cn/simple",
        "--trusted-host", "pypi.tuna.tsinghua.edu.cn"
    ], "安装依赖包")
    
    if not success:
        print("⚠️  依赖安装失败，尝试手动安装")
        print("   运行: pip install -r requirements.txt")
    
    # 检查Tesseract
    print("\n检查Tesseract OCR...")
    if run_command(["tesseract", "--version"], "Tesseract检查"):
        print("✅ Tesseract OCR 已安装")
    else:
        print("❌ Tesseract OCR 未安装")
        print("   请下载并安装: https://github.com/UB-Mannheim/tesseract/wiki")
    
    print("\n" + "=" * 50)
    print("环境设置完成！")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
'''
            
            with open(py_file, 'w', encoding='utf-8') as f:
                f.write(py_content)
            
            print(f"✅ 创建环境设置脚本:")
            print(f"   - {batch_file.name} (Windows批处理)")
            print(f"   - {py_file.name} (Python脚本)")
            
            return True
            
        except Exception as e:
            print(f"❌ 创建设置脚本失败: {e}")
            return False
    
    def run(self):
        """运行依赖检查"""
        print("=" * 60)
        print("依赖检查与自动安装器")
        print("=" * 60)
        print()
        
        # 检查Python版本
        if not self.check_python_version():
            return False
        
        print()
        
        # 检查所有包
        results, missing_packages = self.check_all_packages()
        
        # 如果缺少依赖，尝试修复
        if missing_packages:
            if not self.fix_missing_packages(missing_packages):
                print("\n⚠️  依赖不完整，可能导致运行错误")
                return False
        
        # 创建设置脚本
        self.create_setup_script()
        
        print("\n" + "=" * 60)
        print("依赖检查完成！")
        
        return True

def main():
    """主函数"""
    checker = DependencyChecker()
    
    try:
        if checker.run():
            return 0
        else:
            return 1
    except KeyboardInterrupt:
        print("\n\n操作被用户中断")
        return 1
    except Exception as e:
        print(f"\n\n操作过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())