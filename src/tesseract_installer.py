#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tesseract OCR 自动安装器
功能：自动检测并安装Tesseract OCR
"""

import os
import sys
import subprocess
import shutil
import tempfile
import zipfile
import requests
from pathlib import Path

class TesseractInstaller:
    """Tesseract OCR 安装器"""
    
    def __init__(self):
        self.current_dir = Path(__file__).parent.parent
        self.installers_dir = self.current_dir / "installers"
        self.tesseract_url = "https://github.com/UB-Mannheim/tesseract/wiki"
        self.offline_installer = self.installers_dir / "tesseract_installer.exe"
        
        # 确保installers目录存在
        self.installers_dir.mkdir(exist_ok=True)
    
    def check_tesseract_installed(self):
        """检查Tesseract是否已安装"""
        try:
            # 尝试运行tesseract --version
            result = subprocess.run(
                ['tesseract', '--version'],
                capture_output=True,
                text=True,
                timeout=10
            )
            
            if result.returncode == 0:
                # 提取版本信息
                version_line = result.stdout.split('\n')[0]
                version = version_line.split()[1] if len(version_line.split()) > 1 else "unknown"
                print(f"✅ Tesseract OCR 已安装 (版本: {version})")
                return True
            else:
                return False
                
        except (subprocess.SubprocessError, FileNotFoundError, OSError):
            # 检查常见安装路径
            possible_paths = [
                r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                r"D:\Program Files\Tesseract-OCR\tesseract.exe",
                r"D:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                str(Path.home() / "AppData" / "Local" / "Tesseract-OCR" / "tesseract.exe"),
                str(Path.home() / "AppData" / "Local" / "Programs" / "Tesseract-OCR" / "tesseract.exe")
            ]
            
            for path in possible_paths:
                if os.path.exists(path):
                    print(f"✅ 在以下位置找到 Tesseract OCR: {path}")
                    return True
            
            return False
    
    def download_tesseract_installer(self):
        """下载Tesseract安装程序"""
        print("📥 正在下载 Tesseract OCR 安装程序...")
        
        # 这里可以使用具体的下载链接，或提示用户手动下载
        # 由于网络环境不同，我们提供指引
        print(f"请访问以下网址下载 Tesseract OCR:")
        print(f"  {self.tesseract_url}")
        print()
        print("下载完成后，请将安装程序放入以下目录:")
        print(f"  {self.installers_dir}")
        print()
        print("或者，您可以从以下镜像下载:")
        print("  https://digi.bib.uni-mannheim.de/tesseract/")
        print()
        
        # 检查是否已有安装程序
        if self.offline_installer.exists():
            size_mb = self.offline_installer.stat().st_size / (1024 * 1024)
            print(f"✅ 已找到安装程序: {self.offline_installer.name} ({size_mb:.1f} MB)")
            return True
        else:
            print("⚠️  未找到安装程序，请手动下载")
            return False
    
    def install_tesseract(self):
        """安装Tesseract OCR"""
        print("\n🔧 开始安装 Tesseract OCR...")
        
        if not self.offline_installer.exists():
            print("❌ 安装程序不存在，请先下载")
            return False
        
        try:
            # 运行安装程序
            print(f"正在运行安装程序: {self.offline_installer.name}")
            
            # 使用subprocess运行安装程序
            process = subprocess.Popen(
                [str(self.offline_installer)],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                shell=True
            )
            
            print("✅ 安装程序已启动，请按照提示完成安装")
            print("💡 安装时请注意:")
            print("   1. 选择安装位置")
            print("   2. 勾选'添加到PATH环境变量'")
            print("   3. 安装中文语言包")
            
            # 等待安装完成
            try:
                stdout, stderr = process.communicate(timeout=300)  # 5分钟超时
                
                if process.returncode == 0:
                    print("✅ Tesseract OCR 安装成功")
                    return True
                else:
                    print(f"❌ 安装失败，返回码: {process.returncode}")
                    if stderr:
                        print(f"错误信息: {stderr}")
                    return False
                    
            except subprocess.TimeoutExpired:
                print("⚠️  安装超时，安装程序仍在运行中")
                return False
                
        except Exception as e:
            print(f"❌ 安装过程中发生错误: {e}")
            return False
    
    def check_tesseract_path(self):
        """检查Tesseract是否在PATH中"""
        print("\n🔍 检查系统PATH配置...")
        
        try:
            # 获取系统PATH
            system_path = os.environ.get('PATH', '')
            
            # 检查常见Tesseract路径
            tesseract_paths = [
                r"C:\Program Files\Tesseract-OCR",
                r"C:\Program Files (x86)\Tesseract-OCR",
                r"Tesseract-OCR",
                r"Tesseract"
            ]
            
            found = False
            for path in tesseract_paths:
                if path.lower() in system_path.lower():
                    print(f"✅ PATH中包含Tesseract路径: {path}")
                    found = True
            
            if not found:
                print("⚠️  Tesseract可能不在PATH中")
                print("    如果遇到'找不到tesseract'错误，请检查PATH")
            
            return found
            
        except Exception as e:
            print(f"❌ 检查PATH时发生错误: {e}")
            return False
    
    def create_install_script(self):
        """创建安装脚本"""
        print("\n📝 创建自动化安装脚本...")
        
        # 创建批处理文件
        batch_file = self.current_dir / "install_tesseract.bat"
        
        batch_content = '''@echo off
chcp 65001 > nul
echo ==========================================
echo Tesseract OCR 自动化安装脚本
echo ==========================================

echo.
echo 步骤 1: 检查是否已安装 Tesseract...
where tesseract >nul 2>&1
if %errorlevel% equ 0 (
    echo ✅ Tesseract 已安装
    call tesseract --version
    goto :end
)

echo.
echo 步骤 2: 检查安装程序...
if exist "installers\\tesseract_installer.exe" (
    echo ✅ 找到安装程序
    echo.
    echo 步骤 3: 运行安装程序...
    echo 请按照安装向导完成安装
    echo.
    echo 重要提示:
    echo   1. 建议使用默认安装路径
    echo   2. 必须勾选"添加到 PATH 环境变量"
    echo   3. 必须安装中文语言包
    echo.
    echo 安装完成后，请重启命令提示符
    start /wait "Tesseract OCR Installer" "installers\\tesseract_installer.exe"
    
    echo.
    echo 步骤 4: 验证安装...
    where tesseract >nul 2>&1
    if %errorlevel% equ 0 (
        echo ✅ Tesseract 安装成功
        call tesseract --version
    ) else (
        echo ❌ Tesseract 安装失败
        echo 请手动运行安装程序: installers\\tesseract_installer.exe
    )
) else (
    echo ❌ 找不到安装程序
    echo 请将 Tesseract 安装程序放入 installers 文件夹
    echo 下载地址: https://github.com/UB-Mannheim/tesseract/wiki
)

:end
echo.
echo 安装完成，按任意键退出...
pause >nul
'''
        
        try:
            with open(batch_file, 'w', encoding='gbk') as f:
                f.write(batch_content)
            
            print(f"✅ 创建安装脚本: {batch_file.name}")
            print("   使用方法: 双击此文件运行")
            
            # 创建Python安装脚本
            py_script = self.current_dir / "install_tesseract.py"
            py_content = '''#!/usr/bin/env python3
import os
import subprocess
import sys

def check_tesseract():
    """检查Tesseract是否已安装"""
    try:
        result = subprocess.run(
            ['tesseract', '--version'],
            capture_output=True,
            text=True,
            timeout=10
        )
        if result.returncode == 0:
            return True, result.stdout.strip()
        else:
            return False, None
    except:
        return False, None

def main():
    print("Tesseract OCR 安装检查")
    print("=" * 40)
    
    # 检查是否已安装
    installed, info = check_tesseract()
    if installed:
        print("✅ Tesseract OCR 已安装")
        print(f"版本信息: {info}")
        return 0
    
    print("❌ Tesseract OCR 未安装")
    print()
    print("请按照以下步骤操作:")
    print("1. 访问 https://github.com/UB-Mannheim/tesseract/wiki")
    print("2. 下载最新版 Tesseract 安装程序")
    print("3. 将安装程序放入 installers 文件夹")
    print("4. 运行 install_tesseract.bat")
    print()
    
    return 1

if __name__ == "__main__":
    sys.exit(main())
'''
            
            with open(py_script, 'w', encoding='utf-8') as f:
                f.write(py_content)
            
            print(f"✅ 创建Python安装脚本: {py_script.name}")
            
            return True
            
        except Exception as e:
            print(f"❌ 创建安装脚本失败: {e}")
            return False
    
    def run(self):
        """运行安装流程"""
        print("=" * 60)
        print("Tesseract OCR 自动安装器")
        print("=" * 60)
        print()
        
        # 检查是否已安装
        if self.check_tesseract_installed():
            print("\n✅ Tesseract OCR 已安装，无需再次安装")
            self.check_tesseract_path()
            return True
        
        print("⚠️  Tesseract OCR 未安装")
        print()
        
        # 下载安装程序
        if not self.download_tesseract_installer():
            print("\n❌ 安装程序不可用，无法继续")
            return False
        
        # 确认安装
        print("是否立即安装 Tesseract OCR？ (y/n): ", end="")
        choice = input().strip().lower()
        
        if choice != 'y':
            print("安装已取消")
            return False
        
        # 执行安装
        if self.install_tesseract():
            # 检查PATH
            self.check_tesseract_path()
            
            # 创建安装脚本
            self.create_install_script()
            
            print("\n🎉 Tesseract OCR 安装完成！")
            print("💡 提示：安装完成后建议重启命令提示符")
            return True
        else:
            print("\n❌ Tesseract OCR 安装失败")
            return False

def main():
    """主函数"""
    installer = TesseractInstaller()
    
    try:
        if installer.run():
            return 0
        else:
            return 1
    except KeyboardInterrupt:
        print("\n\n安装被用户中断")
        return 1
    except Exception as e:
        print(f"\n\n安装过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main())