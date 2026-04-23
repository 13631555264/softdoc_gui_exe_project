#!/usr/bin/env python3

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
            print(" [OK]")
            if result.stdout.strip():
                print(f"  输出: {result.stdout.strip()}")
            return True
        else:
            print(" [FAIL]")
            if result.stderr.strip():
                print(f"  错误: {result.stderr.strip()}")
            return False
            
    except subprocess.TimeoutExpired:
        print(" [TIMEOUT]")
        return False
        
    except Exception as e:
        print(f" [ERROR] {e}")
        return False

def main():
    print("软著文档生成工具 - 环境设置")
    print("=" * 50)
    
    # 检查Python版本
    if sys.version_info < (3, 7):
        print("[ERROR] 需要 Python 3.7 或更高版本")
        return 1
    
    print(f"[OK] Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")
    
    # 升级pip
    run_command([sys.executable, "-m", "pip", "install", "--upgrade", "pip"], "升级pip")
    
    # 安装依赖
    print("
安装依赖包...")
    success = run_command([
        sys.executable, "-m", "pip", "install", "-r", "requirements.txt",
        "-i", "https://pypi.tuna.tsinghua.edu.cn/simple",
        "--trusted-host", "pypi.tuna.tsinghua.edu.cn"
    ], "安装依赖包")
    
    if not success:
        print("[WARNING] 依赖安装失败，尝试手动安装")
        print("   运行: pip install -r requirements.txt")
    
    # 检查Tesseract
    print("
检查Tesseract OCR...")
    if run_command(["tesseract", "--version"], "Tesseract检查"):
        print("[OK] Tesseract OCR 已安装")
    else:
        print("[ERROR] Tesseract OCR 未安装")
        print("   请下载并安装: https://github.com/UB-Mannheim/tesseract/wiki")
    
    print("
" + "=" * 50)
    print("环境设置完成！")
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
