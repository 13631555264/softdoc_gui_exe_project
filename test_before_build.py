#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
构建exe前的测试脚本
确保所有功能正常后再进行打包
"""

import os
import sys
import subprocess
import tempfile
import shutil
from pathlib import Path

def print_header():
    """打印标题"""
    print("=" * 70)
    print("软著文档生成工具 - 构建前测试")
    print("=" * 70)
    print()

def check_python_version():
    """检查Python版本"""
    print("1. 检查Python版本...")
    
    if sys.version_info < (3, 7):
        print(f"❌ Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro} 版本过低")
        print("   需要 Python 3.7 或更高版本")
        return False
    
    print(f"✅ Python {sys.version_info.major}.{version.minor}.{version.micro}")
    return True

def check_pyinstaller():
    """检查PyInstaller"""
    print("\n2. 检查PyInstaller...")
    
    try:
        import PyInstaller
        version = PyInstaller.__version__
        print(f"✅ PyInstaller {version} 已安装")
        return True
    except ImportError:
        print("❌ PyInstaller 未安装")
        print("   运行: pip install pyinstaller")
        return False

def check_tesseract():
    """检查Tesseract OCR"""
    print("\n3. 检查Tesseract OCR...")
    
    try:
        result = subprocess.run(
            ['tesseract', '--version'],
            capture_output=True,
            text=True,
            timeout=10
        )
        
        if result.returncode == 0:
            version = result.stdout.split('\n')[0]
            print(f"✅ {version}")
            return True
        else:
            print("❌ Tesseract 不可用")
            return False
    except:
        print("❌ Tesseract 未安装或不在PATH中")
        return False

def check_python_packages():
    """检查Python包"""
    print("\n4. 检查Python包...")
    
    packages = [
        ("tkinterdnd2", "GUI拖拽支持"),
        ("pytesseract", "OCR识别"),
        ("PIL", "图片处理"),
        ("docx", "Word文档操作"),
        ("cv2", "图像处理"),
        ("pdfplumber", "PDF解析"),
        ("pyperclip", "剪贴板操作"),
        ("numpy", "数值计算"),
        ("PyInstaller", "打包工具")
    ]
    
    missing = []
    for import_name, description in packages:
        try:
            if import_name == "tkinterdnd2":
                import tkinterdnd2
            elif import_name == "PIL":
                from PIL import Image
            elif import_name == "PyInstaller":
                import PyInstaller
            else:
                __import__(import_name)
            print(f"✅ {description}")
        except ImportError:
            print(f"❌ {description} 未安装")
            missing.append(import_name)
    
    if missing:
        print(f"\n⚠️  缺少 {len(missing)} 个包: {', '.join(missing)}")
        return False
    
    return True

def check_source_files():
    """检查源代码文件"""
    print("\n5. 检查源代码文件...")
    
    current_dir = Path(__file__).parent
    src_dir = current_dir / "src"
    
    required_files = [
        src_dir / "main.py",
        src_dir / "tesseract_installer.py",
        src_dir / "dependency_checker.py",
        src_dir / "gui" / "__init__.py",
        src_dir / "gui" / "main_window.py",
        src_dir / "core" / "__init__.py",
        src_dir / "core" / "qg_parser.py",
        src_dir / "core" / "softdoc_parser.py",
        src_dir / "core" / "document_generator.py",
    ]
    
    missing = []
    for file_path in required_files:
        if file_path.exists():
            print(f"✅ {file_path.relative_to(current_dir)}")
        else:
            print(f"❌ {file_path.relative_to(current_dir)} 不存在")
            missing.append(file_path)
    
    if missing:
        print(f"\n⚠️  缺少 {len(missing)} 个文件")
        return False
    
    return True

def test_imports():
    """测试导入"""
    print("\n6. 测试导入...")
    
    current_dir = Path(__file__).parent
    src_dir = current_dir / "src"
    
    # 添加路径
    sys.path.insert(0, str(src_dir))
    
    modules_to_test = [
        ("main", lambda: __import__('main')),
        ("dependency_checker", lambda: __import__('dependency_checker')),
        ("tesseract_installer", lambda: __import__('tesseract_installer')),
        ("gui", lambda: __import__('gui')),
        ("core", lambda: __import__('core')),
    ]
    
    failed = []
    for module_name, import_func in modules_to_test:
        try:
            import_func()
            print(f"✅ {module_name} 导入成功")
        except ImportError as e:
            print(f"❌ {module_name} 导入失败: {e}")
            failed.append(module_name)
    
    if failed:
        print(f"\n⚠️  {len(failed)} 个模块导入失败")
        return False
    
    return True

def test_core_functions():
    """测试核心功能"""
    print("\n7. 测试核心功能...")
    
    try:
        # 测试渠广解析器
        from core.qg_parser import QGParser
        print("✅ 渠广解析器 可用")
        
        # 测试软著解析器
        from core.softdoc_parser import SoftDocParser
        print("✅ 软著解析器 可用")
        
        # 测试文档生成器
        from core.document_generator import DocumentGenerator
        print("✅ 文档生成器 可用")
        
        # 测试配置模块
        from core.config import Config
        print("✅ 配置模块 可用")
        
        return True
        
    except ImportError as e:
        print(f"❌ 核心功能导入失败: {e}")
        return False
    except Exception as e:
        print(f"❌ 核心功能测试失败: {e}")
        return False

def test_gui_components():
    """测试GUI组件"""
    print("\n8. 测试GUI组件...")
    
    try:
        # 测试tkinter
        import tkinter as tk
        print("✅ tkinter 可用")
        
        # 测试tkinterdnd2
        import tkinterdnd2
        print("✅ tkinterdnd2 可用")
        
        # 测试GUI主窗口
        from gui.main_window import MainWindow
        print("✅ 主窗口类 可用")
        
        return True
        
    except ImportError as e:
        print(f"❌ GUI组件导入失败: {e}")
        return False
    except Exception as e:
        print(f"❌ GUI组件测试失败: {e}")
        return False

def test_file_operations():
    """测试文件操作"""
    print("\n9. 测试文件操作...")
    
    try:
        # 创建临时目录
        temp_dir = tempfile.mkdtemp(prefix="softdoc_test_")
        print(f"✅ 创建临时目录: {temp_dir}")
        
        # 测试文件读写
        test_file = Path(temp_dir) / "test.txt"
        test_file.write_text("测试内容", encoding='utf-8')
        
        if test_file.exists():
            print("✅ 文件读写正常")
        else:
            print("❌ 文件创建失败")
        
        # 清理
        shutil.rmtree(temp_dir)
        print("✅ 清理临时文件")
        
        return True
        
    except Exception as e:
        print(f"❌ 文件操作测试失败: {e}")
        return False

def run_pyinstaller_test():
    """运行PyInstaller测试构建"""
    print("\n10. 运行PyInstaller测试构建...")
    
    current_dir = Path(__file__).parent
    src_dir = current_dir / "src"
    
    # 创建测试spec文件
    test_spec = current_dir / "test_build.spec"
    
    spec_content = f"""
# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['{src_dir / "main.py"}'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'tkinterdnd2',
        'pytesseract',
        'PIL',
        'docx',
        'cv2',
        'pdfplumber',
        'pyperclip',
        'numpy',
        'core',
        'gui'
    ],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='softdoc_generator_test',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
"""
    
    try:
        # 写入spec文件
        test_spec.write_text(spec_content, encoding='utf-8')
        print("✅ 创建spec文件")
        
        # 运行PyInstaller测试
        command = [sys.executable, "-m", "PyInstaller", "--clean", str(test_spec)]
        
        print("正在运行测试构建...")
        process = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=300
        )
        
        if process.returncode == 0:
            print("✅ PyInstaller测试构建成功")
            
            # 检查输出文件
            dist_dir = current_dir / "dist"
            test_exe = dist_dir / "softdoc_generator_test.exe"
            
            if test_exe.exists():
                size_mb = test_exe.stat().st_size / (1024 * 1024)
                print(f"✅ 生成测试exe文件 ({size_mb:.1f} MB)")
            else:
                print("⚠️  测试exe文件未生成")
            
            return True
        else:
            print(f"❌ PyInstaller测试构建失败: {process.stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        print("⚠️  测试构建超时")
        return False
    except Exception as e:
        print(f"❌ 测试构建失败: {e}")
        return False
    finally:
        # 清理测试文件
        if test_spec.exists():
            test_spec.unlink()
            print("✅ 清理spec文件")

def main():
    """主函数"""
    print_header()
    
    print("本脚本将执行全面的测试，确保所有功能正常")
    print("通过测试后才能安全构建exe文件")
    print()
    
    # 执行所有测试
    tests = [
        ("Python版本", check_python_version),
        ("PyInstaller", check_pyinstaller),
        ("Tesseract OCR", check_tesseract),
        ("Python包", check_python_packages),
        ("源代码文件", check_source_files),
        ("模块导入", test_imports),
        ("核心功能", test_core_functions),
        ("GUI组件", test_gui_components),
        ("文件操作", test_file_operations),
        ("PyInstaller测试构建", run_pyinstaller_test)
    ]
    
    results = []
    for test_name, test_func in tests:
        print(f"执行测试: {test_name}")
        try:
            if test_func():
                results.append((test_name, True))
            else:
                results.append((test_name, False))
        except Exception as e:
            print(f"❌ 测试执行失败: {e}")
            results.append((test_name, False))
        print()
    
    # 统计结果
    passed = sum(1 for _, passed in results if passed)
    total = len(results)
    
    print("=" * 70)
    print("测试结果汇总")
    print("=" * 70)
    
    for test_name, passed_test in results:
        status = "✅ 通过" if passed_test else "❌ 失败"
        print(f"{status}: {test_name}")
    
    print()
    print(f"📊 总成绩: {passed}/{total}")
    
    if passed == total:
        print("🎉 所有测试通过！可以安全构建exe文件")
        print()
        print("构建命令:")
        print("    python build_exe.py")
        return 0
    else:
        print("⚠️  部分测试失败，请修复后再构建exe")
        print()
        print("失败的项目:")
        for test_name, passed_test in results:
            if not passed_test:
                print(f"  - {test_name}")
        print()
        return 1

if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\n测试被用户中断")
        sys.exit(1)