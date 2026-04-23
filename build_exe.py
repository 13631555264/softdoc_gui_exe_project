#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
软著文档生成工具 - exe构建脚本
自动生成可执行的Windows应用程序
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

def print_banner():
    """打印横幅"""
    print("=" * 70)
    print("软著文档生成工具 - exe自动化构建")
    print("=" * 70)
    print()

def check_prerequisites():
    """检查前置条件"""
    print("🔍 检查前置条件...")
    
    # 检查PyInstaller
    try:
        import PyInstaller
        version = PyInstaller.__version__
        print(f"✅ PyInstaller {version} 已安装")
    except ImportError:
        print("❌ PyInstaller 未安装")
        print("   运行: pip install pyinstaller")
        return False
    
    # 检查工作目录
    current_dir = Path(__file__).parent
    src_dir = current_dir / "src"
    
    if not src_dir.exists():
        print("❌ src 目录不存在")
        return False
    
    print(f"✅ 工作目录: {current_dir}")
    print(f"✅ 源代码目录: {src_dir}")
    
    return True

def clean_build_directories():
    """清理构建目录"""
    print("\n🧹 清理构建目录...")
    
    current_dir = Path(__file__).parent
    dirs_to_clean = ["build", "dist", "__pycache__"]
    
    cleaned = []
    for dir_name in dirs_to_clean:
        dir_path = current_dir / dir_name
        
        if dir_path.exists():
            try:
                shutil.rmtree(dir_path)
                cleaned.append(dir_name)
                print(f"✅ 清理: {dir_name}")
            except Exception as e:
                print(f"⚠️  清理 {dir_name} 失败: {e}")
    
    # 清理pycache
    for pycache_dir in current_dir.rglob("__pycache__"):
        try:
            shutil.rmtree(pycache_dir)
            print(f"✅ 清理: {pycache_dir.relative_to(current_dir)}")
        except:
            pass
    
    return len(cleaned) > 0

def create_spec_file():
    """创建PyInstaller spec文件"""
    print("\n📄 创建spec文件...")
    
    current_dir = Path(__file__).parent
    src_dir = current_dir / "src"
    
    spec_content = f"""# -*- mode: python ; coding: utf-8 -*-
# 软著文档生成工具 - GUI版本 spec文件
# 生成时间: 2026-04-22

block_cipher = None

a = Analysis(
    ['{src_dir / "main.py"}'],
    pathex=[{repr(str(current_dir))}],
    binaries=[],
    datas=[
        # 包含配置文件和资源
        ('{src_dir / "core" / "config.py"}', 'core'),
        ('{src_dir / "gui" / "__init__.py"}', 'gui'),
        # 你可以添加其他资源文件
    ],
    hiddenimports=[
        'tkinterdnd2',
        'pytesseract',
        'pytesseract.tesseract_cmd',
        'PIL',
        'PIL._imaging',
        'PIL._imagingtk',
        'PIL._imagingmath',
        'PIL._imagingmorph',
        'docx',
        'docx.oxml',
        'docx.opc',
        'docx.shared',
        'cv2',
        'pdfplumber',
        'pyperclip',
        'numpy',
        'numpy.core._multiarray_umath',
        'core',
        'gui'
    ],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
    optimize=0,
)

# 添加pyz和exe
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='softdoc_generator_gui',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    runtime_tmpdir=None,
    console=False,  # 不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='softdoc_generator_gui',
)
"""
    
    spec_file = current_dir / "softdoc_generator_gui.spec"
    
    try:
        spec_file.write_text(spec_content, encoding='utf-8')
        print(f"✅ 创建spec文件: {spec_file.name}")
        return True
    except Exception as e:
        print(f"❌ 创建spec文件失败: {e}")
        return False

def run_pyinstaller():
    """运行PyInstaller构建"""
    print("\n🔧 运行PyInstaller构建...")
    
    current_dir = Path(__file__).parent
    spec_file = current_dir / "softdoc_generator_gui.spec"
    
    if not spec_file.exists():
        print("❌ spec文件不存在")
        return False
    
    print("正在构建exe文件，这可能需要几分钟...")
    print("请耐心等待...")
    
    try:
        command = [
            sys.executable, "-m", "PyInstaller",
            "--clean",
            "--noconfirm",
            str(spec_file)
        ]
        
        process = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=600,  # 10分钟超时
        )
        
        if process.returnupcode == 0:
            print("✅ PyInstaller构建成功")
            return True
        else:
            print(f"❌ PyInstaller构建失败")
            if process.stderr:
                print(f"错误信息: {process.stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        print("⚠️  构建超时，但可能仍在进行中")
        return True
    except Exception as e:
        print(f"❌ 构建过程中发生错误: {e}")
        return False

def verify_exe_file():
    """验证生成的exe文件"""
    print("\n🔍 验证exe文件...")
    
    current_dir = Path(__file__).parent
    dist_dir = current_dir / "dist"
    exe_path = dist_dir / "softdoc_generator_gui.exe"
    
    if not exe_path.exists():
        print("❌ exe文件不存在")
        
        # 检查是否有其他可能的名称
        for file in dist_dir.glob("*.exe"):
            print(f"  找到可能的exe文件: {file.name}")
            return True
        
        return False
    
    # 获取文件信息
    try:
        size_bytes = exe_path.stat().st_size
        size_mb = size_bytes / (1024 * 1024)
        
        print(f"✅ exe文件已生成")
        print(f"  路径: {exe_path}")
        print(f"  大小: {size_mb:.1f} MB ({size_bytes:,} 字节)")
        
        # 检查文件是否可执行
        if os.access(exe_path, os.X_OK):
            print("✅ 文件可执行")
        else:
            print("⚠️  文件可能无法执行")
        
        return True
        
    except Exception as e:
        print(f"❌ 验证exe文件失败: {e}")
        return False

def create_debug_info():
    """创建调试信息"""
    print("\n📋 创建调试信息...")
    
    current_dir = Path(__file__).parent
    
    # 收集系统信息
    debug_info = {
        "python_version": sys.version,
        "platform": sys.platform,
        "current_dir": str(current_dir),
        "build_time": "2026-04-22",
        "project": "软著文档生成工具",
        "version": "1.0.0"
    }
    
    # 写入文件
    debug_file = current_dir / "build_info.json"
    
    import json
    try:
        debug_file.write_text(json.dumps(debug_info, indent=2, ensure_ascii=False), encoding='utf-8')
        print(f"✅ 创建调试信息文件: {debug_file.name}")
    except Exception as e:
        print(f"⚠️  创建调试信息失败: {e}")
    
    # 创建使用说明
    readme_file = current_dir / "README_EXE.md"
    
    readme_content = """# 软著文档生成工具 - exe版本

## 🚀 快速开始

### 1. 运行程序
双击 `softdoc_generator_gui.exe` 即可启动

### 2. 使用说明

#### 选择文件
1. **软著文件**: PDF或图片格式
2. **渠广文件**: txt格式 (GBK编码)
3. **模板文件夹**: 包含docx模板

#### 生成文档
1. 选择所有文件
2. 点击"开始生成文档"
3. 查看生成的文档

### 3. 输出结构

```
输出文件夹/
├── 游戏名称1/
│   ├── 游戏名称1-单机-公司名称.docx
│   ├── 游戏名称1-免责-公司名称.docx
│   └── 游戏名称1-授权书-主体公司名称-授权方授权方公司.docx
├── 游戏名称2/
│   └── ...
└── 处理日志.txt
```

## 🔧 功能特点

### 1. 自动化处理
- 自动解析文件内容
- 智能判断授权书需求
- 按游戏名称组织输出

### 2. OCR支持
- 支持图片格式软著
- 自动识别文字
- 提取关键信息

### 3. 用户友好
- 图形化界面
- 拖拽文件选择
- 配置自动保存

### 4. 多文档生成
- 单机游戏承诺函
- 免责承诺函
- 授权书（当需要时）

## 💡 使用技巧

### 文件选择
- **拖拽**: 直接将文件拖到输入框
- **浏览**: 点击"浏览"按钮选择文件
- **批量**: 点击"选择所有文件"一次选择多个

### 配置管理
- **自动保存**: 每次选择后自动保存
- **历史记录**: 查看最近使用的文件
- **高级选项**: 自定义OCR参数

### 错误处理
- **详细日志**: 查看处理过程和错误信息
- **保存日志**: 将日志保存到文件
- **复制日志**: 复制到剪贴板

## 🔍 故障排除

### 常见问题

**Q1: 启动时报错"缺少DLL"**
```
解决方案: 安装Visual C++ Redistributable
下载地址: https://aka.ms/vs/17/release/vc_redist.x64.exe
```

**Q2: OCR识别失败**
```
解决方案: 确保已安装Tesseract OCR
下载地址: https://github.com/UB-Mannheim/tesseract/wiki
```

**Q2: 中文显示乱码**
```
解决方案: 尝试修改配置文件中的编码设置
```

**Q3: 模板文件错误**
```
解决方案: 确保模板文件格式正确
```

### 获取帮助

1. **查看日志**: 在日志文件中查找错误信息
2. **提供信息**: 包括错误消息和系统配置
3. **联系支持**: 提供详细的故障描述

## 📄 版本信息

| 版本 | 说明 | 日期 |
|------|------|------|
| 1.0.0 | 初始版本 | 2026-04-22 |

### 系统要求
- **操作系统**: Windows 7/8/10/11
- **依赖**: Visual C++ Redistributable
- **推荐**: 8GB RAM, 1GB可用磁盘空间

## 📞 技术支持

### 问题反馈
1. 保存错误日志
2. 记录操作步骤
3. 提供系统信息

### 获取帮助
- 检查依赖是否完整
- 确保文件路径有效
- 查看日志文件内容

---

**开始使用：**
双击 `softdoc_generator_gui.exe`

**注意：**
首次使用可能需要安装依赖
确保文件路径不要包含特殊字符
"""
    
    try:
        readme_file.write_text(readme_content, encoding='utf-8')
        print(f"✅ 创建使用说明文件: {readme_file.name}")
    except Exception as e:
        print(f"⚠️  创建使用说明失败: {e}")
    
    return True

def create_installer_script():
    """创建安装器脚本"""
    print("\n📦 创建安装器脚本...")
    
    current_dir = Path(__file__).parent
    
    # 创建Windows安装脚本
    install_batch = current_dir / "install_tesseract.bat"
    
    batch_content = '''@echo off
chcp 65001 > nul
echo ==========================================
echo Tesseract OCR 自动安装脚本
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
        install_batch.write_text(batch_content, encoding='gbk')
        print(f"✅ 创建安装脚本: {install_batch.name}")
    except Exception as e:
        print(f"⚠️  创建安装脚本失败: {e}")
    
    return True

def main():
    """主函数"""
    print_banner()
    
    print("本脚本将自动构建软著文档生成工具的exe可执行文件")
    print("构建过程可能需要几分钟，请耐心等待...")
    print()
    
    # 检查前置条件
    if not check_prerequisites():
        print("\n❌ 前置条件检查失败")
        print("请确保:")
        print("   1. PyInstaller 已安装")
        print("   2. src 目录存在")
        print("   3. 所有依赖包已安装")
        input("\n按回车键退出...")
        return 1
    
    print()
    
    # 清理构建目录
    clean_build_directories()
    
    # 创建spec文件
    if not create_spec_file():
        print("\n❌ 创建spec文件失败")
        return 1
    
    # 运行PyInstaller
    if not run_pyinstaller():
        print("\n❌ PyInstaller构建失败")
        return 1
    
    # 验证exe文件
    if not verify_exe_file():
        print("\n⚠️  exe文件验证失败")
    
    # 创建调试信息
    create_debug_info()
    
    # 创建安装器脚本
    create_installer_script()
    
    print("\n" + "=" * 70)
    print("🎉 构建成功完成！")
    print("=" * 70)
    print()
    
    print("📁 输出文件:")
    print("   dist/softdoc_generator_gui.exe - 可执行程序")
    print("   build_info.json - 构建信息")
    print("   README_EXE.md - 使用说明")
    print("   install_tesseract.bat - 安装脚本")
    print()
    
    print("🚀 使用方法:")
    print("   1. 双击 dist/softdoc_generator_gui.exe 运行程序")
    print("   2. 如缺少Tesseract，运行 install_tesseract.bat")
    print("   3. 如需调试，查看 build_info.json")
    print()
    
    print("💡 注意事项:")
    print("   • exe文件为Windows平台专用")
    print("   • 首次运行可能需要安装Visual C++ Redistributable")
    print("   • 确保Tesseract OCR已正确安装")
    print("   • 文件路径不要包含中文或特殊字符")
    print()
    
    print("✅ 构建流程完成，exe文件已生成")
    
    return 0

if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\n构建被用户中断")
        sys.exit(1)
    except Exception as e:
        print(f"\n\n构建过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        input("\n按回车键退出...")
        sys.exit(1)