#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WITRN PD解析器Web版 Nuitka打包脚本
"""

import os
import subprocess
import sys
import shutil

def check_nuitka():
    """检查Nuitka是否安装"""
    try:
        result = subprocess.run([sys.executable, '-m', 'nuitka', '--version'], 
                              capture_output=True, text=True)
        if result.returncode == 0:
            print(f"✓ Nuitka已安装: {result.stdout.strip()}")
            return True
        else:
            print("X Nuitka未安装或无法运行")
            return False
    except Exception as e:
        print(f"X 检查Nuitka时出错: {e}")
        return False

def check_files():
    """检查必要文件是否存在"""
    required_files = [
        'witrn_pd_sniffer_web.py',
        'index.html',
        'requirements.txt'
    ]
    
    missing_files = []
    for file in required_files:
        if not os.path.exists(file):
            missing_files.append(file)
        else:
            print(f"V {file} 存在")
    
    if missing_files:
        print(f"X 缺少文件: {missing_files}")
        return False
    
    return True

def build_with_nuitka():
    """使用Nuitka打包"""
    print("\n开始Nuitka打包...")
    
    # 构建Nuitka命令
    cmd = [
        sys.executable, '-m', 'nuitka',
        '--standalone',                    # 独立模式
        # '--onefile',                       # 单文件模式
        # '--windows-disable-console',       # 禁用控制台窗口
        '--enable-plugin=pywebview',       # 启用pywebview插件
        '--include-data-file=index.html=index.html',  # 包含HTML文件
        '--output-filename=witrn_pd_sniffer_web.exe', # 输出文件名
        '--assume-yes-for-downloads',      # 自动下载依赖
        '--no-prefer-source-code',         # 不优先使用源代码
        '--include-module=codecs',
        '--include-module=encodings',
        'witrn_pd_sniffer_web.py'         # 主文件
    ]
    
    print("执行命令:")
    print(' '.join(cmd))
    print()
    
    try:
        result = subprocess.run(cmd, check=True, text=True)
        print("V 打包成功!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"X 打包失败: {e}")
        return False
    except Exception as e:
        print(f"X 打包出错: {e}")
        return False

def build_with_nuitka_alternative():
    """备用打包方法"""
    print("\n尝试备用打包方法...")
    
    cmd = [
        sys.executable, '-m', 'nuitka',
        '--standalone',
        # '--onefile',
        # '--windows-disable-console',
        '--plugin-enable=pywebview',
        '--include-data-files=index.html=index.html',
        '--include-package-data=pywebview',
        '--include-module=codecs',
        '--include-module=encodings',
        '--assume-yes-for-downloads',
        '--output-dir=dist',
        'witrn_pd_sniffer_web.py'
    ]
    
    print("执行备用命令:")
    print(' '.join(cmd))
    print()
    
    try:
        result = subprocess.run(cmd, check=True, text=True)
        print("V 备用打包成功!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"X 备用打包失败: {e}")
        return False
    except Exception as e:
        print(f"X 备用打包出错: {e}")
        return False

def main():
    """主函数"""
    print("WITRN PD解析器Web版 Nuitka打包工具")
    print("=" * 50)
    
    # 检查环境
    if not check_nuitka():
        print("\n请先安装Nuitka:")
        print("pip install nuitka")
        return 1
    
    if not check_files():
        print("\n请确保所有必要文件都存在")
        return 1
    
    # 尝试打包
    success = False
    
    # 方法1: 标准打包
    if build_with_nuitka():
        success = True
    else:
        # 方法2: 备用打包
        if build_with_nuitka_alternative():
            success = True
    
    if success:
        print("\nV 打包完成!")
        print("可执行文件已生成，可以直接运行")
    else:
        print("\nX 打包失败")
        print("\n可能的解决方案:")
        print("1. 确保pywebview插件可用")
        print("2. 检查HTML文件路径")
        print("3. 尝试手动指定数据文件")
        print("4. 查看Nuitka日志获取详细错误信息")
    
    return 0 if success else 1

if __name__ == "__main__":
    exit(main())
