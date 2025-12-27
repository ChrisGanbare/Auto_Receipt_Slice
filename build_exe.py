"""
打包脚本 - 将项目打包为Windows可执行文件
使用方法：python build_exe.py
"""

import PyInstaller.__main__
import os
import sys

# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))

# 获取main.py的完整路径
main_py_path = os.path.join(current_dir, 'main.py')

# PyInstaller配置
PyInstaller.__main__.run([
    main_py_path,
    '--name=农行电子回单智能拆分工具',
    '--onefile',  # 打包成单个exe文件
    '--windowed',  # 不显示控制台窗口
    '--icon=NONE',  # 如果有图标文件，可以指定路径，如：--icon=icon.ico
    f'--add-data={os.path.join(current_dir, "README.md")};.',  # 包含README文件
    '--hidden-import=tkinter',
    '--hidden-import=tkinter.ttk',
    '--hidden-import=fitz',
    '--hidden-import=pdfplumber',
    '--hidden-import=queue',
    '--hidden-import=operator',
    '--clean',  # 清理临时文件
    '--noconfirm',  # 覆盖输出目录
    f'--distpath={os.path.join(current_dir, "dist")}',  # 输出目录
    f'--workpath={os.path.join(current_dir, "build")}',  # 工作目录
    f'--specpath={os.path.join(current_dir, "build")}',  # spec文件目录
])

print("\n打包完成！")
print(f"可执行文件位置：{os.path.join(current_dir, 'dist', '农行电子回单智能拆分工具.exe')}")

