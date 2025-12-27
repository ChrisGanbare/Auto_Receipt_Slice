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

# 检查图标文件是否存在
icon_path = os.path.join(current_dir, 'icon.ico')
icon_param = f'--icon={icon_path}' if os.path.exists(icon_path) else '--icon=NONE'

# 检查版本信息文件是否存在
version_file_path = os.path.join(current_dir, 'version_info.txt')
version_param = [f'--version-file={version_file_path}'] if os.path.exists(version_file_path) else []

# PyInstaller配置
build_params = [
    main_py_path,
    '--name=农行电子回单智能拆分工具',
    '--onefile',  # 打包成单个exe文件
    '--windowed',  # 不显示控制台窗口
    icon_param,  # 图标文件（如果存在）
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
]

# 如果版本信息文件存在，添加版本信息参数
if version_param:
    build_params.extend(version_param)

PyInstaller.__main__.run(build_params)

print("\n打包完成！")
print(f"可执行文件位置：{os.path.join(current_dir, 'dist', '农行电子回单智能拆分工具.exe')}")

