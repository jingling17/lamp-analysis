import PyInstaller.__main__
import sys
import os

# 获取当前目录
current_dir = os.path.dirname(os.path.abspath(__file__))

PyInstaller.__main__.run([
    'gui.py',  # 主程序文件
    '--name=台灯销售数据分析工具',  # 生成的可执行文件名
    '--windowed',  # 使用GUI模式
    '--onefile',  # 打包成单个文件
    f'--add-data={os.path.join(current_dir, "lamp_analysis.py")}:.',  # 添加依赖文件
    '--clean',  # 清理临时文件
])