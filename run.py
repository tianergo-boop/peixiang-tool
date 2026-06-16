"""
配箱工具 - 直接运行入口（用于PyInstaller打包）
"""

import os
import sys

# 设置项目路径
project_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, project_dir)

from src.gui import PeiXiangApp

if __name__ == '__main__':
    app = PeiXiangApp()
    app.run()
