# launcher_gui.py — 桌面 GUI 版启动器（无 Streamlit）
import sys
import os

# Pandoc 路径
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

pandoc_dir = os.path.join(BASE_DIR, 'pandoc')
os.environ['PATH'] = pandoc_dir + os.pathsep + os.environ.get('PATH', '')
os.environ['PYPANDOC_PANDOC'] = os.path.join(pandoc_dir, 'pandoc.exe')

# 直接启动 GUI
from app_gui import main
main()
