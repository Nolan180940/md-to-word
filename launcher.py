# launcher.py — Streamlit + Pandoc 启动器，用于 PyInstaller 打包
import sys
import os
import webbrowser
import threading
import time

# ------------------------------------------------------------
# 1. 处理 PyInstaller 打包后的路径（解压目录 vs 开发目录）
# ------------------------------------------------------------
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS  # PyInstaller 临时解压目录
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ------------------------------------------------------------
# 2. 将捆绑的 pandoc.exe 加入 PATH（让 pypandoc 能找到）
# ------------------------------------------------------------
pandoc_dir = os.path.join(BASE_DIR, 'pandoc')
os.environ['PATH'] = pandoc_dir + os.pathsep + os.environ.get('PATH', '')
os.environ['PYPANDOC_PANDOC'] = os.path.join(pandoc_dir, 'pandoc.exe')

# ------------------------------------------------------------
# 3. 确保 app.py 的路径正确
# ------------------------------------------------------------
app_path = os.path.join(BASE_DIR, 'app.py')
if not os.path.exists(app_path):
    # 兼容情况：app.py 可能在当前工作目录
    app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'app.py')

# ------------------------------------------------------------
# 4. 自动打开浏览器
# ------------------------------------------------------------
PORT = 8501

def open_browser():
    time.sleep(2)  # 等 Streamlit 启动
    webbrowser.open(f'http://localhost:{PORT}')

threading.Thread(target=open_browser, daemon=True).start()

# ------------------------------------------------------------
# 5. 以编程方式启动 Streamlit
# ------------------------------------------------------------
import streamlit.web.cli as stcli

sys.argv = [
    'streamlit', 'run', app_path,
    '--server.port', str(PORT),
    '--server.headless', 'true',
    '--browser.serverAddress', 'localhost',
    '--server.fileWatcherType', 'none',  # 打包后无需监听文件变化
]
stcli.main()
