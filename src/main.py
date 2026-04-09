"""
银行流水核对系统 - 主入口
"""
import sys
import os

# 将 src 目录加入 Python 路径
if getattr(sys, 'frozen', False):
    # PyInstaller onefile 打包模式
    base_path = sys._MEIPASS
    # onefile exe 解压后，源码在 base_path 根目录（无 src/ 子目录）
    src_path = base_path
else:
    base_path = os.path.dirname(os.path.abspath(__file__))
    src_path = os.path.join(base_path, 'src')

# src/ 目录在 base_path/src/ 下
sys.path.insert(0, os.path.join(base_path, 'src'))

import webview


def main():
    from gui.api import PyApi

    api = PyApi()

    # 获取 HTML 资源路径
    if getattr(sys, 'frozen', False):
        # 打包模式（onedir exe）：HTML 在 _internal/gui/assets/index.html
        html_path = os.path.join(base_path, '_internal', 'gui', 'assets', 'index.html')
    else:
        # 开发模式
        html_path = os.path.join(src_path, 'gui', 'assets', 'index.html')

    window = webview.create_window(
        title='银行流水与三栏账自动核对系统',
        url=html_path,
        width=1200,
        height=800,
        resizable=True,
        js_api=api
    )
    webview.start(debug=False)


if __name__ == '__main__':
    main()
