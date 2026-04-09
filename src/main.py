"""
银行流水核对系统 - 主入口
"""
import sys
import os

# 将 src 目录加入 Python 路径
if getattr(sys, 'frozen', False):
    # PyInstaller 打包模式
    base_path = sys._MEIPASS
    src_path = os.path.join(base_path, 'src')
else:
    base_path = os.path.dirname(os.path.abspath(__file__))
    src_path = os.path.join(base_path, 'src')

sys.path.insert(0, src_path)

import webview


def main():
    from gui.api import PyApi

    api = PyApi()

    # 获取 HTML 资源路径
    if getattr(sys, 'frozen', False):
        # 打包模式：HTML 在 src/gui/assets/index.html
        html_path = os.path.join(src_path, 'gui', 'assets', 'index.html')
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
