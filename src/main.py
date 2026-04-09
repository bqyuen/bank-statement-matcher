"""
银行流水核对系统 - 主入口
"""
import sys
import os

# 将 src 目录加入 Python 路径
if getattr(sys, 'frozen', False):
    # PyInstaller onedir 打包模式
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    internal_dir = os.path.join(exe_dir, '_internal')
    base_path = exe_dir
    # onedir: pylib/base_library 在 _internal/ 下
    src_path = internal_dir
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
        # onedir 模式：exe 和 _internal 同目录，HTML 在 _internal/gui/assets/index.html
        # 用 sys.executable 获取 exe 所在目录（不是 _MEIPASS 临时目录）
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        html_path = os.path.join(exe_dir, '_internal', 'gui', 'assets', 'index.html')
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
