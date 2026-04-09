"""
银行流水核对系统 - 主入口
Onefile mode: sys._MEIPASS is the temp extraction root, exe is at sys.executable
"""
import sys, os

if getattr(sys, 'frozen', False):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
    src_path = base_path  # onefile: src/ is at extraction root
else:
    base_path = os.path.dirname(os.path.abspath(__file__))
    src_path = os.path.join(base_path, 'src')

sys.path.insert(0, src_path)
import webview


def main():
    from gui.api import PyApi
    api = PyApi()

    if getattr(sys, 'frozen', False):
        html_path = os.path.join(base_path, 'gui', 'assets', 'index.html')
    else:
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
