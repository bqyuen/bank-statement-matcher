"""
GUI 后端接口（PyApi）

提供GUI层调用的Python API：
- selectFolder  / selectFile：弹出系统文件/文件夹选择对话框
- readBankFile  / readLedgerFile：读取银行流水 / 三栏账文件
- runMatching：执行匹配
- genExcelReport：生成Excel核对报告
- openFile：用系统默认程序打开文件
"""

import os
import sys
import time
from typing import List, Optional, Dict, Any

# 文件选择对话框（跨平台）
if sys.platform == "darwin":
    import subprocess

    def _select_file(title: str = "选择文件", file_types: str = "") -> Optional[str]:
        """macOS 使用 osascript 弹出文件选择对话框"""
        script = f'''
        tell application "System Events"
            activate
            set chosenFile to (choose file with prompt "{title}" {file_types})
        end tell
        chosenFile
        '''
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=30
            )
            path = result.stdout.strip()
            return path if path and path != "" else None
        except Exception:
            return None

    def _select_folder(title: str = "选择文件夹") -> Optional[str]:
        """macOS 使用 osascript 弹出文件夹选择对话框"""
        script = f'''
        tell application "System Events"
            activate
            set chosenFolder to (choose folder with prompt "{title}")
        end tell
        chosenFolder
        '''
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=30
            )
            path = result.stdout.strip()
            return path if path and path != "" else None
        except Exception:
            return None

elif sys.platform == "win32":
    import ctypes
    from ctypes import wintypes, windll

    OFN_EXPLORER = 0x00080000
    OFN_FILEMUSTEXIST = 0x00001000
    OFN_HIDEREADONLY = 0x00000004

    class OPENFILENAME(ctypes.Structure):
        _fields_ = [
            ('lStructSize', wintypes.DWORD),
            ('hwndOwner', wintypes.HWND),
            ('hInstance', wintypes.HINSTANCE),
            ('lpstrFilter', wintypes.LPCWSTR),
            ('lpstrCustomFilter', wintypes.LPWSTR),
            ('nMaxCustFilter', wintypes.DWORD),
            ('nFilterIndex', wintypes.DWORD),
            ('lpstrFile', wintypes.LPWSTR),
            ('nMaxFile', wintypes.DWORD),
            ('lpstrFileTitle', wintypes.LPWSTR),
            ('nMaxFileTitle', wintypes.DWORD),
            ('lpstrInitialDir', wintypes.LPCWSTR),
            ('lpstrTitle', wintypes.LPCWSTR),
            ('flags', wintypes.DWORD),
            ('nFileOffset', wintypes.WORD),
            ('nFileExtension', wintypes.WORD),
            ('lpstrDefExt', wintypes.LPCWSTR),
            ('lCustData', wintypes.LPARAM),
            ('lpfnHook', wintypes.LPVOID),
            ('lpTemplateName', wintypes.LPCWSTR),
        ]

    _GetOpenFileNameW = ctypes.windll.comdlg32.GetOpenFileNameW
    _GetOpenFileNameW.argtypes = [ctypes.POINTER(OPENFILENAME)]
    _GetOpenFileNameW.restype = wintypes.BOOL

    _GetSaveFileNameW = ctypes.windll.comdlg32.GetSaveFileNameW
    _GetSaveFileNameW.argtypes = [ctypes.POINTER(OPENFILENAME)]
    _GetSaveFileNameW.restype = wintypes.BOOL

    def _select_file(title: str = "选择文件", file_types: str = "") -> Optional[str]:
        title16 = title + '\x00'
        filter_str = 'All Files\x00*.*\x00\x00'
        buf = ctypes.create_unicode_buffer(8192)
        ofn = OPENFILENAME()
        ofn.lStructSize = ctypes.sizeof(OPENFILENAME)
        ofn.lpstrFilter = filter_str
        ofn.lpstrFile = ctypes.cast(buf, ctypes.c_wchar_p)
        ofn.nMaxFile = 8192
        ofn.lpstrTitle = title16
        ofn.flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY
        if _GetOpenFileNameW(ctypes.byref(ofn)):
            return buf.value
        return None

    def _select_folder(title: str = "选择文件夹") -> Optional[str]:
        # 使用 Windows Shell API SHBrowseForFolder
        try:
            from ctypes import POINTER, c_void_p
            from ctypes.wintypes import UINT, HWND, LPCWSTR, LPWSTR

            BIF_RETURNONLYFSDIRS = 0x0001
            BIF_NEWDIALOGSTYLE = 0x0040
            BIF_EDITBOX = 0x0010
            BIF_SHAREABLE = 0x8000

            class BROWSEINFO(ctypes.Structure):
                _fields_ = [
                    ('hwndOwner', HWND),
                    ('pidlRoot', c_void_p),
                    ('pszDisplayName', LPWSTR),
                    ('lpszTitle', LPCWSTR),
                    ('ulFlags', UINT),
                    ('lpfn', c_void_p),
                    ('lParam', c_void_p),
                    ('iImage', ctypes.c_int),
                ]

            buf = ctypes.create_unicode_buffer(8192)
            bi = BROWSEINFO()
            bi.pszDisplayName = buf
            bi.lpszTitle = title
            bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE | BIF_EDITBOX

            pidl = windll.shell32.SHBrowseForFolderW(ctypes.byref(bi))
            if pidl:
                path_buf = ctypes.create_unicode_buffer(8192)
                if windll.shell32.SHGetPathFromIDListW(pidl, path_buf):
                    return path_buf.value
        except Exception:
            pass
        return None
else:
    # Linux 等平台使用 tkinter
    try:
        import tkinter as tk
        from tkinter import filedialog
        _root = tk.Tk()
        _root.withdraw()
        def _select_file(title: str = "选择文件", file_types: str = "") -> Optional[str]:
            path = filedialog.askopenfilename(title=title)
            return path if path else None
        def _select_folder(title: str = "选择文件夹") -> Optional[str]:
            path = filedialog.askdirectory(title=title)
            return path if path else None
    except Exception:
        def _select_file(*_): return None
        def _select_folder(*_): return None


# ---------------------------------------------------------------------------
# PyApi
# ---------------------------------------------------------------------------

class PyApi:
    """GUI 后端接口"""

    def __init__(self):
        # 数据存储
        self.bank_txns: List = []
        self.ledger_entries: List = []
        self.matched: List = []
        self.bank_only: List = []
        self.ledger_only: List = []

        # 匹配参数
        self.tolerance_near: int = 3
        self.tolerance_far: int = 7

        # 输出目录
        ts = time.strftime('%Y年%m月%d日%H时%M分%S')
        if sys.platform == "win32":
            self.out_base_path = os.path.expanduser('~') + f'\\Documents\\银行流水核对\\out\\{ts}'
        else:
            self.out_base_path = os.path.expanduser('~') + f'/Documents/银行流水核对/out/{ts}'
        os.makedirs(self.out_base_path, exist_ok=True)

        # 文件路径记录
        self._bank_file_path: str = ""
        self._ledger_file_path: str = ""

    # -----------------------------------------------------------------------
    # 文件选择
    # -----------------------------------------------------------------------

    def selectFolder(self) -> Optional[str]:
        """弹出文件夹选择对话框，返回选中路径"""
        return _select_folder("选择文件夹")

    def selectFile(self) -> Optional[str]:
        """弹出文件选择对话框，返回选中路径"""
        return _select_file("选择文件")

    # -----------------------------------------------------------------------
    # 数据读取
    # -----------------------------------------------------------------------

    def readBankFile(self, file_path: str = "") -> Dict[str, Any]:
        """
        读取银行流水文件（支持 OFD / PDF / Excel / CSV），自动识别格式

        Args:
            file_path: 文件路径，为空则弹出文件选择框

        Returns:
            {'count': N} 或 {'error': '错误信息'}
        """
        if not file_path:
            file_path = _select_file("选择银行流水文件",
                                     "ofd,pdf,xlsx,xlsm,csv,PDF,OFD,XLSX,CSV")
            if not file_path:
                return {"error": "未选择文件"}

        try:
            from bank_parsers import detect_and_parse
            self.bank_txns = detect_and_parse(file_path)
            self._bank_file_path = file_path
            return {"count": len(self.bank_txns)}
        except Exception as e:
            return {"error": str(e)}

    def readLedgerFile(self, file_path: str = "") -> Dict[str, Any]:
        """
        读取三栏账 Excel 文件

        Args:
            file_path: Excel文件路径，为空则弹出文件选择框

        Returns:
            {'count': N} 或 {'error': '错误信息'}
        """
        if not file_path:
            file_path = _select_file("选择三栏账文件",
                                     "xlsx,XLSX")
            if not file_path:
                return {"error": "未选择文件"}

        try:
            from ledger.ledger_parser import LedgerParser
            parser = LedgerParser()
            self.ledger_entries = parser.parse(file_path)
            self._ledger_file_path = file_path
            return {"count": len(self.ledger_entries)}
        except Exception as e:
            return {"error": str(e)}

    # -----------------------------------------------------------------------
    # 匹配
    # -----------------------------------------------------------------------

    def runMatching(
        self,
        tolerance_near: int = 3,
        tolerance_far: int = 7
    ) -> Dict[str, Any]:
        """
        执行银行流水与三栏账的匹配

        Args:
            tolerance_near: 近端日期容差（天）
            tolerance_far:  远端日期容差（天）

        Returns:
            {'matched': N, 'bank_only': N, 'ledger_only': N}
        """
        if not self.bank_txns:
            return {"error": "请先读取银行流水文件"}
        if not self.ledger_entries:
            return {"error": "请先读取三栏账文件"}

        self.tolerance_near = tolerance_near
        self.tolerance_far = tolerance_far

        try:
            from engine.matcher import BankLedgerMatcher
            matcher = BankLedgerMatcher(tolerance_near, tolerance_far)
            self.matched, self.bank_only, self.ledger_only = matcher.match(
                self.bank_txns, self.ledger_entries
            )
            return {
                "matched": len(self.matched),
                "bank_only": len(self.bank_only),
                "ledger_only": len(self.ledger_only)
            }
        except Exception as e:
            return {"error": str(e)}

    # -----------------------------------------------------------------------
    # Excel 报告
    # -----------------------------------------------------------------------

    def genExcelReport(
        self,
        bank_balance: float = 0.0,
        ledger_balance: float = 0.0,
        company_name: str = "",
        account_no: str = "",
        statement_date: str = ""
    ) -> Dict[str, Any]:
        """
        生成 Excel 核对报告

        Args:
            bank_balance:   银行对账单余额
            ledger_balance: 企业日记账余额
            company_name:   编制单位名称
            account_no:     账户号
            statement_date: 截止日期

        Returns:
            {'path': '...'} 或 {'error': '...'}
        """
        if not self.bank_txns and not self.ledger_entries:
            return {"error": "没有可输出的数据，请先读取文件并执行匹配"}

        try:
            from output.excel_writer import ExcelReportWriter

            if sys.platform == "win32":
                out_path = os.path.join(self.out_base_path, "核对报告.xlsx")
            else:
                out_path = os.path.join(self.out_base_path, "核对报告.xlsx")

            writer = ExcelReportWriter(out_path)
            writer.write_report(
                bank_txns=self.bank_txns,
                ledger_entries=self.ledger_entries,
                matched=self.matched,
                bank_only=self.bank_only,
                ledger_only=self.ledger_only,
                bank_balance=bank_balance,
                ledger_balance=ledger_balance,
                company_name=company_name,
                account_no=account_no,
                statement_date=statement_date
            )
            return {"path": out_path}
        except Exception as e:
            return {"error": str(e)}

    # -----------------------------------------------------------------------
    # 文件操作
    # -----------------------------------------------------------------------

    def openFile(self, file_path: str = "") -> Dict[str, Any]:
        """
        用系统默认程序打开文件

        Args:
            file_path: 文件路径，为空则打开输出目录

        Returns:
            {'path': ...}
        """
        if not file_path:
            file_path = self.out_base_path

        try:
            if sys.platform == "darwin":
                import subprocess
                subprocess.run(["open", file_path], check=True)
            elif sys.platform == "win32":
                os.startfile(file_path)
            else:
                subprocess.run(["xdg-open", file_path], check=True)
            return {"path": file_path}
        except Exception as e:
            return {"error": str(e)}

    # -----------------------------------------------------------------------
    # 状态查询
    # -----------------------------------------------------------------------

    def getStatus(self) -> Dict[str, Any]:
        """获取当前处理状态（供GUI轮询）"""
        return {
            "bank_count": len(self.bank_txns),
            "ledger_count": len(self.ledger_entries),
            "matched_count": len(self.matched),
            "bank_only_count": len(self.bank_only),
            "ledger_only_count": len(self.ledger_only),
            "out_base_path": self.out_base_path,
        }
