"""
银行流水解析器包
"""
from abc import ABC, abstractmethod
from typing import List
from dataclasses import dataclass
from datetime import datetime


@dataclass
class BankTransaction:
    file: str
    bank_type: str
    date: datetime
    amount: float
    direction: str
    counterparty: str
    counterparty_account: str
    counterparty_bank: str
    summary: str
    serial_no: str
    balance: float
    mate: bool = False
    matched_voucher: str = None
    match_level: str = None


class BaseBankParser(ABC):
    @abstractmethod
    def detect(self, file_path: str) -> bool:
        pass

    @abstractmethod
    def parse(self, file_path: str) -> List[BankTransaction]:
        pass


# Registry - module-level variable (NOT a class variable)
_PARSER_REGISTRY = {}


def register_parser(bank_type: str, parser_class: type, priority: int = 100):
    """注册解析器到全局注册表"""
    if bank_type not in _PARSER_REGISTRY:
        _PARSER_REGISTRY[bank_type] = []
    _PARSER_REGISTRY[bank_type].append((priority, parser_class))
    _PARSER_REGISTRY[bank_type].sort(key=lambda x: x[0])


def get_parsers(bank_type: str) -> List:
    return _PARSER_REGISTRY.get(bank_type, [])


def _iter_all_parsers() -> List:
    """按优先级返回全量解析器类（去重）"""
    all_items = []
    for bank, parser_items in _PARSER_REGISTRY.items():
        for priority, parser_cls in parser_items:
            all_items.append((priority, bank, parser_cls))
    all_items.sort(key=lambda x: x[0])

    seen = set()
    result = []
    for priority, bank, parser_cls in all_items:
        key = (bank, parser_cls)
        if key in seen:
            continue
        seen.add(key)
        result.append((priority, bank, parser_cls))
    return result


def detect_bank(file_path: str) -> str:
    """自动识别银行类型"""
    import os
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ('.ofd', '.zip'):
        return _detect_bank_from_ofd_or_zip(file_path)
    elif ext == '.pdf':
        return _detect_bank_from_pdf(file_path)
    elif ext in ('.xlsx', '.xlsm', '.csv'):
        return _detect_bank_from_sheet(file_path)
    return 'unknown'


def _detect_bank_from_ofd_or_zip(file_path: str) -> str:
    import zipfile
    import os

    # 方式1：尝试作为 ZIP 包裹的 OFD 处理
    try:
        with zipfile.ZipFile(file_path, 'r') as zf:
            names = zf.namelist()
            xbrl_files = [n for n in names if n.endswith('.xml')]
            for xf in xbrl_files:
                try:
                    content = zf.read(xf).decode('utf-8', errors='ignore')
                    if 'xbrl.mof.gov.cn/taxonomy/2023-05-15/bkrs' in content or 'xbrl.mof.gov.cn/taxonomy/2021-11-30/bkrs' in content:
                        return 'ccb'
                except Exception:
                    continue
            all_content = ''.join(zf.read(n).decode('utf-8', errors='ignore') for n in names[:50])
            if '建设银行' in all_content or 'CCB' in all_content:
                return 'ccb'
            if '招商银行' in all_content:
                return 'cmb'
            if '农业银行' in all_content or 'ABC' in all_content:
                return 'abc'
    except Exception:
        pass

    # 方式2：直接读取 XML 内容（支持独立 XBRL 文件或 OFD 入口文件）
    try:
        abs_path = os.path.abspath(file_path)
        with open(abs_path, 'r', encoding='utf-8', errors='ignore') as f:
            xml_content = f.read(16384)  # 读前16KB，包含命名空间声明

        # 是 OFD 入口文件 → 查找 Doc_0/Attachs/*.xml 获取 bank type
        if '<ofd:OFD' in xml_content or '<ofd:' in xml_content:
            ofd_dir = os.path.dirname(abs_path)
            # Doc_0 通常与 OFD.xml 同级
            for parent in [ofd_dir, os.path.dirname(ofd_dir)]:
                if not parent or not os.path.isdir(parent):
                    continue
                doc_dir = os.path.join(parent, 'Doc_0')
                if not os.path.isdir(doc_dir):
                    continue
                attachs_dir = os.path.join(doc_dir, 'Attachs')
                if not os.path.isdir(attachs_dir):
                    continue
                for fn in os.listdir(attachs_dir):
                    if fn.endswith('.xml') and (fn.startswith('bkrs') or fn.startswith('BKR')):
                        bkrs_path = os.path.join(attachs_dir, fn)
                        with open(bkrs_path, 'r', encoding='utf-8', errors='ignore') as f:
                            bkrs_content = f.read(32768)
                        if 'xbrl.mof.gov.cn/taxonomy/2023-05-15/bkrs' in bkrs_content or 'xbrl.mof.gov.cn/taxonomy/2021-11-30/bkrs' in bkrs_content:
                            return 'ccb'
                        if '农业银行' in bkrs_content:
                            return 'abc'
                        if '招商银行' in bkrs_content:
                            return 'cmb'
                break

        # 是普通 XML 文件（含 XBRL bank report）→ 直接检测银行类型
        if xml_content.strip().startswith('<?xml') or '<xbrl' in xml_content:
            if 'xbrl.mof.gov.cn/taxonomy/2023-05-15/bkrs' in xml_content or 'xbrl.mof.gov.cn/taxonomy/2021-11-30/bkrs' in xml_content:
                return 'ccb'
            if '农业银行' in xml_content or 'ABC' in xml_content:
                return 'abc'
            if '招商银行' in xml_content:
                return 'cmb'
            if '建设银行' in xml_content or 'CCB' in xml_content:
                return 'ccb'

    except Exception:
        pass

    return 'unknown'


def _detect_bank_from_pdf(file_path: str) -> str:
    import pymupdf
    try:
        with pymupdf.open(file_path) as doc:
            scanned_like = True
            checked_pages = 0
            for page in doc[:3]:
                checked_pages += 1
                text = page.get_text()
                if not text.strip():
                    # 纯图片页按扫描件候选处理
                    if len(page.get_images(full=True)) != 1:
                        scanned_like = False
                    continue

                if '账务明细清单' in text and '招商银行' in text:
                    return 'cmb'
                if any(k in text for k in [
                    '中国建设银行账户明细信息', '建设银行', 'CCB',
                    '综合柜员系统流水账', '中国建设银行'
                ]):
                    return 'ccb'
                if '中交财务有限公司' in text:
                    return 'ccfc'

                # 有较多可提取文本通常不是纯扫描件
                if len(text.strip()) > 50:
                    scanned_like = False
                if len(page.get_images(full=True)) != 1:
                    scanned_like = False

            # 兜底：扫描型PDF没有可提取文本，但仍可能是建行账单
            if checked_pages > 0 and scanned_like:
                return 'ccb'
    except Exception:
        pass
    return 'unknown'


def _detect_bank_from_sheet(file_path: str) -> str:
    """从 Excel/CSV 文本特征识别银行类型"""
    import os
    ext = os.path.splitext(file_path)[1].lower()

    headers = []
    context_text = ''

    try:
        if ext == '.csv':
            import csv
            for enc in ('utf-8-sig', 'gbk', 'gb18030'):
                try:
                    with open(file_path, 'r', encoding=enc, newline='') as f:
                        reader = csv.reader(f)
                        rows = []
                        for i, row in enumerate(reader, start=1):
                            rows.append(row)
                            if i >= 50:
                                break
                    break
                except Exception:
                    rows = []
            if not rows:
                return 'unknown'
        else:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
            try:
                ws = wb.active
                rows = []
                for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
                    rows.append([cell for cell in row])
                    if i >= 50:
                        break
            finally:
                wb.close()

        for row in rows[:30]:
            row_headers = [str(c).replace(' ', '').strip() if c is not None else '' for c in row]
            joined = ''.join(row_headers)
            if ('交易金额' in joined or '交易发生额' in joined or '发生额' in joined) and (
                '交易日期' in joined or '记账日期' in joined or '入账日期' in joined or '日期' in joined
            ):
                headers = row_headers
                break

        context_text = ' '.join(
            str(c).strip().lower()
            for row in rows[:30]
            for c in row[:15]
            if c is not None and str(c).strip()
        )
    except Exception:
        return 'unknown'

    has_abc_mark = any(
        k in context_text
        for k in ('农业银行', '中国农业银行', 'agricultural bank', ' abc ', 'abc银行')
    )
    has_trade_amount_col = any('交易金额' in h for h in headers)

    # 有交易日期+交易金额列时，按农业银行表格流水走解析器；
    # 银行关键词作为增强条件，避免仅凭版头识别失败。
    if headers and (has_abc_mark or has_trade_amount_col):
        return 'abc'
    return 'unknown'


def detect_and_parse(file_path: str) -> List[BankTransaction]:
    """自动检测银行类型，分发到对应解析器"""
    bank_type = detect_bank(file_path)
    tried = []
    parse_errors = []

    # 第一轮：按识别到的银行类型优先尝试
    if bank_type != 'unknown':
        parsers = get_parsers(bank_type)
        for priority, parser_class in parsers:
            parser = parser_class()
            tried.append((bank_type, parser_class.__name__))
            try:
                if not parser.detect(file_path):
                    continue
                data = parser.parse(file_path)
                if data:
                    return data
                parse_errors.append(f"{parser_class.__name__}: 解析结果为空")
            except Exception as e:
                parse_errors.append(f"{parser_class.__name__}: {e}")

    # 第二轮：回退全量解析器，防止银行识别误判导致错过正确规则
    for priority, fallback_bank, parser_class in _iter_all_parsers():
        if (fallback_bank, parser_class.__name__) in tried:
            continue
        parser = parser_class()
        tried.append((fallback_bank, parser_class.__name__))
        try:
            if not parser.detect(file_path):
                continue
            data = parser.parse(file_path)
            if data:
                return data
            parse_errors.append(f"{parser_class.__name__}: 解析结果为空")
        except Exception as e:
            parse_errors.append(f"{parser_class.__name__}: {e}")

    if bank_type == 'unknown':
        detail = f"无法识别该文件类型: {file_path}"
    else:
        detail = f"银行 {bank_type} 已识别，但未提取到有效交易记录"

    if parse_errors:
        raise ValueError(detail + "；" + " | ".join(parse_errors[:3]))
    raise ValueError(detail)


def get_supported_types() -> List[str]:
    return list(_PARSER_REGISTRY.keys())


# ─── 导入并注册解析器 ─────────────────────────────────────────
from .ccb.ccb_pdf import CCB_PDF_Parser
from .ccb.ccb_ofd_2023 import CCB_OFD_2023_Parser
from .ccb.ccb_ofd_2021 import CCB_OFD_2021_Parser
from .ccb.ccb_image_pdf import CCB_Image_PDF_Parser
from .cmb.cmb_pdf import CMB_PDF_Parser
from .abc.abc_excel import ABC_Excel_Parser
from .ccfc.ccfc_pdf import CCFC_PDF_Parser

register_parser('ccb', CCB_PDF_Parser, priority=10)
register_parser('ccb', CCB_OFD_2023_Parser, priority=20)
register_parser('ccb', CCB_OFD_2021_Parser, priority=30)
register_parser('ccb', CCB_Image_PDF_Parser, priority=90)
register_parser('cmb', CMB_PDF_Parser, priority=10)
register_parser('abc', ABC_Excel_Parser, priority=10)
register_parser('ccfc', CCFC_PDF_Parser, priority=10)

__all__ = [
    'BankTransaction',
    'BaseBankParser',
    'detect_bank',
    'detect_and_parse',
    'get_supported_types',
    'get_parsers',
    'register_parser',
]
