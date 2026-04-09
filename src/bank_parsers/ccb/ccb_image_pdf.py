"""
建设银行 图片型PDF解析器（OCR）

解析 CCB 图片型 PDF（扫描件）使用 OCR 识别
识别特征：PDF每页text=0且恰好1张图片
"""
import re
from typing import List, Optional
import pymupdf

from ..base import BaseBankParser, BankTransaction
from ..comm.amount import normalize_amount
from ..comm.date import normalize_date
from .ccb_pdf import _clean_text, CCB_PDF_Parser


class CCB_Image_PDF_Parser(BaseBankParser):
    """CCB 图片型 PDF 解析器（OCR）"""

    BANK_CODE = 'ccb'

    def detect(self, file_path: str) -> bool:
        """
        检测是否为图片型 PDF（需要 OCR 处理）

        判断条件：PDF每页text=0且恰好1张图片
        """
        import os
        ext = os.path.splitext(file_path)[1].lower()
        if ext != '.pdf':
            return False

        try:
            with pymupdf.open(file_path) as doc:
                for page in doc:
                    text = page.get_text().strip()
                    # 检查文字内容是否为0或极少
                    if text and len(text) > 50:
                        # 有足够文字，不是纯图片型
                        return False
                    # 检查图片数量
                    image_list = page.get_images(full=True)
                    if len(image_list) != 1:
                        return False
                # 所有页都通过检测
                return True
        except Exception:
            return False

    def parse(self, file_path: str) -> List[BankTransaction]:
        """
        使用 OCR 解析图片型 CCB PDF

        Args:
            file_path: PDF 文件路径

        Returns:
            BankTransaction 列表
        """
        # 优先 pytesseract，回退 RapidOCR
        ocr_engine = None
        try:
            import pytesseract
            ocr_engine = 'tesseract'
        except ImportError:
            try:
                from rapidocr_onnxruntime import RapidOCR as RapidOCREngine
                ocr_engine = 'rapidocr'
            except ImportError:
                raise ImportError(
                    "图片型PDF需要OCR支持，但系统中未安装 pytesseract 或 rapidocr。"
                    "解决方案:\n"
                    "1. Windows: 下载 Tesseract OCR 安装包 https://github.com/UB-Mannheim/tesseract/wiki\n"
                    "2. Python: pip install pytesseract\n"
                    "3. 或安装 rapidocr_onnxruntime (推荐，无需额外安装软件): pip install rapidocr_onnxruntime"
                )

        results: List[BankTransaction] = []
        _rapid_ocr = None

        def _do_ocr(img_bytes: bytes) -> str:
            """使用可用的OCR引擎提取文字"""
            if ocr_engine == 'tesseract':
                try:
                    return pytesseract.image_to_string(img_bytes, lang='chi_sim', config='--psm 6')
                except Exception:
                    return pytesseract.image_to_string(img_bytes, config='--psm 6')
            else:
                # RapidOCR
                nonlocal _rapid_ocr
                if _rapid_ocr is None:
                    _rapid_ocr = RapidOCREngine()
                result, _, _ = _rapid_ocr(img_bytes)
                if not result:
                    return ''
                # RapidOCR returns list of [text, box, score]
                lines = [item[0] for item in result]
                return '\n'.join(lines)

        try:
            with pymupdf.open(file_path) as doc:
                for page_num, page in enumerate(doc):
                    # 渲染页面为图片
                    # 使用较高分辨率以提高 OCR 准确率
                    mat = pymupdf.Matrix(3, 3)  # 3x缩放
                    pix = page.get_pixmap(matrix=mat)
                    img_data = pix.tobytes('png')

                    # 使用可用OCR引擎进行识别
                    ocr_text = _do_ocr(img_data)

                    # 解析 OCR 文本
                    page_txns = self._parse_ocr_text(ocr_text, file_path, page_num + 1)
                    results.extend(page_txns)

        except Exception:
            pass

        return results

    def _parse_ocr_text(self, ocr_text: str, file_path: str, page_num: int) -> List[BankTransaction]:
        """
        解析 OCR 识别的文本

        使用与 CCB_PDF_Parser 相似的解析逻辑
        """
        transactions: List[BankTransaction] = []

        # OCR 文本可能包含换行符，需要先清洗
        ocr_clean = _clean_text(ocr_text)

        # 尝试按行分割并识别表格结构
        lines = ocr_text.split('\n')
        lines = [l.strip() for l in lines if l.strip()]

        # 查找包含"中国建设银行"的标题行以确认格式
        has_header = any('中国建设银行' in l for l in lines)
        if not has_header:
            # 可能是其他格式的扫描件，跳过
            return transactions

        # 查找表头行以确定列位置
        header_line_idx = -1
        for i, line in enumerate(lines):
            if '交易时间' in line or '记账日期' in line or '借方' in line or '贷方' in line:
                header_line_idx = i
                break

        if header_line_idx == -1:
            # 没找到表头，尝试整体解析
            return self._parse_ocr_raw(lines, file_path, page_num)

        # 从表头行之后开始解析数据行
        data_lines = lines[header_line_idx + 1:]

        # 尝试识别表格格式
        # CCB 格式每行数据包含：日期 借方 贷方 余额 对方户名 摘要 等
        for line in data_lines:
            txn = self._parse_ocr_line(line, file_path)
            if txn:
                transactions.append(txn)

        return transactions

    def _parse_ocr_raw(self, lines: List[str], file_path: str, page_num: int) -> List[BankTransaction]:
        """直接解析 OCR 行（无表头参考）"""
        transactions: List[BankTransaction] = []

        for line in lines:
            # 跳过过短的行
            if len(line) < 10:
                continue

            txn = self._parse_ocr_line(line, file_path)
            if txn:
                transactions.append(txn)

        return transactions

    def _parse_ocr_line(self, line: str, file_path: str) -> Optional[BankTransaction]:
        """
        解析单行 OCR 文本

        OCR 格式与 CCB PDF 相似，但可能有噪声
        """
        # 清理行文本
        line = _clean_text(line)

        # 查找日期（YYYYMMDD 或 YYYY-MM-DD）
        date_patterns = [
            r'(\d{4}[-/年]\d{1,2}[-/月]\d{1,2}[日]?)',
            r'(\d{8})',
        ]
        date_str = None
        date_pos = -1
        for pat in date_patterns:
            m = re.search(pat, line)
            if m:
                date_str = m.group(1)
                date_pos = m.start()
                break

        if not date_str:
            return None

        date = normalize_date(date_str)
        if date is None:
            # 尝试 YYYYMMDD
            date_raw = re.search(r'(\d{8})', line)
            if date_raw:
                date = normalize_date(date_raw.group(1), '%Y%m%d')

        if date is None:
            return None

        # 在日期之后的部分查找金额
        after_date = line[date_pos + len(date_str):]

        # 查找金额（可能带逗号）
        amount_pattern = re.compile(r'([\d,]+\.?\d*)')
        amounts = list(amount_pattern.finditer(after_date))

        if not amounts:
            return None

        # 尝试找到借方和贷方金额
        debit_amount = 0.0
        credit_amount = 0.0
        balance = 0.0

        # OCR 可能把多个金额串在一起，需要智能识别
        # 通常格式：借方金额 贷方金额 余额（都是正数）
        # 在 OCR 中，可能借方金额和贷方金额是分开的字段

        # 简化处理：取最后两个金额作为余额和前一个作为交易金额
        if len(amounts) >= 2:
            try:
                # 最后一个通常为余额
                balance = normalize_amount(amounts[-1].group(1))
                # 倒数第二个为交易金额
                txn_amount_raw = amounts[-2].group(1)
                txn_amount = normalize_amount(txn_amount_raw)

                # 判断是收入还是支出
                # 扫描件中可能包含 "收" "付" "借" "贷" 等标记
                before_last = after_date[:amounts[-2].start()]
                if '收' in before_last or '入' in before_last or '贷' in before_last:
                    credit_amount = txn_amount
                    direction = 'income'
                elif '付' in before_last or '出' in before_last or '借' in before_last:
                    debit_amount = txn_amount
                    direction = 'expense'
                else:
                    # 尝试从 OCR 文本中的标记判断
                    if credit_amount > 0:
                        direction = 'income'
                    else:
                        direction = 'expense'

                amount = debit_amount if debit_amount > 0 else credit_amount
            except (ValueError, IndexError):
                return None
        elif len(amounts) == 1:
            try:
                amount = normalize_amount(amounts[0].group(1))
                direction = 'income' if '收' in after_date or '入' in after_date else 'expense'
            except ValueError:
                return None
        else:
            return None

        if amount == 0:
            return None

        # 尝试提取对手户名（在金额之前的文字）
        # 这部分 OCR 可能很乱，保守处理
        counterparty = ''

        # 尝试找"对方户名"标记
        cp_match = re.search(r'对方[户名]?\s*([^\d\-]{2,20})', line)
        if cp_match:
            counterparty = cp_match.group(1).strip()

        return BankTransaction(
            file=file_path,
            bank_type=self.BANK_CODE,
            date=date,
            amount=amount,
            direction=direction,
            counterparty=counterparty,
            counterparty_account='',
            counterparty_bank='',
            summary='',
            serial_no='',
            balance=balance,
        )
