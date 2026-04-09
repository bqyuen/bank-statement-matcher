"""
招商银行账务明细清单 PDF 解析器
"""
import re
from typing import List, Optional, Tuple
import pymupdf


def _find_amounts(text: str) -> List[str]:
    """
    在文本中查找所有金额，支持处理账户号与金额紧邻的边缘情况。

    格式：-?\d{1,3}(?:,\d{3})*\.\d{2}
    """
    pattern = r'-?[\d,]+\.\d{2}'
    raw = re.findall(pattern, text)

    result: List[str] = []
    for amt_str in raw:
        sign = '-' if amt_str.startswith('-') else ''
        before = amt_str.lstrip('-')
        parts = before.split('.')
        int_part = parts[0]
        dec_part = parts[1]
        digits_only = int_part.replace(',', '')

        if len(digits_only) > 10:
            # 可能为账户号与金额紧邻的误匹配，进行修复
            comma_groups = int_part.split(',')
            if len(comma_groups) >= 2 and len(comma_groups[-1]) == 3:
                second_last = comma_groups[-2].replace(',', '')
                for n in range(1, 5):
                    candidate = second_last[-n:] + ',' + comma_groups[-1]
                    if len(candidate.replace(',', '')) <= 7:
                        result.append(sign + candidate + '.' + dec_part)
                        break
                else:
                    result.append(amt_str)
            else:
                result.append(amt_str)
        else:
            result.append(amt_str)

    return result

from ..base import BaseBankParser, BankTransaction
from ..comm.amount import normalize_amount
from ..comm.date import normalize_date


class CMB_PDF_Parser(BaseBankParser):
    """招商银行账务明细清单解析器"""

    BANK_CODE = 'cmb'

    def detect(self, file_path: str) -> bool:
        """检测文件是否为招商银行账务明细清单"""
        with pymupdf.open(file_path) as doc:
            for page in doc:
                text = page.get_text()
                if '账务明细清单' in text and '招商银行' in text:
                    return True
        return False

    def parse(self, file_path: str) -> List[BankTransaction]:
        """解析招商银行账务明细清单PDF"""
        transactions: List[BankTransaction] = []

        with pymupdf.open(file_path) as doc:
            for page_no, page in enumerate(doc, start=1):
                page_txns = self._parse_page(page, file_path, page_no)
                transactions.extend(page_txns)

        return transactions

    def _parse_page(self, page, file_path: str, page_no: int) -> List[BankTransaction]:
        """
        解析单页内容

        使用 'dict' 模式获取带语义类型的文本块，
        按 y 坐标分行后逐块处理。

        CMB 格式有两种行：
        1. 完整行：日期+类型+票据号+摘要+金额+余额+对手（所有字段在一块）
        2. 拆分行：左块=日期+类型+票据号，右块=摘要续+金额+余额+对手
           在 y 排序后，右块先出现，左块后出现。
        """
        page_dict = page.get_text('dict')
        blocks = page_dict.get('blocks', [])
        text_blocks = [b for b in blocks if b.get('type') == 0]
        text_blocks.sort(key=lambda b: b.get('bbox', [0, 0])[1])

        transactions: List[BankTransaction] = []
        i = 0
        n = len(text_blocks)

        # 跳过表头（y < 140，表头区域约到 y=136）
        while i < n and text_blocks[i].get('bbox', [0, 0])[1] < 140:
            i += 1

        # 待合并的右块队列（同 row_key 的右块）
        pending_right: List[Tuple[str, float]] = []  # [(text, y), ...]
        pending_row_key: List[int] = []  # 对应的 row_key

        while i < n:
            block = text_blocks[i]
            block_text = self._block_text(block).strip()
            bbox = block.get('bbox', [0, 0, 0, 0])
            x0, y0 = bbox[0], bbox[1]

            if not block_text:
                i += 1
                continue

            has_amount = self._contains_amount(block_text)
            starts_with_date = self._starts_with_date(block_text)
            row_key = round(y0 / 10) * 10  # 按 10px 分组行

            if not starts_with_date and has_amount and x0 > 150:
                # 右列块（无日期，有金额）：可能是与下一左块配对的右块
                # 记录为 pending
                pending_right.append((block_text, y0))
                pending_row_key.append(row_key)
                i += 1
                continue

            elif starts_with_date and not has_amount:
                # 左列块（有日期，无金额）：与 pending 右块配对
                merged_text = block_text
                merged_y = y0
                # 合并所有同 row_key 的 pending 右块
                j = 0
                while j < len(pending_row_key):
                    if pending_row_key[j] == row_key:
                        merged_text += pending_right[j][0]
                        merged_y = pending_right[j][1]
                        pending_right.pop(j)
                        pending_row_key.pop(j)
                    else:
                        j += 1

                txn = self._parse_transaction_line(merged_text, file_path, page_no)
                if txn:
                    transactions.append(txn)
                i += 1
                continue

            elif starts_with_date and has_amount:
                # 完整行（含日期和金额）
                txn = self._parse_transaction_line(block_text, file_path, page_no)
                if txn:
                    transactions.append(txn)
                i += 1
                continue

            else:
                # 非日期行且无金额（分隔符等），跳过
                i += 1
                continue

        return transactions

    def _block_text(self, block: dict) -> str:
        """从 dict 格式的文本块中提取所有文字"""
        parts = []
        for line in block.get('lines', []):
            for span in line.get('spans', []):
                parts.append(span.get('text', ''))
        return ''.join(parts)

    def _starts_with_date(self, text: str) -> bool:
        """判断文本是否以日期开头（YYYYMMDD）"""
        return bool(re.match(r'^\d{8}', text.strip()))

    def _contains_amount(self, text: str) -> bool:
        """判断文本是否包含有效金额"""
        return bool(re.search(r'-?[\d,]+\.\d{2}', text))

    def _parse_transaction_line(self, text: str, file_path: str, page_no: int) -> Optional[BankTransaction]:
        """
        解析单行交易数据

        格式：日期 业务类型 票据号 摘要 借方/贷方金额 余额 对手户名
        """
        text = text.strip()
        if not text:
            return None

        # 1. 日期（前8位，YYYYMMDD）
        date_str = text[:8]
        try:
            date = normalize_date(date_str)
        except Exception:
            return None

        remainder = text[8:].strip()

        # 2. 提取所有金额（使用修复版查找函数）
        all_amounts = _find_amounts(remainder)
        if not all_amounts:
            return None

        # 倒数第二个=交易金额，最后一个=余额
        amount_str = all_amounts[-2] if len(all_amounts) >= 2 else all_amounts[-1]
        balance_str = all_amounts[-1]

        amount = normalize_amount(amount_str)
        balance = normalize_amount(balance_str)

        # 方向：负数→expense（取绝对值），正数→income
        if amount < 0:
            amount = abs(amount)
            direction = 'expense'
        else:
            direction = 'income'

        # 3. 对手户名 = 余额之后所有文字
        balance_end = remainder.rfind(balance_str) + len(balance_str)
        counterparty = remainder[balance_end:].strip()

        # 4. 票据号（10位以上数字）
        bill_no = ''
        bill_match = re.search(r'\d{10,}', remainder)
        if bill_match:
            bill_no = bill_match.group(0)

        # 5. 摘要 = 票据号结束后、第一个金额之前
        summary = ''
        if bill_match:
            summary_start = bill_match.end()
            first_amt_pos = remainder.find(amount_str)
            if first_amt_pos > summary_start:
                summary = remainder[summary_start:first_amt_pos].strip()

        return BankTransaction(
            file=file_path,
            bank_type=self.BANK_CODE,
            date=date,
            amount=amount,
            direction=direction,
            counterparty=counterparty,
            counterparty_account='',
            counterparty_bank='',
            summary=summary,
            serial_no=bill_no,
            balance=balance,
        )
