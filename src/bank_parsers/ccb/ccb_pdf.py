"""
建设银行 PDF 对账单解析器

解析 CCB PDF 格式的银行对账单（标准借贷分离式）
识别特征：含"中国建设银行账户明细信息"
"""
import re
from typing import List, Optional
import pymupdf

from ..base import BaseBankParser, BankTransaction
from ..comm.amount import normalize_amount
from ..comm.date import normalize_date


def _clean_text(text: str) -> str:
    """清洗文本：去除换行符和多余空格"""
    if not text:
        return ''
    # 替换换行符为空格，去除多余空格
    text = re.sub(r'[\n\r]+', '', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


class CCB_PDF_Parser(BaseBankParser):
    """CCB PDF 对账单解析器（标准借贷分离式）"""

    BANK_CODE = 'ccb'

    def detect(self, file_path: str) -> bool:
        """
        检测是否为 CCB PDF 对账单格式

        通过内容特征判断：含"中国建设银行账户明细信息"
        """
        import os
        ext = os.path.splitext(file_path)[1].lower()
        if ext != '.pdf':
            return False

        try:
            with pymupdf.open(file_path) as doc:
                for page in doc:
                    text = page.get_text()
                    if '中国建设银行账户明细信息' in text:
                        return True
        except Exception:
            pass
        return False

    def parse(self, file_path: str) -> List[BankTransaction]:
        """
        解析 CCB PDF 对账单

        Args:
            file_path: PDF 文件路径

        Returns:
            BankTransaction 列表
        """
        results: List[BankTransaction] = []

        with pymupdf.open(file_path) as doc:
            for page in doc:
                page_txns = self._parse_page(page, file_path)
                results.extend(page_txns)

        return results

    def _parse_page(self, page, file_path: str) -> List[BankTransaction]:
        """解析单页内容"""
        transactions: List[BankTransaction] = []

        tabs = page.find_tables()
        if not tabs or not tabs.tables:
            return transactions

        table = tabs.tables[0]
        rows = table.extract()
        if not rows or len(rows) < 2:
            return transactions

        # 跳过表头行
        for row in rows[1:]:
            txn = self._parse_row(row, file_path)
            if txn:
                transactions.append(txn)

        return transactions

    def _parse_row(self, row: List[str], file_path: str) -> Optional[BankTransaction]:
        """解析单行数据"""
        if len(row) < 11:
            return None

        # 列定义：
        # 0:账号, 1:交易时间, 2:借方发生额, 3:贷方发生额,
        # 4:余额, 5:币种, 6:对方户名, 7:对方账号,
        # 8:对方开户机构, 9:记账日期, 10:摘要, 11:备注,
        # 12:账户明细编号-交易流水号, 13:企业流水号,
        # 14:凭证种类, 15:凭证号, 16:交易介质编号

        # 解析金额和方向
        debit_str = _clean_text(row[2]) if len(row) > 2 else ''
        credit_str = _clean_text(row[3]) if len(row) > 3 else ''

        debit_amount = normalize_amount(debit_str)
        credit_amount = normalize_amount(credit_str)

        # 收支方向：借方发生额>0→expense，贷方发生额>0→income
        if debit_amount > 0:
            amount = debit_amount
            direction = 'expense'
        elif credit_amount > 0:
            amount = credit_amount
            direction = 'income'
        else:
            # 没有金额，跳过
            return None

        # 解析记账日期
        date_str = _clean_text(row[9]) if len(row) > 9 else ''
        # 日期格式：YYYYMMDD
        date = normalize_date(date_str, '%Y%m%d')
        if date is None:
            date_str_raw = re.search(r'\d{8}', date_str)
            if date_str_raw:
                date = normalize_date(date_str_raw.group(), '%Y%m%d')
        if date is None:
            return None

        # 解析余额
        balance_str = _clean_text(row[4]) if len(row) > 4 else ''
        balance = normalize_amount(balance_str)

        # 解析交易时间（用于流水号）
        txn_time_str = _clean_text(row[1]) if len(row) > 1 else ''
        txn_time_match = re.search(r'\d{8}[\s\d:]*', txn_time_str)

        # 解析对方户名
        counterparty = _clean_text(row[6]) if len(row) > 6 else ''

        # 解析对方账号
        counterparty_account = _clean_text(row[7]) if len(row) > 7 else ''

        # 解析对方开户机构
        counterparty_bank = _clean_text(row[8]) if len(row) > 8 else ''

        # 解析摘要（合并摘要和备注）
        summary = _clean_text(row[10]) if len(row) > 10 else ''
        remark = _clean_text(row[11]) if len(row) > 11 else ''
        if remark and remark not in summary:
            full_summary = f"{summary} {remark}".strip()
        else:
            full_summary = summary

        # 解析流水号（交易流水号）
        serial_no = ''
        if len(row) > 12:
            serial_no_raw = _clean_text(row[12])
            # 格式如 "81-3519818014YGCW3WL7L"，取后半部分
            if '-' in serial_no_raw:
                serial_no = serial_no_raw.split('-', 1)[-1]
            else:
                serial_no = serial_no_raw

        return BankTransaction(
            file=file_path,
            bank_type=self.BANK_CODE,
            date=date,
            amount=amount,
            direction=direction,
            counterparty=counterparty,
            counterparty_account=counterparty_account,
            counterparty_bank=counterparty_bank,
            summary=full_summary,
            serial_no=serial_no,
            balance=balance,
        )
