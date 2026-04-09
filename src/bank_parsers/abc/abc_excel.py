"""
农业银行交易明细（Excel/CSV）解析器

对账口径：
- 交易金额 > 0 -> income（收款）
- 交易金额 < 0 -> expense（付款）
"""

from __future__ import annotations

import csv
import os
from datetime import date, datetime
from typing import Dict, List, Optional, Sequence, Tuple

import openpyxl

from ..base import BaseBankParser, BankTransaction
from ..comm.amount import normalize_amount
from ..comm.date import normalize_date


def _norm_text(val: object) -> str:
    if val is None:
        return ""
    return str(val).replace("\n", " ").replace("\r", " ").strip()


def _norm_header(val: object) -> str:
    return _norm_text(val).replace(" ", "")


def _to_datetime(val: object) -> Optional[datetime]:
    if val is None:
        return None
    if isinstance(val, datetime):
        return val
    if isinstance(val, date):
        return datetime(val.year, val.month, val.day)
    return normalize_date(_norm_text(val))


class ABC_Excel_Parser(BaseBankParser):
    """农业银行流水解析器（Excel/CSV）"""

    BANK_CODE = "abc"
    _DATE_KEYS = ("交易日期", "记账日期", "入账日期", "发生日期", "日期")
    _AMOUNT_KEYS = ("交易金额", "交易发生额", "发生额", "金额")
    _SUMMARY_KEYS = ("摘要", "交易摘要", "附言", "用途", "交易用途", "摘要说明")
    _COUNTERPARTY_KEYS = ("对方户名", "对方名称", "对方账户名称", "收款人", "付款人", "交易对方")
    _COUNTERPARTY_ACCOUNT_KEYS = ("对方账号", "对方账户", "交易对方账号", "收款人账号", "付款人账号")
    _COUNTERPARTY_BANK_KEYS = ("对方开户行", "对方银行", "交易对方开户行")
    _SERIAL_KEYS = ("交易流水号", "流水号", "业务流水号", "交易序号", "凭证号")
    _BALANCE_KEYS = ("余额", "账户余额", "可用余额")

    def detect(self, file_path: str) -> bool:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in (".xlsx", ".xlsm", ".csv"):
            headers, context_text = self._peek_headers_and_context(file_path)
            has_amount = self._find_col(headers, self._AMOUNT_KEYS) is not None
            has_date = self._find_col(headers, self._DATE_KEYS) is not None
            has_abc_mark = any(k in context_text for k in ("农业银行", "中国农业银行", "agricultural bank", "abc"))
            has_trade_amount_col = any("交易金额" in h for h in headers)
            return bool(has_amount and has_date and (has_abc_mark or has_trade_amount_col))
        return False

    def parse(self, file_path: str) -> List[BankTransaction]:
        rows, headers = self._read_rows(file_path)
        col_map = self._build_col_map(headers)

        date_col = col_map.get("date")
        amount_col = col_map.get("amount")
        if date_col is None or amount_col is None:
            raise ValueError("农业银行流水文件缺少必要列：交易日期/交易金额")

        transactions: List[BankTransaction] = []
        for row in rows:
            if amount_col >= len(row):
                continue

            raw_amount = normalize_amount(_norm_text(row[amount_col]))
            if raw_amount == 0:
                continue

            direction = "income" if raw_amount > 0 else "expense"
            amount = abs(raw_amount)

            raw_date = row[date_col] if date_col < len(row) else None
            txn_date = _to_datetime(raw_date)
            if txn_date is None:
                continue

            txn = BankTransaction(
                file=file_path,
                bank_type=self.BANK_CODE,
                date=txn_date,
                amount=amount,
                direction=direction,
                counterparty=self._get_value(row, col_map.get("counterparty")),
                counterparty_account=self._get_value(row, col_map.get("counterparty_account")),
                counterparty_bank=self._get_value(row, col_map.get("counterparty_bank")),
                summary=self._get_value(row, col_map.get("summary")),
                serial_no=self._get_value(row, col_map.get("serial")),
                balance=normalize_amount(self._get_value(row, col_map.get("balance"))),
            )
            transactions.append(txn)

        if not transactions:
            raise ValueError("农业银行流水解析失败：未识别到有效交易记录")

        return sorted(transactions, key=lambda t: t.date)

    def _peek_headers_and_context(self, file_path: str) -> Tuple[List[str], str]:
        rows, headers = self._read_rows(file_path, max_rows=50)
        context_cells: List[str] = []
        for row in rows[:30]:
            for cell in row[:15]:
                txt = _norm_text(cell)
                if txt:
                    context_cells.append(txt.lower())
        return headers, " ".join(context_cells)

    def _read_rows(self, file_path: str, max_rows: Optional[int] = None) -> Tuple[List[List[object]], List[str]]:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".csv":
            return self._read_csv(file_path, max_rows=max_rows)
        return self._read_excel(file_path, max_rows=max_rows)

    def _read_excel(self, file_path: str, max_rows: Optional[int] = None) -> Tuple[List[List[object]], List[str]]:
        wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
        try:
            ws = wb.active
            raw_rows: List[List[object]] = []
            for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
                raw_rows.append(list(row))
                if max_rows and i >= max_rows:
                    break
        finally:
            wb.close()

        header_idx, headers = self._detect_header_row(raw_rows)
        data_rows = raw_rows[header_idx + 1:]
        return data_rows, headers

    def _read_csv(self, file_path: str, max_rows: Optional[int] = None) -> Tuple[List[List[object]], List[str]]:
        encodings = ("utf-8-sig", "gbk", "gb18030")
        rows: List[List[object]] = []
        last_err: Optional[Exception] = None
        for enc in encodings:
            try:
                with open(file_path, "r", encoding=enc, newline="") as f:
                    reader = csv.reader(f)
                    for i, row in enumerate(reader, start=1):
                        rows.append([c for c in row])
                        if max_rows and i >= max_rows:
                            break
                last_err = None
                break
            except Exception as exc:
                last_err = exc
                rows = []
        if last_err:
            raise ValueError(f"CSV 文件读取失败: {last_err}")

        header_idx, headers = self._detect_header_row(rows)
        data_rows = rows[header_idx + 1:]
        return data_rows, headers

    def _detect_header_row(self, rows: Sequence[Sequence[object]]) -> Tuple[int, List[str]]:
        for idx, row in enumerate(rows[:30]):
            headers = [_norm_header(c) for c in row]
            if self._find_col(headers, self._AMOUNT_KEYS) is not None and self._find_col(headers, self._DATE_KEYS) is not None:
                return idx, headers
        raise ValueError("未识别到农业银行流水表头（缺少交易日期/交易金额列）")

    def _build_col_map(self, headers: List[str]) -> Dict[str, Optional[int]]:
        return {
            "date": self._find_col(headers, self._DATE_KEYS),
            "amount": self._find_col(headers, self._AMOUNT_KEYS),
            "summary": self._find_col(headers, self._SUMMARY_KEYS),
            "counterparty": self._find_col(headers, self._COUNTERPARTY_KEYS),
            "counterparty_account": self._find_col(headers, self._COUNTERPARTY_ACCOUNT_KEYS),
            "counterparty_bank": self._find_col(headers, self._COUNTERPARTY_BANK_KEYS),
            "serial": self._find_col(headers, self._SERIAL_KEYS),
            "balance": self._find_col(headers, self._BALANCE_KEYS),
        }

    def _find_col(self, headers: Sequence[str], candidates: Sequence[str]) -> Optional[int]:
        for i, header in enumerate(headers):
            if not header:
                continue
            for key in candidates:
                if key in header:
                    return i
        return None

    def _get_value(self, row: Sequence[object], idx: Optional[int]) -> str:
        if idx is None or idx >= len(row):
            return ""
        return _norm_text(row[idx])
