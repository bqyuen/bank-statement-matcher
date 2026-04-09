"""
中交财务有限公司 PDF 对账单解析器

规则：
- 借方金额 > 0 -> expense（支付）
- 贷方金额 > 0 -> income（收入）
"""

from __future__ import annotations

import re
import tempfile
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import pymupdf

from ..base import BaseBankParser, BankTransaction
from ..comm.amount import normalize_amount
from ..comm.date import normalize_date


class CCFC_PDF_Parser(BaseBankParser):
    """中交财务公司 PDF 解析器（优先 OCR，兼容扫描件）"""

    BANK_CODE = "ccfc"
    _MONEY_RE = re.compile(r"^-?\d{1,3}(?:,\d{3})*\.\d{2}$")
    _MONEY_TOKEN_RE = re.compile(r"-?\d[\d,\s]{0,24}\.\s*\d{1,2}")
    _SERIAL_RE = re.compile(r"99\d{10,16}")

    def __init__(self) -> None:
        self._ocr_engine = None

    def detect(self, file_path: str) -> bool:
        if Path(file_path).suffix.lower() != ".pdf":
            return False

        # 文本型 PDF：检查关键字
        try:
            with pymupdf.open(file_path) as doc:
                for page in doc[:2]:
                    txt = page.get_text()
                    if "中交财务有限公司" in txt and "借方金额" in txt and "贷方金额" in txt:
                        return True
        except Exception:
            pass

        # 扫描件/图像型 PDF：检查页数（CCFC 对账单通常 >= 1 页）
        # 只要是 PDF 且有内容，就认为是可能的 CCFC 格式，交给 parse() 负责提取
        try:
            with pymupdf.open(file_path) as doc:
                if len(doc) >= 1:
                    return True
        except Exception:
            pass

        return False

    def parse(self, file_path: str) -> List[BankTransaction]:
        transactions: List[BankTransaction] = []
        seen = set()
        account_no = ""
        last_date = None

        with pymupdf.open(file_path) as doc:
            for page_idx in range(len(doc)):
                rows = self._ocr_page_lines(file_path, page_index=page_idx)
                if not rows:
                    continue

                header_info = self._find_header_info(rows)
                if not header_info:
                    continue
                header_cols, header_idx = header_info

                serial_indices = self._find_serial_row_indices(rows, header_idx + 1)

                # 兜底：处理首个流水号之前、无流水号但有金额的记录（常见于首页首笔）
                first_serial_idx = serial_indices[0] if serial_indices else len(rows)
                for ridx in range(header_idx + 1, first_serial_idx):
                    row_text = " ".join(tok for _, tok in rows[ridx])
                    if self._is_skip_row(row_text):
                        continue
                    date_val = self._extract_date(row_text)
                    if date_val is None:
                        continue
                    parsed = self._extract_amounts_from_block([rows[ridx]], header_cols)
                    if parsed is None:
                        continue
                    debit, credit, balance = parsed
                    if debit > 0 and credit == 0:
                        direction = "expense"
                        amount = debit
                    elif credit > 0 and debit == 0:
                        direction = "income"
                        amount = credit
                    else:
                        continue

                    summary_parts = [row_text]
                    if ridx > header_idx + 1:
                        summary_parts.insert(0, " ".join(tok for _, tok in rows[ridx - 1]))
                    if ridx + 1 < first_serial_idx:
                        summary_parts.append(" ".join(tok for _, tok in rows[ridx + 1]))
                    summary = self._extract_summary(" ".join(summary_parts))
                    counterparty = self._extract_counterparty(rows[ridx], header_cols)

                    if not self._is_valid_transaction(
                        row_text=row_text,
                        summary=summary,
                        counterparty=counterparty,
                        serial_no="",
                        debit=debit,
                        credit=credit,
                    ):
                        continue

                    dedup_key = (
                        date_val.strftime("%Y-%m-%d"),
                        f"{amount:.2f}",
                        direction,
                        "",
                        counterparty,
                        summary,
                        f"{balance:.2f}",
                    )
                    if dedup_key in seen:
                        continue
                    seen.add(dedup_key)
                    transactions.append(
                        BankTransaction(
                            file=file_path,
                            bank_type=self.BANK_CODE,
                            date=date_val,
                            amount=amount,
                            direction=direction,
                            counterparty=counterparty,
                            counterparty_account=account_no,
                            counterparty_bank="中交财务有限公司",
                            summary=summary,
                            serial_no="",
                            balance=balance,
                        )
                    )
                    last_date = date_val

                blocks = self._build_transaction_blocks(rows, header_idx, serial_indices)
                for block_rows in blocks:
                    block_text = " ".join(" ".join(tok for _, tok in row) for row in block_rows)

                    acc = self._extract_account_no(block_text)
                    if acc:
                        account_no = acc

                    serial_no = self._extract_serial_no(block_text)
                    if not serial_no:
                        continue

                    date_val = self._extract_date_from_block(block_rows, last_date)
                    if date_val is None:
                        continue
                    last_date = date_val

                    parsed = self._extract_amounts_from_block(block_rows, header_cols)
                    if parsed is None:
                        continue
                    debit, credit, balance = parsed

                    if debit > 0 and credit == 0:
                        direction = "expense"
                        amount = debit
                    elif credit > 0 and debit == 0:
                        direction = "income"
                        amount = credit
                    elif debit > 0 and credit > 0:
                        # 异常双边同时有值，优先按更接近业务语义的关键词判断
                        if self._looks_like_income(block_text):
                            direction = "income"
                            amount = credit
                            debit = 0.0
                        else:
                            direction = "expense"
                            amount = debit
                            credit = 0.0
                    else:
                        continue

                    summary = self._extract_summary(block_text)
                    counterparty = self._extract_counterparty_from_block(block_rows, header_cols)

                    if not self._is_valid_transaction(
                        row_text=block_text,
                        summary=summary,
                        counterparty=counterparty,
                        serial_no=serial_no,
                        debit=debit,
                        credit=credit,
                    ):
                        continue

                    dedup_key = (
                        date_val.strftime("%Y-%m-%d"),
                        f"{amount:.2f}",
                        direction,
                        serial_no,
                        counterparty,
                        summary,
                        f"{balance:.2f}",
                    )
                    if dedup_key in seen:
                        continue
                    seen.add(dedup_key)

                    transactions.append(
                        BankTransaction(
                            file=file_path,
                            bank_type=self.BANK_CODE,
                            date=date_val,
                            amount=amount,
                            direction=direction,
                            counterparty=counterparty,
                            counterparty_account=account_no,
                            counterparty_bank="中交财务有限公司",
                            summary=summary,
                            serial_no=serial_no,
                            balance=balance,
                        )
                    )

        return transactions

    def _get_ocr_engine(self):
        if self._ocr_engine is not None:
            return self._ocr_engine
        try:
            from rapidocr_onnxruntime import RapidOCR  # type: ignore

            self._ocr_engine = RapidOCR()
            return self._ocr_engine
        except Exception:
            self._ocr_engine = False
            return False

    def _ocr_page_lines(self, file_path: str, page_index: int) -> List[List[Tuple[float, str]]]:
        engine = self._get_ocr_engine()
        if not engine:
            return []

        with pymupdf.open(file_path) as doc:
            page = doc[page_index]
            pix = page.get_pixmap(matrix=pymupdf.Matrix(2, 2))
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                pix.save(tmp.name)
                img_path = tmp.name

        try:
            result, _ = engine(img_path)
        except Exception:
            result = None
        finally:
            Path(img_path).unlink(missing_ok=True)

        if not result:
            return []

        rows: Dict[int, List[Tuple[float, str]]] = {}
        for box, text, _score in result:
            token = str(text).strip()
            if not token:
                continue
            ys = [p[1] for p in box]
            xs = [p[0] for p in box]
            y_key = int(round((sum(ys) / 4) / 6) * 6)
            rows.setdefault(y_key, []).append((float(min(xs)), token))

        merged_rows: List[List[Tuple[float, str]]] = []
        for y in sorted(rows.keys()):
            merged_rows.append(sorted(rows[y], key=lambda t: t[0]))
        return merged_rows

    def _find_header_info(
        self, rows: Sequence[Sequence[Tuple[float, str]]]
    ) -> Optional[Tuple[Dict[str, float], int]]:
        for idx, row in enumerate(rows):
            cols = {txt: x for x, txt in row}
            joined = " ".join(txt for _, txt in row)
            if "借方金额" in joined and "贷方金额" in joined:
                return (
                    {
                        "debit": self._pick_x(cols, "借方金额"),
                        "credit": self._pick_x(cols, "贷方金额"),
                        "balance": self._pick_x(cols, "余额"),
                    },
                    idx,
                )
        return None

    def _pick_x(self, cols: Dict[str, float], key: str) -> float:
        for txt, x in cols.items():
            if key in txt:
                return x
        if key == "余额" and cols:
            return max(cols.values())
        return 0.0

    def _find_serial_row_indices(
        self, rows: Sequence[Sequence[Tuple[float, str]]], start_idx: int
    ) -> List[int]:
        serial_rows = []
        for idx in range(start_idx, len(rows)):
            row_text = " ".join(tok for _, tok in rows[idx])
            if self._SERIAL_RE.search(row_text):
                serial_rows.append(idx)
        return serial_rows

    def _build_transaction_blocks(
        self,
        rows: Sequence[Sequence[Tuple[float, str]]],
        header_idx: int,
        serial_rows: Optional[Sequence[int]] = None,
    ) -> List[List[Sequence[Tuple[float, str]]]]:
        if serial_rows is None:
            serial_rows = self._find_serial_row_indices(rows, header_idx + 1)

        if not serial_rows:
            return []

        blocks: List[List[Sequence[Tuple[float, str]]]] = []
        for i, current_idx in enumerate(serial_rows):
            start = serial_rows[i - 1] + 1 if i > 0 else header_idx + 1
            end = serial_rows[i + 1] if i + 1 < len(serial_rows) else len(rows)
            if current_idx < start:
                start = current_idx
            block_rows: List[Sequence[Tuple[float, str]]] = []
            for row_idx in range(start, end):
                block_rows.append(rows[row_idx])
            if block_rows:
                blocks.append(block_rows)
        return blocks

    def _is_skip_row(self, row_text: str) -> bool:
        norm = self._normalize_text(row_text)
        return any(
            k in norm
            for k in (
                "页共",
                "对账单",
                "开始日期",
                "结束日期",
                "业务日期",
                "借方金额",
                "贷方金额",
                "日总计",
                "月总计",
                "本页合计",
                "发生额合计",
            )
        )

    def _extract_money_values(self, token: str) -> List[float]:
        raw = token.replace("，", ",").replace("。", ".")
        vals: List[float] = []
        for m in self._MONEY_TOKEN_RE.findall(raw):
            t = m.replace(" ", "")
            t = re.sub(r"(?<=\d)\.(?=\d{3},)", "", t)
            if re.fullmatch(r"-?\d{1,3}(?:,\d{3})*\.\d", t):
                t = f"{t}0"
            if not self._MONEY_RE.fullmatch(t):
                continue
            vals.append(normalize_amount(t))
        return vals

    def _assign_amounts_by_column(
        self, money_items: Sequence[Tuple[float, float]], header_cols: Dict[str, float]
    ) -> Tuple[float, float, float]:
        debit = 0.0
        credit = 0.0
        balance = 0.0
        best_dist = {"debit": 1e9, "credit": 1e9, "balance": 1e9}
        targets = {
            "debit": header_cols.get("debit", 0.0),
            "credit": header_cols.get("credit", 0.0),
            "balance": header_cols.get("balance", 0.0),
        }

        for x, val in money_items:
            nearest = min(targets.keys(), key=lambda k: abs(x - targets[k]))
            dist = abs(x - targets[nearest])
            if dist < best_dist[nearest]:
                best_dist[nearest] = dist
                if nearest == "debit":
                    debit = val
                elif nearest == "credit":
                    credit = val
                else:
                    balance = val

        if balance == 0.0 and money_items:
            # OCR 丢列时兜底：最右侧金额当余额
            balance = max(money_items, key=lambda t: t[0])[1]
        return debit, credit, balance

    def _extract_amounts_from_block(
        self, block_rows: Sequence[Sequence[Tuple[float, str]]], header_cols: Dict[str, float]
    ) -> Optional[Tuple[float, float, float]]:
        money_items: List[Tuple[float, float]] = []
        signed_vals: List[float] = []
        for row in block_rows:
            for x, txt in row:
                vals = self._extract_money_values(txt)
                for v in vals:
                    signed_vals.append(v)
                    money_items.append((x, abs(v)))

        if not money_items:
            return None

        debit, credit, balance = self._assign_amounts_by_column(money_items, header_cols)
        if debit > 0 or credit > 0:
            return debit, credit, balance

        # 兜底：列定位失败时，取块内第一个非零金额作为业务金额
        non_zero = [abs(v) for v in signed_vals if abs(v) > 0.0001]
        if not non_zero:
            return None
        amount = non_zero[0]
        block_text = " ".join(" ".join(txt for _, txt in row) for row in block_rows)
        if self._looks_like_income(block_text):
            return 0.0, amount, balance
        return amount, 0.0, balance

    def _extract_date_from_block(self, block_rows, last_date):
        ordered_rows = sorted(
            block_rows,
            key=lambda r: 0 if self._SERIAL_RE.search(" ".join(tok for _, tok in r)) else 1,
        )
        for row in ordered_rows:
            row_text = " ".join(tok for _, tok in row)
            if self._is_skip_row(row_text):
                continue
            dt = self._extract_date(row_text)
            if dt:
                return dt
        return last_date

    def _extract_date(self, row_text: str) -> Optional[object]:
        s = row_text or ""
        text = s.strip()[:40]
        m = re.search(r"^\D{0,2}(20\d{2})\D{0,3}([01]?\d)\D{0,3}([0-3]?\d)", text)
        if m:
            try:
                year = int(m.group(1))
                month = int(m.group(2))
                day = int(m.group(3))
            except Exception:
                year = month = day = 0
            if 1 <= month <= 12 and 1 <= day <= 31:
                dt = normalize_date(f"{year:04d}-{month:02d}-{day:02d}", "%Y-%m-%d")
                if dt:
                    return dt

        m2 = re.search(r"^\D{0,2}(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\b", text)
        if m2:
            dt = normalize_date(f"{m2.group(1)}-{m2.group(2)}-{m2.group(3)}", "%Y-%m-%d")
            if dt:
                return dt
        return None

    def _extract_account_no(self, row_text: str) -> str:
        m = re.search(r"01-\d{2}-\d{6}-\d{2}", row_text)
        return m.group(0) if m else ""

    def _extract_serial_no(self, row_text: str) -> str:
        m = self._SERIAL_RE.search(row_text)
        if m:
            s = m.group(0)
            return s[:14] if len(s) > 14 else s
        m = re.search(r"\b\d{8,16}\b", row_text)
        return m.group(0) if m else ""

    def _extract_counterparty(self, row: Sequence[Tuple[float, str]], header_cols: Dict[str, float]) -> str:
        candidates = []
        debit_x = header_cols.get("debit", 0.0)
        for x, txt in row:
            if x >= debit_x:
                continue
            if re.search(r"[\u4e00-\u9fff]{2,}", txt):
                if any(k in txt for k in ("对账单", "日总计", "月总计", "业务日期", "借方金额", "贷方金额")):
                    continue
                candidates.append(txt)
        if not candidates:
            return ""
        return max(candidates, key=len)

    def _extract_counterparty_from_block(
        self, block_rows: Sequence[Sequence[Tuple[float, str]]], header_cols: Dict[str, float]
    ) -> str:
        candidates = []
        debit_x = header_cols.get("debit", 0.0)
        for row in block_rows:
            for x, txt in row:
                if x >= debit_x:
                    continue
                if not re.search(r"[\u4e00-\u9fff]{2,}", txt):
                    continue
                if any(k in txt for k in ("对账单", "日总计", "月总计", "业务日期", "借方金额", "贷方金额")):
                    continue
                if "01-10-" in txt:
                    continue
                candidates.append(txt)
        if not candidates:
            return ""
        return max(candidates, key=len)

    def _extract_summary(self, row_text: str) -> str:
        s = row_text
        s = re.sub(r"20\d{2}\D{0,3}[01]?\d\D{0,3}[0-3]?\d", " ", s)
        s = re.sub(r"99\d{10,16}", " ", s)
        s = re.sub(r"01-\d{2}-\d{6}-\d{2}", " ", s)
        s = re.sub(r"-?\d[\d,\.\s]{2,}", " ", s)
        for k in (
            "借方金额",
            "贷方金额",
            "余额",
            "业务日期",
            "操作日期",
            "凭证号",
            "对方账号",
            "对方户名",
            "起息日期",
            "发生额合计",
        ):
            s = s.replace(k, " ")
        s = re.sub(r"\s+", " ", s).strip()
        return s

    def _looks_like_income(self, text: str) -> bool:
        norm = self._normalize_text(text)
        return any(
            k in norm
            for k in (
                "自动资金上划",
                "日终调平生成下拨单",
                "收款",
                "贷方",
                "收入",
            )
        )

    def _normalize_text(self, text: str) -> str:
        return re.sub(r"\s+", "", text or "")

    def _is_valid_transaction(
        self,
        row_text: str,
        summary: str,
        counterparty: str,
        serial_no: str,
        debit: float,
        credit: float,
    ) -> bool:
        row_norm = self._normalize_text(row_text)
        summary_norm = self._normalize_text(summary)
        counterparty_norm = self._normalize_text(counterparty)

        # 有流水号+金额即保留，避免 OCR 夹杂“日总计”等噪声词被误杀
        if serial_no and (debit > 0 or credit > 0):
            return True

        if any(k in row_norm for k in ("日总计", "月总计", "本页合计", "总计", "发生额合计")):
            return False
        if summary_norm in ("日总计", "月总计", "总计", "发生额合计"):
            return False

        if debit > 0 and credit > 0:
            return False

        if not summary_norm and not counterparty_norm:
            return False

        return True
