"""
银行流水与三栏账匹配引擎

当前匹配规则：
- 金额精确一致
- 收支方向一致（income/expense）
- 不使用日期容差
"""

from typing import List, Tuple, Any

from bank_parsers.base import BankTransaction
from ledger.ledger_parser import LedgerEntry


class BankLedgerMatcher:
    """银行流水与三栏账匹配器"""

    def __init__(self, tolerance_near: int = 3, tolerance_far: int = 7):
        """
        Args:
            tolerance_near: 近端日期容差（天），用于L1级精确匹配
            tolerance_far:  远端日期容差（天），用于L2~L9级模糊匹配
        """
        self.tolerance_near = tolerance_near
        self.tolerance_far = tolerance_far

    def match(
        self,
        bank_txns: List[BankTransaction],
        ledger_entries: List[LedgerEntry]
    ) -> Tuple[
        List[Tuple[BankTransaction, LedgerEntry, str]],  # matched
        List[BankTransaction],                             # bank_only
        List[LedgerEntry]                                 # ledger_only
    ]:
        """
        执行匹配

        Returns:
            (matched, bank_only, ledger_only)
        """
        matched: List[Tuple[BankTransaction, LedgerEntry, str]] = []
        bank_used = set()
        ledger_used = set()

        # 按金额分组以加速查找
        amount_groups: dict = {}
        for i, entry in enumerate(ledger_entries):
            amount_groups.setdefault(entry.amount, []).append(i)

        # L1 ~ L9 逐级匹配
        for level in ["L1", "L2", "L3", "L4", "L5", "L6", "L7", "L8", "L9"]:
            if bank_used == set(range(len(bank_txns))) and ledger_used == set(range(len(ledger_entries))):
                break

            for bi, bank_txn in enumerate(bank_txns):
                if bi in bank_used:
                    continue

                candidates = amount_groups.get(bank_txn.amount, [])
                for li in candidates:
                    if li in ledger_used:
                        continue

                    ledger_entry = ledger_entries[li]
                    if self._match_level(bank_txn, ledger_entry, level):
                        matched.append((bank_txn, ledger_entry, level))
                        bank_used.add(bi)
                        ledger_used.add(li)
                        break

        # 未匹配记录
        bank_only = [t for i, t in enumerate(bank_txns) if i not in bank_used]
        ledger_only = [e for i, e in enumerate(ledger_entries) if i not in ledger_used]

        return matched, bank_only, ledger_only

    def _match_level(
        self,
        bank_txn: BankTransaction,
        ledger_entry: LedgerEntry,
        level: str
    ) -> bool:
        """判断是否满足指定匹配级别"""
        # 收支方向必须一致：银行income<->三栏账借方，银行expense<->三栏账贷方
        if bank_txn.direction != ledger_entry.direction:
            return False

        # 不进行日期容差判断：金额+方向一致即匹配
        return True
