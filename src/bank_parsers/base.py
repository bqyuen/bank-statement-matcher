"""
银行流水解析器基类
"""
from abc import ABC, abstractmethod
from typing import List
from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class BankTransaction:
    """银行交易记录"""
    file: str                    # 来源文件
    bank_type: str               # 银行类型: 'ccb' / 'cmb'
    date: datetime               # 交易日期
    amount: float                # 交易金额
    direction: str               # 方向: 'income' (收入) / 'expense' (支出)
    counterparty: str            # 交易对手名称
    counterparty_account: str     # 交易对手账号
    counterparty_bank: str       # 交易对手开户行
    summary: str                  # 摘要
    serial_no: str                # 流水号/交易序号
    balance: float                # 余额
    mate: bool = False           # 是否已匹配
    matched_voucher: str = None  # 匹配的凭证号
    match_level: str = None       # 匹配级别: 'L1'~'L9'


class BaseBankParser(ABC):
    """银行解析器抽象基类"""

    @abstractmethod
    def detect(self, file_path: str) -> bool:
        """
        检测文件是否属于当前解析器支持的格式

        Args:
            file_path: 文件路径

        Returns:
            True if this parser can handle the file, False otherwise
        """
        pass

    @abstractmethod
    def parse(self, file_path: str) -> List[BankTransaction]:
        """
        解析文件并返回交易记录列表

        Args:
            file_path: 文件路径

        Returns:
            BankTransaction 列表
        """
        pass
