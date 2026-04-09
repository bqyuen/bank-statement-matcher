"""
金额标准化工具

处理各种格式的金额字符串转换为统一的 float 类型
"""
import re
from typing import Optional


class AmountNormalizer:
    """金额标准化工具类"""

    @staticmethod
    def normalize(amount_str: str) -> float:
        """
        将金额字符串标准化为 float

        支持格式：
        - 1,234.56
        - 1.234,56
        - 1 234.56
        - ¥1,234.56
        - $1,234.56
        - (1,234.56) 表示负数
        - -1,234.56

        Args:
            amount_str: 金额字符串

        Returns:
            标准化后的金额（float）
        """
        if not amount_str:
            return 0.0

        amount_str = str(amount_str).strip()

        # 判断是否为负数
        is_negative = False
        if amount_str.startswith('-') or amount_str.startswith('('):
            is_negative = True
            amount_str = amount_str.lstrip('-()')

        # 移除货币符号和空格
        amount_str = re.sub(r'[\¥\$\s]', '', amount_str)

        # 处理千分位逗号分隔的格式（美式：1,234.56）
        if ',' in amount_str and '.' in amount_str:
            # 判断哪种是千分位，哪种是小数点
            comma_pos = amount_str.rfind(',')
            dot_pos = amount_str.rfind('.')
            if comma_pos > dot_pos:
                # 1.234.567,89 格式（欧式）
                amount_str = amount_str.replace('.', '').replace(',', '.')
            else:
                # 1,234.567,89 格式（美式）
                amount_str = amount_str.replace(',', '')
        elif ',' in amount_str:
            # 只有逗号，可能是千分位分隔符
            # 尝试判断：123,456 通常是千分位，123,4 通常是小数点
            parts = amount_str.split(',')
            if len(parts[-1]) == 2 or len(parts[-1]) == 3:
                # 小数部分较短，可能是欧式小数
                if len(parts[-1]) == 2:
                    amount_str = amount_str.replace(',', '.')
                else:
                    amount_str = amount_str.replace(',', '')
            else:
                amount_str = amount_str.replace(',', '')

        try:
            result = float(amount_str)
            return -result if is_negative else result
        except ValueError:
            return 0.0

    @staticmethod
    def is_income(amount: float) -> bool:
        """
        判断金额是否为收入

        Args:
            amount: 金额（正数）

        Returns:
            True if income, False if expense
        """
        return amount >= 0

    @staticmethod
    def parse_direction(direction_str: str) -> str:
        """
        解析方向字符串

        Args:
            direction_str: 方向描述（如 '收入', '支出', '借', '贷' 等）

        Returns:
            'income' 或 'expense'
        """
        income_keywords = ['收入', '借', '来账', '+', '进账']
        expense_keywords = ['支出', '贷', '出账', '-', '扣款']

        direction_str = str(direction_str).strip().lower()

        for kw in income_keywords:
            if kw in direction_str:
                return 'income'

        for kw in expense_keywords:
            if kw in direction_str:
                return 'expense'

        # 默认根据正负号判断
        return 'income'


def normalize_amount(amount_str: str) -> float:
    """将金额字符串标准化为 float（兼容函数）"""
    return AmountNormalizer.normalize(amount_str)
