"""
日期标准化工具

处理各种格式的日期字符串转换为 datetime 对象
"""
import re
from datetime import datetime
from typing import Optional, Tuple


class DateNormalizer:
    """日期标准化工具类"""

    # 支持的日期格式列表
    DATE_FORMATS = [
        '%Y-%m-%d',
        '%Y/%m/%d',
        '%Y.%m.%d',
        '%Y年%m月%d日',
        '%d-%m-%Y',
        '%d/%m/%Y',
        '%d.%m.%Y',
        '%m-%d-%Y',
        '%m/%d/%Y',
        '%Y%m%d',
        '%Y%m%d%H%M%S',
        '%Y-%m-%d %H:%M:%S',
        '%Y/%m/%d %H:%M:%S',
    ]

    @classmethod
    def normalize(cls, date_str: str, fmt_hint: Optional[str] = None) -> Optional[datetime]:
        """
        将日期字符串标准化为 datetime 对象

        Args:
            date_str: 日期字符串
            fmt_hint: 可选的格式提示

        Returns:
            datetime 对象，解析失败返回 None
        """
        if not date_str:
            return None

        date_str = str(date_str).strip()

        # 尝试使用提示的格式
        if fmt_hint:
            try:
                return datetime.strptime(date_str, fmt_hint)
            except ValueError:
                pass

        # 尝试所有支持的格式
        for fmt in cls.DATE_FORMATS:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue

        # 尝试从字符串中提取日期
        extracted = cls.extract_date(date_str)
        if extracted:
            return extracted

        return None

    @classmethod
    def extract_date(cls, text: str) -> Optional[datetime]:
        """
        从文本中提取日期

        支持格式：
        - 2024年01月15日
        - 2024-01-15
        - 2024/01/15
        - 20240115
        - 01/15/2024
        - 15/01/2024
        """
        patterns = [
            # 年月日
            (r'(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})日?', '%Y', '%m', '%d'),
            # 年月
            (r'(\d{4})[年\-/](\d{1,2})月?', '%Y', '%m', None),
        ]

        for pattern, y_fmt, m_fmt, d_fmt in patterns:
            match = re.search(pattern, text)
            if match:
                year = int(match.group(1))
                month = int(match.group(2))
                day = int(match.group(3)) if d_fmt else 1

                # 简单校验
                if 1900 <= year <= 2100 and 1 <= month <= 12 and 1 <= day <= 31:
                    try:
                        return datetime(year, month, day)
                    except ValueError:
                        continue

        return None

    @classmethod
    def parse_year_month(cls, year_month_str: str) -> Tuple[int, int]:
        """
        解析年月字符串

        Args:
            year_month_str: 年月字符串（如 '2024-03', '2024年3月'）

        Returns:
            (year, month) 元组
        """
        if not year_month_str:
            return (datetime.now().year, datetime.now().month)

        year_month_str = str(year_month_str).strip()

        # 尝试各种格式
        patterns = [
            r'(\d{4})[年\-/](\d{1,2})',
            r'(\d{4})(\d{2})',
        ]

        for pattern in patterns:
            match = re.search(pattern, year_month_str)
            if match:
                year = int(match.group(1))
                month = int(match.group(2))
                if 1900 <= year <= 2100 and 1 <= month <= 12:
                    return (year, month)

        # 返回当前年月
        now = datetime.now()
        return (now.year, now.month)


def normalize_date(date_str: str, fmt_hint: Optional[str] = None) -> Optional[datetime]:
    """将日期字符串标准化为 datetime 对象（兼容函数）"""
    return DateNormalizer.normalize(date_str, fmt_hint)
