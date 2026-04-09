"""
模糊匹配工具

使用 rapidfuzz 进行高效的模糊字符串匹配
"""
from typing import List, Tuple, Optional, Dict, Any, Union
from rapidfuzz import fuzz, process


class FuzzyMatcher:
    """模糊匹配工具类"""

    def __init__(self, threshold: int = 80):
        """
        初始化模糊匹配器

        Args:
            threshold: 匹配阈值（0-100），低于此分数的不返回
        """
        self.threshold = threshold

    def match_one(self, query: str, choices: List[str], score_cutoff: Optional[int] = None) -> Tuple[Optional[str], Optional[int]]:
        """
        从候选列表中找出一个最佳匹配

        Args:
            query: 查询字符串
            choices: 候选字符串列表
            score_cutoff: 最低得分阈值（覆盖默认值）

        Returns:
            (最佳匹配字符串, 匹配得分) 或 (None, None) 如果没有匹配
        """
        if not query or not choices:
            return None, None

        threshold = score_cutoff if score_cutoff is not None else self.threshold

        result = process.extractOne(
            query,
            choices,
            scorer=fuzz.ratio,
            score_cutoff=threshold
        )

        if result:
            return (result[0], int(result[1]))
        return None, None

    def match_multi(self, query: str, choices: List[str], limit: int = 5, score_cutoff: Optional[int] = None) -> List[Tuple[str, int]]:
        """
        从候选列表中找出多个最佳匹配

        Args:
            query: 查询字符串
            choices: 候选字符串列表
            limit: 返回的最大结果数
            score_cutoff: 最低得分阈值

        Returns:
            [(匹配字符串, 匹配得分), ...] 列表，按得分降序排列
        """
        if not query or not choices:
            return []

        threshold = score_cutoff if score_cutoff is not None else self.threshold

        results = process.extract(
            query,
            choices,
            scorer=fuzz.ratio,
            limit=limit,
            score_cutoff=threshold
        )

        return [(r[0], int(r[1])) for r in results]

    def match_with_type(self, query: str, choices: List[Dict[str, Any]], key: str, limit: int = 5, score_cutoff: Optional[int] = None) -> List[Tuple[Dict[str, Any], int]]:
        """
        从字典列表中模糊匹配

        Args:
            query: 查询字符串
            choices: 字典列表
            key: 要匹配的字段名
            limit: 返回的最大结果数
            score_cutoff: 最低得分阈值

        Returns:
            [(匹配的字典, 匹配得分), ...] 列表
        """
        if not query or not choices:
            return []

        # 提取要匹配的字段值列表
        values = [str(item.get(key, '')) for item in choices]

        # 执行匹配
        threshold = score_cutoff if score_cutoff is not None else self.threshold
        results = process.extract(
            query,
            values,
            scorer=fuzz.ratio,
            limit=limit,
            score_cutoff=threshold
        )

        # 映射回原始字典
        matched = []
        for value, score, index in results:
            matched.append((choices[index], int(score)))

        return matched

    @staticmethod
    def similarity(s1: str, s2: str) -> int:
        """
        计算两个字符串的相似度得分

        Args:
            s1: 字符串1
            s2: 字符串2

        Returns:
            相似度得分（0-100）
        """
        if not s1 or not s2:
            return 0
        return int(fuzz.ratio(str(s1), str(s2)))

    @staticmethod
    def partial_similarity(s1: str, s2: str) -> int:
        """
        计算两个字符串的部分相似度（子串匹配）

        Args:
            s1: 字符串1
            s2: 字符串2

        Returns:
            部分相似度得分（0-100）
        """
        if not s1 or not s2:
            return 0
        return int(fuzz.partial_ratio(str(s1), str(s2)))

    @staticmethod
    def token_similarity(s1: str, s2: str) -> int:
        """
        计算两个字符串的分词相似度（考虑词序）

        Args:
            s1: 字符串1
            s2: 字符串2

        Returns:
            分词相似度得分（0-100）
        """
        if not s1 or not s2:
            return 0
        return int(fuzz.token_sort_ratio(str(s1), str(s2)))


def is_fuzzy_match(str_a: str, str_b: str, threshold: int = 50) -> bool:
    """
    判断两个字符串是否模糊匹配（partial_ratio >= threshold）

    Args:
        str_a: 字符串A
        str_b: 字符串B
        threshold: 匹配阈值（默认50）

    Returns:
        True if fuzzy match, False otherwise
    """
    if not str_a or not str_b:
        return False
    return fuzz.partial_ratio(str_a, str_b) >= threshold


def best_fuzzy_match(target: str, candidates: List[str]) -> Tuple[Optional[str], int]:
    """
    在候选列表中找到与目标字符串最佳模糊匹配的项

    Args:
        target: 目标字符串
        candidates: 候选字符串列表

    Returns:
        (最佳匹配字符串, 最高得分) tuple，若候选为空则返回 (None, 0)
    """
    if not candidates:
        return None, 0

    best_score = 0
    best_match = None
    for c in candidates:
        score = fuzz.partial_ratio(target, c)
        if score > best_score:
            best_score = score
            best_match = c
    return best_match, best_score
