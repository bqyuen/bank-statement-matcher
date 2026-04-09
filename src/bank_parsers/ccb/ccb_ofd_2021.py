"""
建设银行 OFD 2021版 XBRL 解析器

解析 CCB OFD 格式（2021版）的 XBRL 银行流水
识别特征：OFD包内XBRL命名空间 = "http://xbrl.mof.gov.cn/taxonomy/2021-11-30/bkrs"
特殊逻辑：每2个Tuple成对，取非0金额的那个
"""
import zipfile
import re
from typing import List, Optional
from datetime import datetime
import xml.etree.ElementTree as ET

from ..base import BaseBankParser, BankTransaction
from ..comm.amount import normalize_amount
from ..comm.date import normalize_date


class CCB_OFD_2021_Parser(BaseBankParser):
    """CCB OFD 2021版 XBRL 解析器"""

    BANK_CODE = 'ccb'
    # 2021版 XBRL 命名空间
    NS_2021 = 'http://xbrl.mof.gov.cn/taxonomy/2021-11-30/bkrs'

    def detect(self, file_path: str) -> bool:
        """
        检测是否为 CCB OFD 2021版 XBRL 格式

        通过 XBRL 命名空间判断
        """
        import os
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in ('.ofd', '.zip', ''):
            # 也可能是已解压的目录
            pass

        try:
            # 尝试作为 zip 文件打开（OFD 本质是 zip）
            with zipfile.ZipFile(file_path, 'r') as zf:
                names = zf.namelist()
            bank = self._detect_bank_from_zip(file_path)
            return bank == 'ccb_2021'
        except Exception:
            # 可能是已解压的目录
            return self._detect_from_directory(file_path)

    def _detect_bank_from_zip(self, zip_path: str) -> str:
        """从 ZIP 包中检测银行类型和版本"""
        try:
            with zipfile.ZipFile(zip_path, 'r') as zf:
                for name in zf.namelist():
                    if name.lower().endswith('.xml') and 'attachs' in name.lower():
                        try:
                            content = zf.read(name).decode('utf-8', errors='ignore')
                            if self.NS_2021 in content:
                                return 'ccb_2021'
                            # 2023 命名空间
                            if 'http://xbrl.mof.gov.cn/taxonomy/2023-05-15/bkrs' in content:
                                return 'ccb_2023'
                        except Exception:
                            continue
        except Exception:
            pass
        return 'unknown'

    def _detect_from_directory(self, file_or_dir_path: str) -> bool:
        """从已解压的 OFD 文件或目录中检测"""
        import os
        # 如果传入的是文件（OFD.xml/document.ofd），找到其所在目录
        if os.path.isfile(file_or_dir_path):
            search_dir = os.path.dirname(os.path.abspath(file_or_dir_path))
        else:
            search_dir = os.path.abspath(file_or_dir_path)

        # OFD 包结构：OFD.xml 同级有 Doc_0/Attachs/
        for parent in [search_dir, os.path.dirname(search_dir)]:
            if not parent or not os.path.isdir(parent):
                continue
            attachs_dir = os.path.join(parent, 'Doc_0', 'Attachs')
            if not os.path.isdir(attachs_dir):
                continue
            try:
                for fname in os.listdir(attachs_dir):
                    if fname.endswith('.xml') and fname != 'Attachments.xml':
                        fpath = os.path.join(attachs_dir, fname)
                        with open(fpath, 'r', encoding='utf-8', errors='ignore') as f:
                            content = f.read()
                        if self.NS_2021 in content:
                            return True
            except Exception:
                pass
        return False

    def parse(self, file_path: str) -> List[BankTransaction]:
        """
        解析 CCB OFD 2021版 XBRL 格式

        Args:
            file_path: OFD 文件或解压目录路径

        Returns:
            BankTransaction 列表
        """
        results: List[BankTransaction] = []
        tuple_count = 0

        try:
            # 尝试作为 zip 文件打开
            with zipfile.ZipFile(file_path, 'r') as zf:
                xbrl_files = [
                    name for name in zf.namelist()
                    if name.lower().endswith('.xml') and 'attachs' in name.lower()
                    and name != 'Doc_0/Attachs/Attachments.xml'
                ]
                for xbrl_file in xbrl_files:
                    try:
                        content = zf.read(xbrl_file).decode('utf-8', errors='ignore')
                        if self.NS_2021 not in content:
                            continue
                        tuple_count += self._count_tuples(content)
                        txns = self._parse_xbrl_content(content, file_path)
                        results.extend(txns)
                    except Exception:
                        continue

            # 部分 OFD 的 XBRL 附件金额字段不完整（大量为0），
            # 此时回退到分页文本解析以恢复完整记录。
            if self._needs_page_fallback(results_count=len(results), tuple_count=tuple_count):
                page_txns = self._parse_from_page_content_zip(file_path)
                if len(page_txns) > len(results):
                    return page_txns
        except Exception:
            # 尝试作为解压目录处理
            results = self._parse_from_directory(file_path)

        return results

    def _count_tuples(self, content: str) -> int:
        return len(re.findall(
            r'<bkrs:InformationOfReconcileDetailsTuple[^>]*>.*?</bkrs:InformationOfReconcileDetailsTuple>',
            content,
            re.DOTALL
        ))

    def _needs_page_fallback(self, results_count: int, tuple_count: int) -> bool:
        if results_count == 0:
            return True
        if tuple_count >= 20 and results_count <= int(tuple_count * 0.6):
            return True
        return False

    def _parse_from_directory(self, file_or_dir_path: str) -> List[BankTransaction]:
        """从已解压的 OFD 文件或目录解析"""
        import os
        results: List[BankTransaction] = []
        # 如果传入的是文件，找到其所在目录
        if os.path.isfile(file_or_dir_path):
            search_dir = os.path.dirname(os.path.abspath(file_or_dir_path))
        else:
            search_dir = os.path.abspath(file_or_dir_path)

        # OFD 包结构：Doc_0/Attachs/ 与 OFD.xml 同级
        for parent in [search_dir, os.path.dirname(search_dir)]:
            if not parent or not os.path.isdir(parent):
                continue
            attachs_dir = os.path.join(parent, 'Doc_0', 'Attachs')
            if not os.path.isdir(attachs_dir):
                continue
            try:
                for fname in os.listdir(attachs_dir):
                    if fname.endswith('.xml') and fname != 'Attachments.xml':
                        fpath = os.path.join(attachs_dir, fname)
                        with open(fpath, 'r', encoding='utf-8', errors='ignore') as f:
                            content = f.read()
                        if self.NS_2021 not in content:
                            continue
                        txns = self._parse_xbrl_content(content, file_or_dir_path)
                        results.extend(txns)
            except Exception:
                pass
            if results:
                break
        return results

    def _parse_from_page_content_zip(self, file_path: str) -> List[BankTransaction]:
        """从 OFD 分页 Content.xml 提取交易明细（兜底方案）"""
        try:
            with zipfile.ZipFile(file_path, 'r') as zf:
                page_files = [
                    n for n in zf.namelist()
                    if re.match(r'Doc_0/Pages/Page_\d+/Content\.xml$', n)
                ]
                page_files.sort(key=self._page_sort_key)

                tokens: List[str] = []
                for page_file in page_files:
                    content = zf.read(page_file).decode('utf-8', errors='ignore')
                    tokens.extend(self._extract_text_tokens(content))
        except Exception:
            return []

        return self._build_transactions_from_page_tokens(tokens, file_path)

    def _page_sort_key(self, page_path: str) -> int:
        m = re.search(r'Page_(\d+)/Content\.xml$', page_path)
        return int(m.group(1)) if m else 0

    def _extract_text_tokens(self, page_content: str) -> List[str]:
        tokens: List[str] = []
        try:
            root = ET.fromstring(page_content)
        except Exception:
            return tokens

        ns = {'ofd': 'http://www.ofdspec.org'}
        for text_node in root.findall('.//ofd:TextCode', ns):
            txt = (text_node.text or '').strip()
            if txt:
                tokens.append(txt)
        return tokens

    def _build_transactions_from_page_tokens(self, tokens: List[str], file_path: str) -> List[BankTransaction]:
        if not tokens:
            return []

        date_pattern = re.compile(r'20\d{2}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])$')
        money_pattern = re.compile(r'-?\d{1,3}(?:,\d{3})*\.\d{2}$')

        date_indices = [i for i, tok in enumerate(tokens) if date_pattern.fullmatch(tok)]
        transactions: List[BankTransaction] = []

        for i, start in enumerate(date_indices):
            end = date_indices[i + 1] if i + 1 < len(date_indices) else len(tokens)
            seg = tokens[start:end]
            if not seg:
                continue

            date_token = seg[0]
            txn_date = normalize_date(date_token, '%Y%m%d')
            if txn_date is None:
                txn_date = normalize_date(date_token)
            if txn_date is None:
                continue

            money_tokens = [tok for tok in seg if money_pattern.fullmatch(tok)]
            if len(money_tokens) < 2:
                continue

            # 标准行是 [交易金额, 余额]，最后一行可能额外带页尾汇总金额，
            # 因此优先取首对金额。
            amount_raw = money_tokens[0]
            balance_raw = money_tokens[1]

            amount_val = normalize_amount(amount_raw)
            if amount_val == 0:
                continue
            direction = 'expense' if amount_raw.startswith('-') else 'income'
            amount = abs(amount_val)
            balance = normalize_amount(balance_raw)

            summary = seg[1] if len(seg) > 1 else ''
            counterparty = self._extract_counterparty(seg)
            counterparty_account = self._extract_counterparty_account(seg)
            serial_no = self._extract_serial(seg, amount_raw)

            transactions.append(BankTransaction(
                file=file_path,
                bank_type=self.BANK_CODE,
                date=txn_date,
                amount=amount,
                direction=direction,
                counterparty=counterparty,
                counterparty_account=counterparty_account,
                counterparty_bank='',
                summary=summary,
                serial_no=serial_no,
                balance=balance,
            ))

        return transactions

    def _extract_counterparty(self, seg: List[str]) -> str:
        ignore = {'转取', '转存', '汇款退回', '批量扣费'}
        for tok in seg[2:]:
            if tok in ignore:
                continue
            if re.fullmatch(r'\d+', tok):
                continue
            if re.fullmatch(r'-?\d{1,3}(?:,\d{3})*\.\d{2}', tok):
                continue
            if tok.startswith('第 ') and '页' in tok:
                continue
            if '中国建设银行' in tok or '中国农业银行' in tok:
                continue
            if re.search(r'[\u4e00-\u9fffA-Za-z]', tok):
                return tok
        return ''

    def _extract_counterparty_account(self, seg: List[str]) -> str:
        for tok in seg:
            if re.fullmatch(r'\d{10,}', tok):
                return tok
        return ''

    def _extract_serial(self, seg: List[str], amount_raw: str) -> str:
        for idx, tok in enumerate(seg):
            if tok == amount_raw:
                # 交易金额前最近的 6~12 位数字通常是流水号
                for j in range(idx - 1, -1, -1):
                    cand = seg[j]
                    if re.fullmatch(r'\d{6,12}', cand):
                        return cand
                break
        return ''

    def _parse_xbrl_content(self, content: str, file_path: str) -> List[BankTransaction]:
        """
        解析 XBRL XML 内容

        2021版特殊逻辑：
        - 每2个Tuple成对出现（借贷双方）
        - 取非0金额的那个作为实际交易
        - Tuple[N] ICD=0, Amt=0 → 无效占位
        - Tuple[N+1] ICD=0, Amt=-780 → 实际支出780元（direction='expense'，amount=780）
        - ICD=1 的为收入条目（金额通常在对面占位条中）
        """
        transactions: List[BankTransaction] = []

        # 提取所有 InformationOfReconcileDetailsTuple
        tuple_pattern = re.compile(
            r'<bkrs:InformationOfReconcileDetailsTuple[^>]*>(.*?)</bkrs:InformationOfReconcileDetailsTuple>',
            re.DOTALL
        )

        raw_tuples = []
        for match in tuple_pattern.finditer(content):
            raw_tuples.append(match.group(0))

        i = 0
        while i < len(raw_tuples):
            tuple_xml = raw_tuples[i]

            # 解析当前 Tuple 的 ICD 和金额
            icd_val, amt_val, is_nil = self._extract_icd_and_amount(tuple_xml)

            if icd_val == '0':
                # 借方（支出）
                if amt_val != 0:
                    # 非零金额 → 实际支出交易
                    txn = self._build_transaction(tuple_xml, file_path, 'expense', abs(amt_val))
                    if txn:
                        transactions.append(txn)
                    i += 1
                else:
                    # Amt=0 → 可能是占位，检查下一个
                    if i + 1 < len(raw_tuples):
                        next_xml = raw_tuples[i + 1]
                        next_icd, next_amt, next_nil = self._extract_icd_and_amount(next_xml)
                        if next_amt != 0:
                            # 下一条是非零 → 这是实际交易
                            txn = self._build_transaction(next_xml, file_path, 'expense', abs(next_amt))
                            if txn:
                                transactions.append(txn)
                            i += 2
                        else:
                            # 两条都是零，跳过
                            i += 1
                    else:
                        i += 1
            elif icd_val == '1':
                # 贷方（收入）—— 金额在对面占位条中
                if amt_val != 0:
                    # 非零金额 → 实际收入交易
                    txn = self._build_transaction(tuple_xml, file_path, 'income', abs(amt_val))
                    if txn:
                        transactions.append(txn)
                    i += 1
                else:
                    # Amt=0 → 可能是占位，检查下一个
                    if i + 1 < len(raw_tuples):
                        next_xml = raw_tuples[i + 1]
                        next_icd, next_amt, next_nil = self._extract_icd_and_amount(next_xml)
                        if next_amt != 0:
                            txn = self._build_transaction(next_xml, file_path, 'income', abs(next_amt))
                            if txn:
                                transactions.append(txn)
                            i += 2
                        else:
                            i += 1
                    else:
                        i += 1
            else:
                # 未知 ICD，跳过
                i += 1

        return transactions

    def _extract_icd_and_amount(self, tuple_xml: str) -> tuple[str, float, bool]:
        """
        从 Tuple XML 中提取 ICD 值和金额

        Returns:
            (icd_str, amount_float, is_nil)
        """
        # ICD
        icd_match = re.search(
            r'<bkrs:IdentificationOfCreditOrDebit[^>]*>([^<]*)<',
            tuple_xml
        )
        icd_str = icd_match.group(1).strip() if icd_match else '0'

        # 金额
        amt_match = re.search(
            r'<bkrs:TransactionAmount[^>]*>([^<]+)<',
            tuple_xml
        )
        if not amt_match:
            return icd_str, 0.0, True

        is_nil = 'xsi:nil="true"' in tuple_xml[amt_match.start():amt_match.end() + 20]

        if is_nil:
            return icd_str, 0.0, True

        try:
            amount = float(amt_match.group(1).strip())
        except ValueError:
            amount = 0.0

        return icd_str, amount, False

    def _build_transaction(
        self,
        tuple_xml: str,
        file_path: str,
        direction: str,
        amount: float
    ) -> Optional[BankTransaction]:
        """从 Tuple XML 构建交易记录"""
        # 记账日期
        date_match = re.search(
            r'<bkrs:DateOfBookkeeping[^>]*>([^<]+)<',
            tuple_xml
        )
        if not date_match:
            return None
        date_str = date_match.group(1).strip()
        date = normalize_date(date_str, '%Y-%m-%d')
        if date is None:
            date = normalize_date(date_str)
        if date is None:
            return None

        # 余额
        bal_match = re.search(
            r'<bkrs:AccountBalance[^>]*>([^<]+)<',
            tuple_xml
        )
        balance = 0.0
        if bal_match:
            try:
                balance = abs(float(bal_match.group(1).strip()))
            except ValueError:
                pass

        # 对方户名
        name_match = re.search(
            r'<bkrs:NameOfCounterparty[^>]*>([^<]*)<',
            tuple_xml
        )
        counterparty = name_match.group(1).strip() if name_match else ''

        # 对方账号
        acct_match = re.search(
            r'<bkrs:AccountOfCounterparty[^>]*>([^<]*)<',
            tuple_xml
        )
        counterparty_account = acct_match.group(1).strip() if acct_match else ''

        # 对方开户机构
        bank_match = re.search(
            r'<bkrs:DepositoryBankOfCounterparty[^>]*>([^<]*)<',
            tuple_xml
        )
        counterparty_bank = bank_match.group(1).strip() if bank_match else ''

        # 摘要
        summary_match = re.search(
            r'<bkrs:NotesOfBankElectronicReceipt[^>]*>([^<]*)<',
            tuple_xml
        )
        summary = summary_match.group(1).strip() if summary_match else ''

        # 流水号
        serial_match = re.search(
            r'<bkrs:JournalAccountOfBookkeeping[^>]*>([^<]+)<',
            tuple_xml
        )
        serial_no = serial_match.group(1).strip() if serial_match else ''

        return BankTransaction(
            file=file_path,
            bank_type=self.BANK_CODE,
            date=date,
            amount=amount,
            direction=direction,
            counterparty=counterparty,
            counterparty_account=counterparty_account,
            counterparty_bank=counterparty_bank,
            summary=summary,
            serial_no=serial_no,
            balance=balance,
        )
