"""
建设银行 OFD 2023版 XBRL 解析器

解析 CCB OFD 格式（2023版）的 XBRL 银行流水
识别特征：OFD包内XBRL命名空间 = "http://xbrl.mof.gov.cn/taxonomy/2023-05-15/bkrs"
"""
import zipfile
import re
from typing import List, Optional
from datetime import datetime

from ..base import BaseBankParser, BankTransaction
from ..comm.date import normalize_date


class CCB_OFD_2023_Parser(BaseBankParser):
    """CCB OFD 2023版 XBRL 解析器"""

    BANK_CODE = 'ccb'
    # 2023版 XBRL 命名空间
    NS_2023 = 'http://xbrl.mof.gov.cn/taxonomy/2023-05-15/bkrs'

    def detect(self, file_path: str) -> bool:
        """
        检测是否为 CCB OFD 2023版 XBRL 格式

        支持 ZIP 包裹的 OFD 和独立的 OFD 文件/目录
        """
        import os
        import zipfile
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in ('.ofd', '.zip'):
            return False

        # 方式1：作为 ZIP 包裹的 OFD 处理
        try:
            with zipfile.ZipFile(file_path, 'r') as zf:
                for name in zf.namelist():
                    if name.lower().endswith(('.xml', '.xbrl')) and 'attachs' in name.lower():
                        try:
                            content = zf.read(name).decode('utf-8', errors='ignore')
                            if self.NS_2023 in content:
                                return True
                        except Exception:
                            continue
        except Exception:
            pass

        # 方式2：作为独立 OFD 文件（OFD.xml）或解压目录处理
        try:
            # 判断是否是 OFD 入口文件
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                header = f.read(256)
            if '<ofd:OFD' not in header and '<ofd:' not in header:
                return False

            # 在 OFD 目录树中查找 bkrs XBRL 文件
            search_dir = os.path.dirname(os.path.abspath(file_path))
            for parent in [search_dir, os.path.dirname(search_dir), os.path.dirname(os.path.dirname(search_dir))]:
                # 在 Doc_0/Attachs 中查找 bkrs*.xml
                for fn in os.listdir(attachs_dir):
                    if fn.endswith('.xml') and fn.startswith('bkrs'):
                        bkrs_xml = os.path.join(attachs_dir, fn)
                        try:
                            with open(bkrs_xml, 'r', encoding='utf-8', errors='ignore') as f:
                                content = f.read(8192)
                            if self.NS_2023 in content:
                                return True
                        except Exception:
                            continue
        except Exception:
            pass

        return False

    def parse(self, file_path: str) -> List[BankTransaction]:
        """
        解析 CCB OFD 2023版 XBRL 格式

        Args:
            file_path: OFD/XBRL 文件路径

        Returns:
            BankTransaction 列表
        """
        results: List[BankTransaction] = []

        try:
            with zipfile.ZipFile(file_path, 'r') as zf:
                # 查找所有 XBRL XML 文件
                xbrl_files = [
                    name for name in zf.namelist()
                    if name.lower().endswith('.xml') and 'attachs' in name.lower()
                    and name != 'Doc_0/Attachs/Attachments.xml'
                ]

                for xbrl_file in xbrl_files:
                    try:
                        content = zf.read(xbrl_file).decode('utf-8', errors='ignore')
                        if self.NS_2023 not in content:
                            continue
                        txns = self._parse_xbrl_content(content, file_path)
                        results.extend(txns)
                    except Exception:
                        continue
        except Exception:
            pass

        return results

    def _parse_xbrl_content(self, content: str, file_path: str) -> List[BankTransaction]:
        """解析 XBRL XML 内容"""
        transactions: List[BankTransaction] = []

        # 提取所有 InformationOfReconcileDetailsTuple
        tuple_pattern = re.compile(
            r'<bkrs:InformationOfReconcileDetailsTuple>(.*?)</bkrs:InformationOfReconcileDetailsTuple>',
            re.DOTALL
        )

        for match in tuple_pattern.finditer(content):
            tuple_xml = match.group(1)
            txn = self._parse_tuple(tuple_xml, file_path)
            if txn:
                transactions.append(txn)

        return transactions

    def _parse_tuple(self, tuple_xml: str, file_path: str) -> Optional[BankTransaction]:
        """解析单个 XBRL Tuple"""
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

        # 收支方向：ICD=0→expense，ICD=1→income
        icd_match = re.search(
            r'<bkrs:IdentificationOfCreditOrDebit[^>]*>([^<]+)<',
            tuple_xml
        )
        icd = icd_match.group(1).strip() if icd_match else '0'
        direction = 'income' if icd == '1' else 'expense'

        # 交易金额（从 TransactionAmount 取，去掉符号）
        amt_match = re.search(
            r'<bkrs:TransactionAmount[^>]*>([^<]+)<',
            tuple_xml
        )
        if not amt_match:
            return None
        amount_str = amt_match.group(1).strip()
        try:
            amount = abs(float(amount_str))
        except ValueError:
            return None
        if amount == 0:
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

        # 流水号（AccDtlSn 或 BusinessSerialNumber）
        serial_match = re.search(
            r'<bkrs:AccDtlSn[^>]*>([^<]+)<',
            tuple_xml
        )
        serial_no = serial_match.group(1).strip() if serial_match else ''

        # 如果没有 AccDtlSn，尝试 BusinessSerialNumber
        if not serial_no:
            bs_match = re.search(
                r'<bkrs:BusinessSerialNumber[^>]*>([^<]+)<',
                tuple_xml
            )
            serial_no = bs_match.group(1).strip() if bs_match else ''

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
