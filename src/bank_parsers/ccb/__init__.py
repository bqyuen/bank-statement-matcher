"""
建设银行（CCB）解析器包
"""
from .ccb_pdf import CCB_PDF_Parser
from .ccb_ofd_2023 import CCB_OFD_2023_Parser
from .ccb_ofd_2021 import CCB_OFD_2021_Parser
from .ccb_image_pdf import CCB_Image_PDF_Parser

__all__ = [
    'CCB_PDF_Parser',
    'CCB_OFD_2023_Parser',
    'CCB_OFD_2021_Parser',
    'CCB_Image_PDF_Parser',
]
