from .DocxFiller import DocxFiller
from .XlsxFiller import XlsxFiller
from .OdtFiller import OdtFiller

from threading import Lock


class DocumentFillerFactory:

    DEBUG = PDF = PDF_ONLY = False
    LOCK = None

    def __init__(
        self,
        debug: bool = False,
        pdf: bool = False,
        pdfonly: bool = False,
        lock: Lock = None,
    ):
        self.DEBUG = debug
        self.PDF = pdf
        self.PDF_ONLY = pdfonly
        self.LOCK = lock

    def inst(self, extention: str):
        if self.__ext_equal(extention, "docx"):
            return DocxFiller(self.DEBUG, self.PDF, self.PDF_ONLY, self.LOCK)
        elif self.__ext_equal(extention, "xlsx"):
            return XlsxFiller(self.DEBUG, self.PDF, self.PDF_ONLY, self.LOCK)
        elif self.__ext_equal(extention, "odt"):
            return OdtFiller(self.DEBUG, self.PDF, self.PDF_ONLY, self.LOCK)
        elif self.DEBUG:
            print(f"INFO - File format not supported so ignored: {extention}")

    def __ext_equal(self, ext1: str, ext2: str) -> bool:
        ext2 = ext2 if ext2[0] != "." else ext2[1:]
        return ext1 in [ext2, "." + ext2]
