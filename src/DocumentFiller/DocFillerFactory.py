from .DocxFiller import DocxFiller
from .XlsxFiller import XlsxFiller
from .OdtFiller import OdtFiller


class DocumentFillerFactory:

    DEBUG = PDF = PDF_ONLY = False

    def __init__(
        self, debug: bool = False, pdf: bool = False, pdfonly: bool = False
    ):
        self.DEBUG = debug
        self.PDF = pdf
        self.PDF_ONLY = pdfonly
        pass

    def inst(self, extention: str):
        if self.__ext_equal(extention, "docx"):
            return DocxFiller(self.DEBUG, self.PDF, self.PDF_ONLY)
        elif self.__ext_equal(extention, "xlsx"):
            return XlsxFiller(self.DEBUG, self.PDF, self.PDF_ONLY)
        elif self.__ext_equal(extention, "odt"):
            return OdtFiller(self.DEBUG, self.PDF, self.PDF_ONLY)
        elif self.DEBUG:
            print(f"INFO - File format not supported so ignored: {extention}")

    def __ext_equal(self, ext1: str, ext2: str) -> bool:
        return ext1 in [ext2, "." + ext2]
