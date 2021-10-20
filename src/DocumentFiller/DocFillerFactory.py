from .DocxFiller import DocxFiller
from .XlsxFiller import XlsxFiller


class DocumentFillerFactory:

    DEBUG = PDF = PDF_ONLY = False

    def __init__(
        self, debug: bool = False, pdf: bool = False, pdfonly: bool = False
    ):
        self.DEBUG = debug
        self.PDF = pdf
        self.PDF_ONLY = pdfonly
        pass

    def inst(self, format: str):
        if format == "docx":
            return DocxFiller(self.DEBUG, self.PDF, self.PDF_ONLY)
        elif format == "xlsx":
            return XlsxFiller(self.DEBUG, self.PDF, self.PDF_ONLY)
        else:
            if self.DEBUG:
                print(f"INFO - File format not supported so ignored: {format}")
