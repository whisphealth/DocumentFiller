from DocumentFiller.DocFillerFactory import DocumentFillerFactory
from DocumentFiller.DocxFiller import DocxFiller
from DocumentFiller.OdtFiller import OdtFiller
from DocumentFiller.XlsxFiller import XlsxFiller


def test_document_filler_factory():
    dff = DocumentFillerFactory()

    dff_docx = dff.inst("docx")

    assert isinstance(dff_docx, DocxFiller)

    dff_odf = dff.inst("odt")

    assert isinstance(dff_odf, OdtFiller)

    dff_xlsx = dff.inst("xlsx")

    assert isinstance(dff_xlsx, XlsxFiller)
