from zipfile import ZipFile
from os.path import exists
from subprocess import run
from .DocFillerFamily import DocumentFillerFamilly
from re import findall, search, sub
from lxml import etree


class XlsxFiller(DocumentFillerFamilly):

    # Flags can't be changed for now
    BEFORE_FLAG = "{{"
    AFTER_FLAG = "}}"
    SEPARATOR = "_"  # MUST BE ANYTHING OTHER THAN A NUMBER !!!!!
    DEBUG = PDF = PDF_ONLY = False

    def __init__(
        self, debug: bool = False, pdf: bool = False, pdf_only: bool = False
    ):
        self.DEBUG = debug
        self.PDF = pdf
        self.PDF_ONLY = pdf_only

    def __replace_text(self, text: bytes, tags: dict) -> bytes:
        try:
            root = etree.fromstring(text)
        except Exception:
            raise Exception("The cleaning process broke the xml")

        # child_to_parent = {c: p for p in root.iter() for c in p}

        for i in root.iter():
            # If there is no flags, no need to check for them
            if not search(
                self.BEFORE_FLAG + ".*?" + self.AFTER_FLAG, str(i.text)
            ):
                continue

            if search(
                self.BEFORE_FLAG
                + r".*?"
                + self.SEPARATOR
                + r"\d+?"
                + self.AFTER_FLAG,
                i.text,
            ):
                i.text = self.__split_replace(i.text, tags)
            else:
                i.text = self.__simple_replace(i.text, tags)

        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'.encode(
                "utf-8"
            )
            + etree.tostring(root, encoding="utf-8")
        )

    def __simple_replace(self, text: str, values: dict) -> str:
        """
        Look for the flags and replace them with there corresponding values
        """
        for key, value in values.items():
            simple_key = self.BEFORE_FLAG + key + self.AFTER_FLAG

            for flag in findall(simple_key, text):
                if self.DEBUG:
                    print(f"KEY - Simple Replaced {flag} with {value}")
                text = text.replace(flag, str(value))

        return text

    def __split_replace(self, text: str, values: dict) -> str:
        """
        Look for the splitted flags ( {{%FLAG_01%}}, {{%FLAG_02%}} ) and
        replace them with the correct value
        """
        for key, value in values.items():
            split_key = (
                self.BEFORE_FLAG
                + key
                + self.SEPARATOR
                + r"\d+?"
                + self.AFTER_FLAG
            )

            if search(split_key, text):
                flags = findall(split_key, text)
                for flag in flags:
                    nb = int(
                        search(self.SEPARATOR + r"\d+", flag)
                        .group(0)
                        .replace(self.SEPARATOR, "")
                    )
                    value_2 = str(value)[nb] if nb < len(str(value)) else ""
                    if self.DEBUG:
                        print(f"KEY - Split Replaced {flag} with {value_2}")
                    text = text.replace(flag, value_2)
        return text

    def fill_document(self, src_path: str, values: dict, dest_path: str):
        """
        Function that fill a document by replacing flags and put the result in
        an other directory
        :param path: (String) Path to the document that needs to be filled
        :param values: (Dictionary) Dictionary containing the flags (keys) and
        the values (items)
        :param destpath: (String) Path to the document that will be created /
        replaced
        """

        # Input param check
        if not isinstance(src_path, str):
            raise TypeError("Source path provided is not type String")
        if not isinstance(values, dict):
            raise TypeError("Values provided are not type Dict")
        if not isinstance(dest_path, str):
            raise TypeError("Destination path provided is not type String")
        if src_path == dest_path:
            raise ValueError(
                "The source and destination path must be different"
            )
        if not exists(src_path):
            raise FileNotFoundError("The source file does not exist")

        with ZipFile(src_path) as in_zip, ZipFile(dest_path, "w") as out_zip:
            # For each files in the input zip file
            for in_zip_info in in_zip.infolist():
                with in_zip.open(in_zip_info) as in_file:
                    # Try reading it
                    try:
                        content = in_file.read()
                    except Exception:
                        raise Exception(
                            "Can't read properly the file "
                            f"{in_zip_info.filename} in {src_path}"
                        )

                    if self.DEBUG:
                        print(f"INFO - Read {in_zip_info.filename}")

                    # Replace the text (to fill the document)
                    if "xl/sharedStrings.xml" == in_zip_info.filename:
                        if self.DEBUG:
                            print(f"INFO - Modifying {in_zip_info.filename}")
                        content = self.__replace_text(content, values)

                    # Try writing to the output file
                    try:
                        out_zip.writestr(in_zip_info.filename, content)
                    except Exception:
                        raise Exception(
                            f"Can't write the file {in_zip_info.filename} "
                            f"in {src_path}"
                        )

                    if self.DEBUG:
                        print(f"INFO - Wrote {in_zip_info.filename}")

        # PDF Stuff
        if self.PDF or self.PDF_ONLY:
            run(
                [
                    "localc",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    sub(r"[^\/]+?\.xlsx", "", dest_path),
                    dest_path,
                ]
            )
        if self.PDF_ONLY:
            run(["rm", dest_path])
