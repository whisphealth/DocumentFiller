from re import sub, findall, search
from os.path import exists, splitext, split
from subprocess import run
from zipfile import ZipFile

from odf.style import Style, TextProperties
from odf.opendocument import OpenDocumentText, load
from odf import text as odfText
from odf.element import Element, Text

from .DocFillerFamily import DocumentFillerFamilly


class OdtFiller(DocumentFillerFamilly):

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

    def __recursive_get_text(self, ele) -> list:
        """
        Parcour recursively the childs of an element and return a list of
        childs that are Text and contain a flag
        """
        if isinstance(ele, Text):
            if search(r"{{[^{}]+?}}", ele.data):
                return ele
        elif ele:
            return list(map(self.__recursive_get_text, ele.childNodes))

    def __recursive_get_parents(self, ele, child_list) -> dict:
        """
        Parcour recursively the childs of an element and return a dictionary
        of all the parents of each childs
        """
        parents = {}
        child = [i for i in ele.childNodes if i in child_list]
        future_parents = [i for i in ele.childNodes if "childNodes" in dir(i)]

        for i in child:
            parents[i] = ele
        for i in future_parents:
            for k, j in self.__recursive_get_parents(i, child_list).items():
                parents[k] = j

        return parents

    def __flatten(self, S) -> list:
        """
        Flatten a list recursively
        """
        if S == []:
            return S
        if isinstance(S[0], list):
            return self.__flatten(S[0]) + self.__flatten(S[1:])
        return S[:1] + self.__flatten(S[1:])

    def __replace_text(
        self, doc: OpenDocumentText, tags: dict
    ) -> OpenDocumentText:
        """
        Function that replace le flags in text with the values associated to
        the flag in values
        :param text: (String) Xml input text that contain the flags to replace
        :param values: (Dictionary) Dictionary containing the flags (keys) and
        the values (items)
        :return: (String) Xml Text with all the flags replaced
        """

        under_tags, check_tags, simple_tags = self.__split_tags(tags)

        elements = doc.getElementsByType(odfText.H) + doc.getElementsByType(
            odfText.P
        )

        flagged_elements = [
            i for i in elements if search(r"{{[^{}]+?}}", str(i))
        ]

        flagged_text_elements = self.__flatten(
            list(map(self.__recursive_get_text, flagged_elements))
        )
        flagged_text_elements = [i for i in flagged_text_elements if i]

        flagged_parent_text_elements = {}
        for i in flagged_elements:
            for j, k in self.__recursive_get_parents(
                i, flagged_text_elements
            ).items():
                flagged_parent_text_elements[j] = k

        for text in flagged_text_elements:
            # If there is no flags, no need to check for them
            if not search(
                self.BEFORE_FLAG + ".*?" + self.AFTER_FLAG, text.data
            ):
                continue

            if search(
                self.BEFORE_FLAG
                + "UNDER"
                + self.SEPARATOR
                + ".*?"
                + self.AFTER_FLAG,
                text.data,
            ):
                parent_text = flagged_parent_text_elements[text]
                parent_text = self.__underline(
                    text, parent_text, under_tags, doc
                )
            elif search(
                self.BEFORE_FLAG
                + "CHECK"
                + self.SEPARATOR
                + r".*?"
                + self.AFTER_FLAG,
                text.data,
            ):
                text.data = self.__check_boxes(text.data, check_tags)
            elif search(
                self.BEFORE_FLAG
                + r".*?"
                + self.SEPARATOR
                + r"\d+?"
                + self.AFTER_FLAG,
                text.data,
            ):
                text.data = self.__split_replace(text.data, simple_tags)
            else:
                text.data = self.__simple_replace(text.data, simple_tags)

        return doc

    def __clean_formatting(self, path: str) -> str:
        with ZipFile(path) as in_zip, ZipFile(
            ".clean".join(splitext(path)), "w"
        ) as out_zip:
            # For each files in the input zip file
            for in_zip_info in in_zip.infolist():
                with in_zip.open(in_zip_info) as in_file:
                    # Try reading it
                    try:
                        content = in_file.read()
                    except Exception:
                        raise Exception(
                            "Can't read properly the file "
                            f"{in_zip_info.filename} in {path}"
                        )

                    # Replace the text (to fill the document)
                    if in_zip_info.filename == "content.xml":

                        # Unifie the tags
                        content = sub(
                            r"{.*?{[^{}]+?}.*?}",
                            lambda m: sub("<[^<>]+?>", "", m.group(0)),
                            content.decode("utf-8"),
                        ).encode("utf-8")

                    # Try writing to the output file
                    try:
                        out_zip.writestr(in_zip_info.filename, content)
                    except Exception:
                        raise Exception(
                            f"Can't write the file {in_zip_info.filename} "
                            f"in {path}"
                        )
        return ".clean".join(splitext(path))

    def __split_tags(self, tags: dict):
        """
        Split the tags and values in 3 categories:
        - underTags: Tags used to underline text
        - checkTags: Tags used to check boxes
        - simpleTags: Tags that will simply be replaced with the values when
        seen
        """

        under_tags = {}
        check_tags = {}
        simple_tags = {}

        for key, val in tags.items():
            if search("UNDER" + self.SEPARATOR + ".*?", key):
                under_tags[key] = val
            elif search("CHECK" + self.SEPARATOR + ".*?", key):
                check_tags[key] = val
            else:
                simple_tags[key] = val

        return under_tags, check_tags, simple_tags

    def __simple_replace(self, text: str, values: dict) -> str:
        """
        Look for the flags and replace them with there corresponding values
        """
        for key, value in values.items():
            simple_key = self.BEFORE_FLAG + key + self.AFTER_FLAG

            if search(simple_key, text):
                if self.DEBUG:
                    print(f"KEY - Simple Replaced {simple_key} with {value}")
                text = text.replace(simple_key, str(value))

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

    def __underline(
        self,
        element: Text,
        parent_element: Element,
        values: dict,
        doc: OpenDocumentText,
    ):
        """
        For each underline key found in the values, we remove the key while
        keeping the text inside and if the underline value is true, we will
        add the underline subelement to the parent of the paragraph that
        contain the key
        """

        if "underline" not in doc._styles_dict.keys():
            # Create underline if it doesn't already exist
            u_style = Style(name="underline", family="text")
            u_style.addElement(
                TextProperties(
                    attributes={
                        "textunderlinestyle": "solid",
                        "textunderlinewidth": "auto",
                        "textunderlinecolor": "font-color",
                    }
                )
            )
            doc.automaticstyles.addElement(u_style)

        # Apply the style

        text = element.data
        text_list = sub(
            "(" + self.BEFORE_FLAG + ")|(" + self.AFTER_FLAG + ")",
            "|BLEG|",
            text,
        ).split("|BLEG|")

        step = 0 if text[:2] != "{{" else 1

        element.data = ""

        for index, sub_text in enumerate(text_list):
            p_text = sub(
                "^.+" + self.SEPARATOR + "(?=[^" + self.SEPARATOR + "]+$)",
                "",
                sub_text,
            )
            if (
                index % 2 == step
                and p_text in values.keys()
                and values[p_text]
            ):
                p_stylename = ""
                p_text = sub_text
            else:
                p_stylename = "underline"
            parent_element.addElement(
                odfText.Span(stylename=p_stylename, text=p_text)
            )

        return parent_element

    def __check_boxes(self, text: str, values: dict) -> str:
        """
        Look for a check tag and if there is, replace the tag with ✓ if the
        value is true and replaced the tag by □ in the other case.
        """
        for key, value in values.items():
            if key[:6] != "CHECK_":
                continue

            simple_key = self.BEFORE_FLAG + key + self.AFTER_FLAG
            new_value = "✓" if value else "□"

            if self.DEBUG:
                print(f"KEY - Checked replaced {simple_key} with {new_value}")

            text = text.replace(simple_key, new_value)
        return text

    def fill_document(
        self, src_path: str, values: dict, dest_path: str
    ) -> None:
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

        src_path = self.__clean_formatting(src_path)

        # Open the odt file
        doc = load(src_path)

        if self.DEBUG:
            print(f"INFO - Read {src_path}")

        doc = self.__replace_text(doc, values)

        doc.save(dest_path)

        if self.DEBUG:
            print(f"INFO - Saved {dest_path}")

        run(["rm", src_path, "-f"])

        if self.PDF or self.PDF_ONLY:
            run(
                [
                    "lowriter",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    split(dest_path)[0],
                    dest_path,
                ]
            )
        if self.PDF_ONLY:
            run(["rm", dest_path])
