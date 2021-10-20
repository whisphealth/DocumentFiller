from re import sub, findall, search, DOTALL, finditer
from zipfile import ZipFile
from os.path import exists
from lxml import etree
from subprocess import run
from .DocFillerFamily import DocumentFillerFamilly


class DocxFiller(DocumentFillerFamilly):

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
        """
        Function that replace le flags in text with the values associated to
        the flag in values
        :param text: (String) Xml input text that contain the flags to replace
        :param values: (Dictionary) Dictionary containing the flags (keys) and
        the values (items)
        :return: (String) Xml Text with all the flags replaced
        """

        text = self.__clean_formatting(text.decode("utf-8")).encode("utf-8")

        try:
            root = etree.fromstring(text)
        except Exception:
            raise Exception("The cleaning process broke the xml")

        under_tags, check_tags, simple_tags = self.__split_tags(tags)

        child_to_parent = {c: p for p in root.iter() for c in p}

        for i in root.iter():
            # If there is no flags, no need to check for them
            if not search(
                self.BEFORE_FLAG + ".*?" + self.AFTER_FLAG, str(i.text)
            ):
                continue

            if search(
                self.BEFORE_FLAG
                + "UNDER"
                + self.SEPARATOR
                + ".*?"
                + self.AFTER_FLAG,
                i.text,
            ):
                p = child_to_parent[i]
                p = self.__underline(p, under_tags, root.nsmap)
            elif search(
                self.BEFORE_FLAG
                + "CHECK"
                + self.SEPARATOR
                + r".*?"
                + self.AFTER_FLAG,
                i.text,
            ):
                i.text = self.__check_boxes(i.text, check_tags)
            elif search(
                self.BEFORE_FLAG
                + r".*?"
                + self.SEPARATOR
                + r"\d+?"
                + self.AFTER_FLAG,
                i.text,
            ):
                i.text = self.__split_replace(i.text, simple_tags)
            else:
                i.text = self.__simple_replace(i.text, simple_tags)

        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'.encode(
                "utf-8"
            )
            + etree.tostring(root, encoding="utf-8")
        )

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

    def __underline(self, tree: etree, values: dict, ns) -> etree:
        """
        For each underline key found in the values, we remove the key while
        keeping the text inside and if the underline value is true, we will
        add the underline subelement to the parent of the paragraph that
        contain the key
        """

        for key, value in values.items():
            under_key = (
                self.BEFORE_FLAG
                + key
                + self.SEPARATOR
                + ".*?"
                + self.AFTER_FLAG
            )

            for i in tree.iter():
                if isinstance(i.text, str) and search(under_key, i.text):

                    if self.DEBUG:
                        print(
                            f"KEY - Cleaned and underlined {i.text} ", end=""
                        )

                    i.text = sub(
                        self.BEFORE_FLAG + key + self.SEPARATOR,
                        "",
                        sub(self.AFTER_FLAG, "", i.text),
                    )

                    if self.DEBUG:
                        print(f"with {i.text}")

                    if value:

                        # Underlining trick
                        pr = tree.find("{*}rPr")
                        etree.SubElement(
                            pr,
                            "{" + ns["w"] + "}u",
                            {"{" + ns["w"] + "}val": "single"},
                        )

        return tree

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

    def __clean_formatting(self, text: str) -> str:
        """
        Clean the formatting of the string to make it properly readable and
        editable by :
        - cleaning the xml tags inside the flags so they can be read
        - Isolating the flags in a paragraph so that if any formatting should
          be applied, it won't affect the text next to the tag
        """

        text = self.__remove_tags(text)

        text = self.__divide_paragraphs_btw_flags(text)

        text = self.__divide_paragraphs(text)

        return text

    def __remove_tags(self, text: str) -> str:
        """
        Cleaning the xml tags inside the flags so they can be read
        """
        text = sub(
            r"(?<={)(<[^>]*>)+(?=[\{%])|(?<=[%\}])(<[^>]*>)+(?=\})",
            "",
            text,
            flags=DOTALL,
        )

        def striptags(m):
            return sub(
                r"</w:t>.*?(<w:t>|<w:t [^>]*>)", "", m.group(0), flags=DOTALL
            )

        text = sub(r"{{(?:(?!}}).)*", striptags, text, flags=DOTALL)

        try:
            etree.fromstring(text.encode("utf-8"))
        except Exception:
            raise Exception(
                "The cleaning process broke the xml by cleaning the "
                "formatting inside the flags.\nMaybe one or more flags in "
                "the document don't have a single formatting per flag (part "
                "of a flag underlined / different size)"
            )

        return text

    def __divide_paragraphs_btw_flags(self, text: str) -> str:
        """
        When two flags are on the same paragraph, this function will separate
        them on different paragraphs so that one's formatting doesn't affect
        the other
        """
        # Finding the start and end position of each paragraph
        pos_start_par = [m.start() for m in finditer("<w:r>", text)]
        pos_stop_par = [m.end() for m in finditer("</w:r>", text)]

        # Keeping the text in an another var so that the modifications to the
        # original don't mess up the start and end position of each paragraph
        old_text = text

        for tag in finditer(r"{{\w+?}}[^<>]+?{{\w+?}}", text):
            # Find the start and end position of the paragraph it is in
            start_par = max([i for i in pos_start_par if i < tag.start()])
            stop_par = min([i for i in pos_stop_par if i > tag.end()])

            # Get the whole paragraph, the beggining and the end
            tag_text = old_text[start_par:stop_par]
            begin_tag = search(r"^.+?\>(?!\<)", tag_text).group(0)
            after_tag = search(r"[^\>]\<.+?$", tag_text).group(0)
            new_text = ""

            # Separate each flags with the text that precede it
            for flag in finditer(r".+?(}}|(?=<))", tag_text):
                flag_txt = sub(r"<.*?>", "", str(flag.group(0)))
                if flag_txt not in ["", " " * len(flag_txt)]:
                    new_text += begin_tag + flag_txt + after_tag

            text = text.replace(tag_text, new_text)

            if self.DEBUG:
                print(f"FORMAT - Separate\n{tag_text}\nwith\n{new_text}")

        # Check if the function ruined the document
        try:
            etree.fromstring(text.encode("utf-8"))
        except Exception:
            raise Exception(
                "The cleaning process broke the xml by replacing\n"
                f"{old_text}\nwith\n{text}"
            )

        return text

    def __divide_paragraphs(self, text: str) -> str:
        """
        Isolating the flags in a paragraph so that if any formatting should be
        applied, it won't affect the text next to the tag
        """
        # Finding the start and end position of each paragraph
        pos_start_par = [m.start() for m in finditer("<w:r( [^>]+?)?>", text)]
        pos_stop_par = [m.end() for m in finditer("</w:r>", text)]

        # Keeping the text in an another var so that the modifications to the
        # original don't mess up the start and end position of each paragraph
        old_text = text

        # Foreach tag that we find
        for tag in finditer(r"{{UNDER_.+?}}", text):
            # Find the start and end position of the paragraph it is in
            start_par = max([i for i in pos_start_par if i < tag.start()])
            stop_par = min([i for i in pos_stop_par if i > tag.end()])

            # Get the whole paragraph
            tag_text = old_text[start_par:stop_par]

            # The first tag contain the text before the flag so we copy the
            # paragraph removing just the flag and the text after
            tag_a = sub(r"{{.+?}}.*?(?=<\/w:t>)", "", tag_text)
            # The second tag contain only the tag in the paragraph
            tag_b = sub(
                r"([^\>]*?(?={{))|((?<=}}).*?(?=<\/w:t>))", "", tag_text
            )
            # The third tag contain the text after the flag
            tag_c = sub(r"(?<=\>)[^\>]+?}}", "", tag_text)

            # If the content of the tags are empty (so no text before or after
            # the flag) then remove them
            if len(findall(r">(?![<{])", tag_a)) <= 1:
                tag_a = ""
            if len(findall(r">(?![<{])", tag_c)) <= 1:
                tag_c = ""

            try:
                prev = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes'
                    '"?><w:document xmlns:o="urn:schemas-microsoft-com:offic'
                    'e:office" xmlns:r="http://schemas.openxmlformats.org/of'
                    'ficeDocument/2006/relationships" xmlns:v="urn:schemas-m'
                    'icrosoft-com:vml" xmlns:w="http://schemas.openxmlformat'
                    's.org/wordprocessingml/2006/main" xmlns:w10="urn:schema'
                    's-microsoft-com:office:word" xmlns:wp="http://schemas.o'
                    'penxmlformats.org/drawingml/2006/wordprocessingDrawing"'
                    ' xmlns:wps="http://schemas.microsoft.com/office/word/20'
                    '10/wordprocessingShape" xmlns:wpg="http://schemas.micro'
                    'soft.com/office/word/2010/wordprocessingGroup" xmlns:mc'
                    '="http://schemas.openxmlformats.org/markup-compatibilit'
                    'y/2006" xmlns:wp14="http://schemas.microsoft.com/office'
                    '/word/2010/wordprocessingDrawing" xmlns:w14="http://sch'
                    'emas.microsoft.com/office/word/2010/wordml" mc:Ignorabl'
                    'e="w14 wp14"> <w:body>'
                )

                after = "</w:body></w:document>"

                etree.fromstring(
                    (prev + tag_a + tag_b + tag_c + after).encode("utf-8")
                )
            except Exception:
                raise Exception(
                    "The cleaning process broke the xml by replacing\n"
                    f"{tag_text}\nwith\n{tag_a}\n{tag_b}\n{tag_c}"
                )

            if self.DEBUG:
                print(
                    f"FORMAT - Replaced\n{tag_text}\nwith\n{tag_a}\n{tag_b}"
                    f"\n{tag_c}"
                )
            text = text.replace(tag_text, tag_a + tag_b + tag_c)

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

        # Open the two zip files
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
                    if in_zip_info.filename == "word/document.xml":

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

        if self.PDF or self.PDF_ONLY:
            run(
                [
                    "lowriter",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    sub(r"[^\/]+?\.docx", "", dest_path),
                    dest_path,
                ]
            )
        if self.PDF_ONLY:
            run(["rm", dest_path])
