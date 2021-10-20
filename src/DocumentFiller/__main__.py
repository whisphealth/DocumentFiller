from argparse import ArgumentParser
from genericpath import isdir, isfile
from json import load, loads
from posix import listdir
from re import search
from .DocFillerFactory import DocumentFillerFactory


def parse_path(path_in: str, path_out: str) -> zip:
    """
    Take the two path to a folder or a file and parse them to a list of
    paths to files
    :param path_in: (String) Path to the input file or folder
    :param path_out: (String) Path to the output file or folder
    """
    path_in = path_in + (
        "/" if (path_in[-1:] != "/" and isdir(path_in)) else ""
    )
    path_out = path_out + (
        "/" if (path_out[-1:] != "/" and isdir(path_out)) else ""
    )

    if isfile(path_in):
        r_path_in = [path_in]
    elif isdir(path_in):
        r_path_in = [
            path_in + f
            for f in listdir(path_in)
            if (isfile(path_in + f) and f[:2] != "~$")
        ]
    else:
        raise Exception("srcpath is not a file / directory")

    if isfile(path_in) and not isdir(path_out):
        r_path_out = [path_out]
    elif isfile(path_in) and isdir(path_out):
        r_path_out = [path_out + search(r"[^\/]+?\..+?$", path_in).group(0)]
    elif isdir(path_in) and isdir(path_out):
        r_path_out = [f.replace(path_in, path_out) for f in r_path_in]
    else:
        raise Exception("srcpath and / or destpath is invalid")

    return zip(r_path_in, r_path_out)


def main():
    parser = ArgumentParser(
        description="Fill documents by replacing flags in docx files and put"
        " the result in an other directory",
        usage="DocumentFiller [-d|--debug] [--pdf] [--pdfonly] scrpath "
        "jsonpath destpath",
    )

    parser.add_argument(
        "srcpath", type=str, help="Source path to a docx file or folder"
    )
    parser.add_argument(
        "jsonpath", type=str, help="Json file path containing the keys"
    )
    parser.add_argument(
        "destpath", type=str, help="Destination path to a docx file or folder"
    )
    parser.add_argument(
        "-d", "--debug", action="store_true", help="Show debug infos"
    )
    parser.add_argument(
        "--pdf",
        action="store_true",
        help="Generate a pdf for every docx generated",
    )
    parser.add_argument(
        "--pdfonly",
        action="store_true",
        help="Generate only pdf and no docx documents",
    )

    args = parser.parse_args()

    # Arguments errors detection
    if (not isfile(args.srcpath)) and (not isdir(args.srcpath)):
        raise Exception("First argument file/folder does not exist")
    if not isfile(args.jsonpath):
        try:
            json = loads(args.jsonpath)
        except Exception:
            raise Exception(
                "Second argument is not a file and is not parsable"
            )
    else:
        try:
            json = load(open(args.jsonpath))
        except Exception:
            raise Exception("Json file is not parsable")

    # File path correction
    file_dest_path = args.destpath + (
        "/" if (args.destpath[-1:] != "/" and isdir(args.destpath)) else ""
    )
    file_src_path = args.srcpath + (
        "/" if (args.srcpath[-1:] != "/" and isdir(args.srcpath)) else ""
    )

    dff = DocumentFillerFactory(args.debug, args.pdf, args.pdfonly)
    # Calling the function
    for file_in, file_out in parse_path(file_src_path, file_dest_path):
        doc_filler = dff.inst(search(r"[^\.]+?$", file_in).group(0))
        if doc_filler:
            doc_filler.fill(file_in, json, file_out)


if __name__ == "__main__":
    main()
