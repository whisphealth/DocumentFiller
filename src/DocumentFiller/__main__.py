from argparse import ArgumentParser
from genericpath import isdir, isfile
from json import load, loads
from re import sub
from .DocFillerFactory import DocumentFillerFactory
from os.path import splitext, split


def parse_path(path_in: list, path_out: str) -> zip:
    """
    Take the list of input files path and infos on the output
    file and parse them to two list of input and output files
    :param path_in: (List) List of path to the input files
    :param path_out: (String) Path to the output file or folder
    """

    if not path_out:
        r_path_out = [sub(r"\.(?=[^\.\/]+?$)", ".out.", f) for f in path_in]
    elif isdir(path_out):
        path_out = path_out + ("/" if (path_out[-1:] != "/") else "")

        r_path_out = [path_out + split(f)[1] for f in path_in]
    elif isfile(path_out) and len(path_in) == 1:
        r_path_out = [path_out]
    else:
        raise Exception("srcpath and / or destpath is invalid")

    return zip(path_in, r_path_out)


def main():
    parser = ArgumentParser(
        description="Fill documents by replacing flags in docx files and put"
        " the result in an other directory",
        usage="DocumentFiller [-d|--debug] [--pdf] [--pdfonly] "
        "-j|-json|-jsonpath jsonpath [-o|-output outputpath] inputfile(s)",
    )

    parser.add_argument(
        "input",
        type=str,
        nargs="*",
        help="Source path to a docx file(s)",
    )
    parser.add_argument(
        "-j",
        "--json",
        "--jsonpath",
        type=str,
        help="Json file path containing the keys",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=str,
        help="Destination path to a folder or a docx file",
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
    if not all([isfile(f) for f in args.input]):
        raise Exception("One or more input file does not exist")

    if not isfile(args.json):
        try:
            json = loads(args.json)
        except Exception:
            raise Exception(
                "Second argument is not a file and is not parsable"
            )
    else:
        try:
            json = load(open(args.json))
        except Exception:
            raise Exception("Json file is not parsable")

    # DocFiller

    dff = DocumentFillerFactory(args.debug, args.pdf, args.pdfonly)
    # Calling the function
    for file_in, file_out in parse_path(args.input, args.output):
        doc_filler = dff.inst(splitext(file_in)[1])
        if doc_filler:
            doc_filler.fill(file_in, json, file_out)


if __name__ == "__main__":
    main()
