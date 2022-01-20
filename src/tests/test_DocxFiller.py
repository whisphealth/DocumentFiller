from DocumentFiller.DocFillerFactory import DocumentFillerFactory
from DocumentFiller.DocxFiller import DocxFiller
import zipfile
import json
import subprocess

doc_path = "./src/tests/documents/"


def test_DocxFiller():
    src_path = doc_path + "Test1.docx"
    json_path = doc_path + "Flags1.json"
    out_path = doc_path + "Test1.out.docx"
    goal_path = doc_path + "Test1.goal.docx"

    dff = DocumentFillerFactory()
    dff_docx = dff.inst("docx")

    dff_docx.fill(
        src_path,
        json.load(open(json_path, "r")),
        out_path,
    )

    with zipfile.ZipFile(out_path) as out, zipfile.ZipFile(goal_path) as goal:
        for out_zip_info in out.infolist():
            if out_zip_info.filename == "word/document.xml":
                out_content = out_zip_info
                break

        for goal_zip_info in goal.infolist():
            if goal_zip_info.filename == "word/document.xml":
                goal_content = goal_zip_info
                break

        with out.open(out_content) as f_out, goal.open(goal_content) as f_goal:
            assert f_out.read() == f_goal.read()

    subprocess.run(["rm", out_path])


def test_replace_text():
    dff_docx = DocxFiller()

    with open(doc_path + "docx 1 document.xml", "rb") as file:
        shared_strings = file.read()
    with open(doc_path + "docx 1 Goal document.xml", "rb") as file:
        goal_shared_strings = file.read()
    flags = json.load(open(doc_path + "Flags1.json", "r"))

    treated_shared_strings = dff_docx.replace_text(shared_strings, flags)

    assert treated_shared_strings == goal_shared_strings


def test_split_tags():
    dff_docx = DocxFiller()

    flags = json.load(open(doc_path + "Flags1.json", "r"))

    under_tags, check_tags, simple_tags = dff_docx.split_tags(flags)

    goal_under_tags = {"UNDER_A": True, "UNDER_B": True, "UNDER_C": False}

    goal_check_tags = {"CHECK_A": True, "CHECK_B": False}

    goal_simple_tags = {
        "NAME": "Mathieu",
        "ADDRESS": "Rue du fromage baguette",
        "IF_MANGER": True,
        "IF_FAIM": False,
    }

    assert under_tags == goal_under_tags
    assert check_tags == goal_check_tags
    assert simple_tags == goal_simple_tags


def test_simple_replace():
    dff_docx = DocxFiller()
    flags = json.load(open(doc_path + "Flags1.json", "r"))

    text = "Bonjour {{NAME}}, address: {{ADDRESS}}"
    goal_text = "Bonjour Mathieu, address: Rue du fromage baguette"

    replaced_text = dff_docx.simple_replace(text, flags)

    assert replaced_text == goal_text


def test_split_replace():
    dff_docx = DocxFiller()
    flags = json.load(open(doc_path + "Flags1.json", "r"))

    text = "Bonjour {{NAME_00}}, {{NAME_01}}, {{NAME_02}}"
    goal_text = "Bonjour M, a, t"

    replaced_text = dff_docx.split_replace(text, flags)

    assert replaced_text == goal_text


def test_check_boxes():
    dff_docx = DocxFiller()
    flags = json.load(open(doc_path + "Flags1.json", "r"))

    text = "List:\n{{CHECK_A}} Buy some food\n{{CHECK_B}} Pet the cat"
    goal_text = "List:\n✓ Buy some food\n□ Pet the cat"

    replaced_text = dff_docx.check_boxes(text, flags)

    assert replaced_text == goal_text


def test_remove_tags():
    dff_docx = DocxFiller()

    with open(doc_path + "docx 1 remove tags document.xml", "r") as file:
        text = file.read()
    with open(doc_path + "docx 1 goal remove tags document.xml", "r") as file:
        goal_text = file.read()

    replaced_text = dff_docx.remove_tags(text)

    assert replaced_text == goal_text
