from DocumentFiller.DocFillerFactory import DocumentFillerFactory
from DocumentFiller.XlsxFiller import XlsxFiller
import zipfile
import json
import subprocess

doc_path = "./src/tests/documents/"


def test_XlsxFiller():
    src_path = doc_path + "Test1.xlsx"
    json_path = doc_path + "Flags1.json"
    out_path = doc_path + "Test1.out.xlsx"
    goal_path = doc_path + "Test1.goal.xlsx"

    dff = DocumentFillerFactory()
    dff_xlsx = dff.inst("xlsx")

    dff_xlsx.fill(
        src_path,
        json.load(open(json_path)),
        out_path,
    )

    with zipfile.ZipFile(out_path) as out, zipfile.ZipFile(goal_path) as goal:
        for out_zip_info in out.infolist():
            if out_zip_info.filename == "xl/sharedStrings.xml":
                out_content = out_zip_info
                break

        for goal_zip_info in goal.infolist():
            if goal_zip_info.filename == "xl/sharedStrings.xml":
                goal_content = goal_zip_info
                break

        with out.open(out_content) as f_out, goal.open(goal_content) as f_goal:
            assert f_out.read() == f_goal.read()

    subprocess.run(["rm", out_path])


def test_replace_text():
    dff_xlsx = XlsxFiller()

    with open(doc_path + "xlsx 1 sharedStrings.xml", "rb") as file:
        shared_strings = file.read()
    with open(doc_path + "xlsx 1 Goal sharedStrings.xml", "rb") as file:
        goal_shared_strings = file.read()
    flags = json.load(open(doc_path + "Flags1.json"))

    treated_shared_strings = dff_xlsx.replace_text(shared_strings, flags)

    assert treated_shared_strings == goal_shared_strings


def test_simple_replace():
    dff_xlsx = XlsxFiller()
    flags = json.load(open(doc_path + "Flags1.json"))

    text = "Bonjour {{NAME}}, address: {{ADDRESS}}"
    goal_text = "Bonjour Mathieu, address: Rue du fromage baguette"

    replaced_text = dff_xlsx.simple_replace(text, flags)

    assert replaced_text == goal_text


def test_split_replace():
    dff_xlsx = XlsxFiller()
    flags = json.load(open(doc_path + "Flags1.json"))

    text = "Bonjour {{NAME_00}}, {{NAME_01}}, {{NAME_02}}"
    goal_text = "Bonjour M, a, t"

    replaced_text = dff_xlsx.split_replace(text, flags)

    assert replaced_text == goal_text
