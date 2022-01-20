from DocumentFiller.DocFillerFactory import DocumentFillerFactory
from DocumentFiller.OdtFiller import OdtFiller

import zipfile
import json
import subprocess

doc_path = "./src/tests/documents/"


def test_OdtFiller():
    src_path = doc_path + "Test1.odt"
    json_path = doc_path + "Flags1.json"
    out_path = doc_path + "Test1.out.odt"
    goal_path = doc_path + "Test1.goal.odt"

    dff = DocumentFillerFactory()
    dff_odt = dff.inst("odt")

    dff_odt.fill(
        src_path,
        json.load(open(json_path)),
        out_path,
    )

    with zipfile.ZipFile(out_path) as out, zipfile.ZipFile(goal_path) as goal:
        for out_zip_info in out.infolist():
            if out_zip_info.filename == "content.xml":
                out_content = out_zip_info
                break

        for goal_zip_info in goal.infolist():
            if goal_zip_info.filename == "content.xml":
                goal_content = goal_zip_info
                break

        with out.open(out_content) as f_out, goal.open(goal_content) as f_goal:
            assert f_out.read() == f_goal.read()

    subprocess.run(["rm", out_path])


def test_clean_formatting():
    dff_odf = OdtFiller()

    path_clean = dff_odf.clean_formatting(doc_path + "Test1.odt")

    with zipfile.ZipFile(path_clean) as out, zipfile.ZipFile(
        doc_path + "Test1.clean.goal.odt"
    ) as goal:
        for out_zip_info in out.infolist():
            if out_zip_info.filename == "content.xml":
                out_content = out_zip_info
                break

        for goal_zip_info in goal.infolist():
            if goal_zip_info.filename == "content.xml":
                goal_content = goal_zip_info
                break

        with out.open(out_content) as f_out, goal.open(goal_content) as f_goal:
            assert f_out.read() == f_goal.read()

    subprocess.run(["rm", path_clean])


def test_split_tags():
    dff_odt = OdtFiller()

    flags = json.load(open(doc_path + "Flags1.json", "r"))

    under_tags, check_tags, simple_tags, if_tags = dff_odt.split_tags(flags)

    goal_under_tags = {"UNDER_A": True, "UNDER_B": True, "UNDER_C": False}

    goal_check_tags = {"CHECK_A": True, "CHECK_B": False}

    goal_simple_tags = {
        "NAME": "Mathieu",
        "ADDRESS": "Rue du fromage baguette",
    }

    assert under_tags == goal_under_tags
    assert check_tags == goal_check_tags
    assert simple_tags == goal_simple_tags


def test_simple_replace():
    dff_odt = OdtFiller()
    flags = json.load(open(doc_path + "Flags1.json", "r"))

    text = "Bonjour {{NAME}}, address: {{ADDRESS}}"
    goal_text = "Bonjour Mathieu, address: Rue du fromage baguette"

    replaced_text = dff_odt.simple_replace(text, flags)

    assert replaced_text == goal_text


def test_split_replace():
    dff_odt = OdtFiller()
    flags = json.load(open(doc_path + "Flags1.json", "r"))

    text = "Bonjour {{NAME_00}}, {{NAME_01}}, {{NAME_02}}"
    goal_text = "Bonjour M, a, t"

    replaced_text = dff_odt.split_replace(text, flags)

    assert replaced_text == goal_text


def test_check_boxes():
    dff_odt = OdtFiller()
    flags = json.load(open(doc_path + "Flags1.json", "r"))

    text = "List:\n{{CHECK_A}} Buy some food\n{{CHECK_B}} Pet the cat"
    goal_text = "List:\n✓ Buy some food\n□ Pet the cat"

    replaced_text = dff_odt.check_boxes(text, flags)

    assert replaced_text == goal_text


def test_if_replace():
    dff_odt = OdtFiller()
    flags_1 = {"IF_MANGER": True, "IF_FAIM": False}
    flags_2 = {"IF_MANGER": False, "IF_FAIM": True}

    text_1 = "J'ai {{IF_MANGER_THEN_pas faim_ELSE_très faim_END}} aujourd'hui"
    text_2 = (
        "J'ai {{IF_FAIM_THEN_plutot faim_ELSE_mal au ventre_END}} aujourd'hui"
    )
    goal_text_1_1 = "J'ai pas faim aujourd'hui"
    goal_text_1_2 = "J'ai mal au ventre aujourd'hui"
    goal_text_2_1 = "J'ai très faim aujourd'hui"
    goal_text_2_2 = "J'ai plutot faim aujourd'hui"

    replaced_text_1_1 = dff_odt.if_replace(text_1, flags_1)
    replaced_text_1_2 = dff_odt.if_replace(text_2, flags_1)
    replaced_text_2_1 = dff_odt.if_replace(text_1, flags_2)
    replaced_text_2_2 = dff_odt.if_replace(text_2, flags_2)

    assert replaced_text_1_1 == goal_text_1_1
    assert replaced_text_1_2 == goal_text_1_2
    assert replaced_text_2_1 == goal_text_2_1
    assert replaced_text_2_2 == goal_text_2_2
