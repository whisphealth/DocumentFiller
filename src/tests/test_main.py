from DocumentFiller.__main__ import parse_path


def test_parse_path():
    test1 = parse_path(["test/path.ext", "test2/path2.ext"], None)
    result1 = zip(
        ["test/path.ext", "test2/path2.ext"],
        ["test/path.out.ext", "test2/path2.out.ext"],
    )

    assert list(test1) == list(result1)

    test2 = parse_path(
        ["test/path.ext", "test2/path2.ext", "test3/path3.ext"], "src/tests"
    )
    result2 = zip(
        ["test/path.ext", "test2/path2.ext", "test3/path3.ext"],
        ["src/tests/path.ext", "src/tests/path2.ext", "src/tests/path3.ext"],
    )

    assert list(test2) == list(result2)

    test3 = parse_path(["test/path.ext"], "README.md")
    result3 = zip(["test/path.ext"], ["README.md"])

    assert list(test3) == list(result3)
