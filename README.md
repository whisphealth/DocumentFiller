# Document Filler

## What is it ?

Document filler is a python library that allow someone to replace tags in a word document with personalized content

## What does DocumentFiller have that other libraries don't ?

- Doesn't require Microsoft Word
- Doesn't mess up the formatting of documents
- Can replace tags character by character

## Installation

Download the release file DocumentFiller-... .whl

`pip install ./DocumentFiller-version-py3-none-any.whl`

## How to use it

### CLI

Usage: `DocumentFiller [-d|--debug] [--pdf] [--pdfonly] -j|-json|-jsonpath jsonpath [-o|-output outputpath] inputfile(s) `

Example:

`DocumentFiller -j "{\"key\": \"value\"}" -o out.docx in.docx`

This command will replace all the `{{KEY}}` present in the `in.docx` with `value` and save the result in `out.docx`

`DocumentFiller -j values.json document.docx`

This command will replace all the keys defined in `values.json` present in `document.docx` and save the result in `document.out.docx`

### In a python file

```python
from DocumentFiller.DocFillerFactory import DocumentFillerFactory

dff = DocumentFillerFactory()

docxFiller = dff.inst("docx")
docxFiller.fill("sourcePath", {"key": "value"}, "destinationPath")

xlsxFiller = dff.inst("xlsx")
xlsxFiller.fill("sourcePath", {"key": "value"}, "destinationPath")
```

### In the files

#### Simple flag replacement

Use this style of flags: `{{FLAG}}`

All the character of a flag must keep the same formatting.

#### Split flag replacement

Use this style of flags: `{{FLAG_00}}`, `{{FLAG_01}}`, ...

The flags will be replaced by a single character.

#### Check boxes (docx only)

Use this style of flags: `{{CHECK_FLAG}}`

If the value `CHECK_FLAG` in the json is true, the flag will be replaced with `✓`. In the other case, it will be replaced with `□`.

#### Underline a text (docx only)

Use this style of flags: `{{UNDER_FLAG_textThatMightGetUnderline}}`

If the value `UNDER_FLAG` in the json is true, the text associated will be underline. In the other case, it will stay untouched.

## Supported files

- Microsoft word files xml format (.docx)
- Microsoft excel files xml format (.xlsx)

## Know issues

- PDF conversion is done by Libre Office so some formatting issues might apply
