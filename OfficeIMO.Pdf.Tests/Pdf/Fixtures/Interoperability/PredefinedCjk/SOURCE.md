# Predefined CJK font fixtures

`generate.py` creates synthetic one-page PDFs using ReportLab 4.4.9 and pypdf
6.10.0. The four positive files contain Japanese, Simplified Chinese, Traditional
Chinese or Korean text through horizontal UCS-2 predefined encodings. They contain
no embedded font program and no ToUnicode map. Source versions, exact hashes,
selected text and independently calculated line advances are in `source.json`.

ReportLab 4.4.9 declares `UniGB-UCS2-H` for `MSung-Light`, whose character
collection is Adobe/CNS1. The generator retains that mismatch as
`mismatched-collection.pdf`, then corrects the positive Traditional Chinese file
to `UniCNS-UCS2-H` using pypdf. OfficeIMO refuses the mismatched source rather
than interpreting it with a different character collection.

The documents and generator are MIT licensed. ReportLab and pypdf are test-only
independent producers under their BSD licenses. They are not product runtime or
ordinary test-run prerequisites. Regenerate explicitly with:

```sh
python generate.py
```

The regression cases cover extraction, CID-based advances, precise text-only
redaction, neighboring text, verified removal and source immutability. Run the
[independent reader checks](../../../../../Build/PdfViewerVerification/README.md)
on exported source/result pairs to qualify the selected viewer behavior.
