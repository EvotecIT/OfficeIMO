# OpenDocument interoperability fixtures

`libreoffice-writer-basic.odt` was produced with LibreOfficeDev 26.8 by converting `OfficeIMO.TestAssets/Documents/DocumentWithTables.docx` with the `writer8` filter on 2026-07-10. The preservation test changes one text paragraph and verifies that every other package entry retains its original bytes.

`microsoft-word-basic.odt` was saved from the same source document with Microsoft Word for Mac 16.99 on 2026-07-10 using Word's OpenDocument Text export.

Keep fixtures small and record the producing application, source file, and date here. Binary fixtures are test evidence; they are not public examples.
