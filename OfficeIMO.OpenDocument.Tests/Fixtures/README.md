# OpenDocument interoperability fixtures

`libreoffice-writer-basic.odt` was produced with LibreOfficeDev 26.8 by converting `OfficeIMO.TestAssets/Documents/DocumentWithTables.docx` with the `writer8` filter on 2026-07-10. The preservation test changes one text paragraph and verifies that every other package entry retains its original bytes.

`microsoft-word-basic.odt` was saved from the same source document with Microsoft Word for Mac 16.99 on 2026-07-10 using Word's OpenDocument Text export.

`libreoffice-calc-basic.ods` was produced with LibreOfficeDev 26.8 by converting `OfficeIMO.TestAssets/Documents/BasicExcel.xlsx` with the `calc8` filter on 2026-07-10.

`microsoft-excel-basic.ods` was saved from the same source workbook with Microsoft Excel for Mac 16.99 on 2026-07-10 using Excel's OpenDocument Spreadsheet export.

`microsoft-excel-column-chart.ods`, `microsoft-excel-bar-chart.ods`, and `microsoft-excel-line-chart.ods` were saved with Microsoft Excel for Windows 16.0 on 2026-09-24. `Generate-MicrosoftExcelCharts.ps1` creates the small source workbook and exports all three variants. They cover the embedded chart object, source-cell ranges, series values, and package preservation.

`extreme-repeats.ods` was generated with OfficeIMO.OpenDocument on 2026-07-10. Its last used cell is XFD1048576, represented by row and cell repeat runs instead of expanded XML nodes.

`libreoffice-impress-basic.odp` was produced with LibreOfficeDev 26.8 by converting `Assets/PowerPointTemplates/PowerPointWithTitle.pptx` with the `impress8` filter on 2026-07-10.

`microsoft-powerpoint-basic.odp` was saved from the same source presentation with Microsoft PowerPoint for Mac 16.99 on 2026-07-10 using PowerPoint's OpenDocument Presentation export.

Keep fixtures small and record the producing application, source file, and date here. Binary fixtures are test evidence; they are not public examples.

`producer-manifest.json` is the machine-readable source of truth for the authored producer fixtures, their exact hashes, and the tests that own style, formula, drawing, embedded-content, and unknown-package-content contracts. It also records a public Google Docs ODT export as external-only evidence; the source bytes are intentionally not redistributed. `Build/Test-OpenDocumentExternalProducerEvidence.ps1` verifies its bounded package size and stable semantic hash, then loads, validates, preserves, saves, reopens, and validates the artifact through `OfficeIMO.OpenDocument` with strict loss reporting.
