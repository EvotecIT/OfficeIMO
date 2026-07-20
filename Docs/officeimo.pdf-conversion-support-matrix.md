# OfficeIMO PDF Conversion Support Matrix

This matrix is generated from `Docs/pdf-conversion-scenarios.json`. Fidelity status describes the current evidence, not the intended destination.

Premium claim rule: A converter can be marked externally-verified only when its declared reference policy has stable source artifacts, producer/version metadata, page geometry, visual comparison results, text/link/structure proof, and no unexpected conversion warnings.

| Source | Formats | Mode | Evidence status | Reference policy |
| --- | --- | --- | --- | --- |
| word | doc, docx, docm, dotx, dotm | native-paged | candidate | microsoft-office-plus-officeimo-regression |
| excel | xlsx, xlsm, xltx, xltm, xls, xlsb | native-paged | candidate | microsoft-excel-plus-officeimo-regression |
| powerpoint | pptx, pptm | native-slide-canvas | candidate | microsoft-powerpoint-plus-officeimo-regression |
| markdown | md, markdown | semantic-document | regression-proven | officeimo-regression |
| html | html, htm, mhtml, mht | native-static-paged | candidate | standards-corpus-plus-officeimo-regression |
| rtf | rtf | semantic-document | regression-proven | officeimo-regression |
| onenote | one, onetoc2, onepkg | explicit-semantic-document | accepted-degradation | semantic-contract |
| asciidoc | adoc, asciidoc, asc | loss-aware-semantic-document | accepted-degradation | semantic-contract |
| latex | tex, latex | loss-aware-semantic-document | accepted-degradation | semantic-contract |

## Capability Claims

| Source | Capability | Fidelity level | Evidence scenarios |
| --- | --- | --- | --- |
| word | semantic-text-headings-lists-and-tables | exact | word-native-report |
| word | advanced-word-layout-and-drawing-objects | supported-with-approximation | word-native-report |
| excel | worksheet-values-formats-links-and-basic-pagination | exact | excel-native-daily-workbook |
| excel | worksheet-canvas-drawings-charts-and-dashboard-layout | supported-with-approximation | excel-native-daily-workbook, excel-dashboard-report |
| powerpoint | slide-geometry-basic-text-images-and-tables | exact | powerpoint-native-dense-layout |
| powerpoint | advanced-theme-drawingml-smartart-and-chart-layout | supported-with-approximation | powerpoint-layout-theme-groups |
| html | static-business-html-shared-png-svg-and-searchable-pdf-scene | exact | html-static-market-corpus |
| html | advanced-css-fragmentation-typography-and-svg-effects | supported-with-approximation | html-static-market-corpus, html-css-resource-policy |
| html | javascript-and-interactive-browser-state | unsupported | html-css-resource-policy |

## Direct, Composed, And Planned Routes

| Route | Formats | Status | Implementation owner | Contract evidence | Diagnostic contract |
| --- | --- | --- | --- | --- | --- |
| opendocument-text-via-word | odt, ott | direct-loss-aware-adapter | `OfficeIMO.OpenDocument.Pdf/OfficeIMO.OpenDocument.Pdf.csproj` | `OfficeIMO.OpenDocument.Converters.Tests/OpenDocumentPdfConversionContracts.cs#OdtFacadePreservesProjectionLossAndProducesReadablePdf` | The direct adapter merges ODT-to-Word feature mappings with Word-to-PDF diagnostics in PdfDocumentConversionResult. |
| opendocument-spreadsheet-via-excel | ods, ots | direct-loss-aware-adapter | `OfficeIMO.OpenDocument.Pdf/OfficeIMO.OpenDocument.Pdf.csproj` | `OfficeIMO.OpenDocument.Converters.Tests/OpenDocumentPdfConversionContracts.cs#OdsFacadeUsesExcelPdfEngineAndExposesInformationEvidence` | The direct adapter merges ODS-to-Excel feature mappings with Excel-to-PDF diagnostics in PdfDocumentConversionResult. |
| opendocument-presentation-via-powerpoint | odp, otp | direct-loss-aware-adapter | `OfficeIMO.OpenDocument.Pdf/OfficeIMO.OpenDocument.Pdf.csproj` | `OfficeIMO.OpenDocument.Converters.Tests/OpenDocumentPdfConversionContracts.cs#OdpFacadeUsesPowerPointPdfEngineAndKeepsAnimationLoss` | The direct adapter merges ODP-to-PowerPoint feature mappings with PowerPoint-to-PDF diagnostics in PdfDocumentConversionResult. |
| email-document | eml, msg, tnef | not-yet-direct | — | — | A premium route still needs deterministic body selection, inline-resource resolution, attachment policy, and combined source/PDF diagnostics. |
| epub-book | epub | not-yet-direct | — | — | A premium route still needs book-level chapter order, stylesheet/font/image resolution, navigation, pagination, and combined diagnostics. |
| visio-diagram | vsdx, vssx, vstx | not-yet-direct | — | — | A premium route should preserve vector page output and hyperlinks; raster-only composition is not advertised as a direct Visio PDF converter. |
