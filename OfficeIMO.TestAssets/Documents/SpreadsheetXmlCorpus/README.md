# XML Spreadsheet data interchange evidence

`officeimo-data.xml` is OfficeIMO-authored XML Spreadsheet 2003 data containing
typed cells, a date number format and seven constant-rate payment formulas.
The formula caches are deliberately zero before the independent consumer runs.
`libreoffice-readback.xlsx` is the exact LibreOffice 26.8.1.1 read/save result.
`provenance.json` records the producer, command, qualification limits and hashes.

The readback retains the text, numeric data, date serial and formulas, and
provides independently calculated caches for IPMT/PPMT, including payments due
at the beginning of the period. LibreOffice omits the Boolean at C1; this file
does not qualify Boolean interoperability. The XML file is an OfficeIMO output,
not an independent XML producer fixture.

Run the existing `ExcelSpreadsheetXmlTests` and `ExcelPaymentSplitTests` through
`OfficeIMO.Excel.Tests`. Reproduce the external consumer with an installed
LibreOffice in a disposable output directory and an isolated user profile.
LibreOffice is validation-only and is not required by OfficeIMO.Excel.

`excel-financial-results.xlsx` and `excel-financial-coercions.xlsx` are exact
Microsoft Excel for Mac 16.113.4 (26100421) saved controls. The initial XLSX files
were authored through the public `ExcelDocument` API with zero caches, opened
without repair, explicitly recalculated with F9 and saved through the native UI.
The fixtures cover RATE/IRR/XIRR/XNPV numeric results and errors, multiple roots,
negative returns, fractional dates, ranges, array constants, blank positions and
nonzero payment types. `ExcelCashFlowTests` recalculates and compares these saved
results, with a relative tolerance of 1e-8 (absolute 1e-8 near zero).

`excel-financial.xml` is the first control saved by native Excel as XML
Spreadsheet 2004, the Mac name for the XML Spreadsheet family. It independently
qualifies XML import of the included formulas and caches. It does not qualify
OfficeIMO XML export, Boolean exchange, rich appearance or arbitrary producers.
`native-provenance.json` records the native build, method, scope and SHA-256 hashes.


`excel-geometry.xml` and `excel-geometry.xlsx` are exact native Excel controls
with explicit column widths, a two-dimensional merge and a comment. The XML
widths 48/96/192 points correspond to stored XLSX widths 8/16/32 in this native
cohort. The codec test supplies that observed eight-pixel digit width explicitly.
The installed bundled Calibri font measures 7.43359375 pixels per digit at
11 points/96 DPI; platform pixel rounding is a separate qualification boundary.
The third financial formula in this control has a native cache near zero for
an extreme XIRR guess despite a root of 0.2. It is retained as an observation,
not used as a numerical correctness oracle.

`excel-preservation.xml` was saved by native Excel from an OfficeIMO-authored
styled XLSX control. `officeimo-preservation.xml` is the corrected OfficeIMO
XML export, opened in the same native Excel build without repair and saved as
`excel-preservation-readback.xlsx`. The saved readback retains the tan solid
fill, medium dashed bottom border, bold/italic Calibri, 135-degree Open XML
rotation, formula hiding, comment, merge, hyperlink tooltip, scoped names and
ordinary print/view settings. Native Excel quantizes the requested font colour
`#123456` to indexed colour 56 on XML import; exact custom-font-colour fidelity
is not qualified. The XML retains the requested RGB. These finite controls do
not establish general rendered appearance, rich-text or arbitrary-producer parity.
