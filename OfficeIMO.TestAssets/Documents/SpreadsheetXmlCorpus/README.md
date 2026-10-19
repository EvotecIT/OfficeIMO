# XML Spreadsheet data interchange evidence

`officeimo-data.xml` is OfficeIMO-authored XML Spreadsheet 2003 data containing
typed cells, a date number format and seven constant-rate payment formulas.
The formula caches are deliberately zero before the independent consumer runs.
`libreoffice-readback.xlsx` is the exact LibreOffice 26.8.1.1 read/save result.
`provenance.json` records the producer, command, qualification limits and hashes.

The readback retains the text, numeric data, date serial and formulas, and
provides independently calculated caches for IPMT/PPMT, including payments due
at the beginning of the period. LibreOffice omits the Boolean at C1; this file
does not qualify Boolean interoperability. Native Microsoft Excel XML read/save
is not qualified by these fixtures. The XML file is an OfficeIMO output, not an
independent XML producer fixture.

Run the existing `ExcelSpreadsheetXmlTests` and `ExcelPaymentSplitTests` through
`OfficeIMO.Excel.Tests`. Reproduce the external consumer with an installed
LibreOffice in a disposable output directory and an isolated user profile.
LibreOffice is validation-only and is not required by OfficeIMO.Excel.
