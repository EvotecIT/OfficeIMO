# Themed table row bands

The two themed report PDFs were generated with OfficeIMO at `c000d5ce86`. They preserve the inputs that exposed missing trailing rows, a duplicated risk-table header, and wrapped mitigation text moving into the next row. The manifest records the exact bytes; keep the fixtures unchanged when the writer changes.

`word-themed-report.pdf` contains a seven-row revenue table (including its header and total) and a five-row risk register. `native-wrapped-report.pdf` contains the header and four signal rows: DMARC, TLS, DNS and PDF. Poppler text extraction confirms that the source PDFs contain the full rows and wrapped cell text. The reader tests protect table membership and row ownership; plain text containing the missing values is insufficient.

The report PDFs are OfficeIMO-generated regression evidence. All fixtures are test inputs and are not copied into product packages.

`form-layout.pdf` is the LibreOffice export of the audit's employee equipment-loan form. It guards the preceding identity row when a wrapped candidate overlaps a later form grid. The fixture is test-only and comes from the same authored audit corpus.
