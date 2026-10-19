# Native workbook preservation controls

The XLSM control contains a new inert VBA project: host identity attributes and
comments, with no procedures or event handlers. Its worksheet/chart content is
an OfficeIMO projection of the repository's CC0 OpenPreserve `valid.xls` fixture.
The original VBA was replaced. Native Excel opened the corrected XLSM without
repair and saved the paired XLSB and XLS files with macros disabled.

`provenance.json` records exact hashes, source attribution and the native encoder.
These files qualify a preservation cohort; they do not establish every legacy
workbook feature, chart appearance, or macro execution behavior. Fixture tools
are independent validation and are not product runtime requirements.
