# SYLK and DIF producer fixtures

`libreoffice-cells.slk` and `libreoffice-cells.dif` are independent LibreOffice exports of the authored CSV below. They contain three rows and three columns, semicolons, embedded quotes, stored text values and a hyperlink formula result. They were produced with LibreOfficeDev `26.8.0.0.alpha0`, revision `2c87e51eeaa2b413ff4ae097b2705eea1995d8e5`. The data is synthetic and has no personal or licensed document content.

```csv
Label,Amount,Flag
"Semi;quoted ""value""",42.25,true
"=HYPERLINK(""https://example.invalid"")",-3.5,false
```

The CSV import used this producer's default settings, so amounts and flags are text in the exports. LibreOffice interpreted the leading equals sign in the third row as a formula: the SYLK export retains its `E` expression and stored URL, while the DIF export retains the stored URL alone. Tests assert these source distinctions, quoted text, row/cell identities and Reader table content. Separate parser cases cover typed numbers and booleans, literal formula-looking text, multiline DIF strings, corruption, limits and XLSX reopening.

Reproduce the files with an isolated LibreOffice profile and the same producer revision:

```sh
soffice --headless -env:UserInstallation=file:///absolute/scratch/lo-profile --convert-to 'slk:SYLK' --outdir /absolute/scratch/slk cells.csv
soffice --headless -env:UserInstallation=file:///absolute/scratch/lo-profile --convert-to 'dif:DIF' --outdir /absolute/scratch/dif cells.csv
```

| Fixture | SHA-256 |
| --- | --- |
| `libreoffice-cells.slk` | `1c52a72fd1dc30d923ec17651336f0ce166f66b349c5cf227fbb9b66a97aeb51` |
| `libreoffice-cells.dif` | `4122e40939cc626e40d4aa608364a0097cee09b337f99be536c30aa18b1104d2` |

LibreOffice is an opt-in evidence producer. It is not required by the importer or its ordinary test suite. This pair qualifies the documented stored-value profile for one producer; it does not qualify every historical SYLK/DIF dialect, formula preservation or source-format writing.
