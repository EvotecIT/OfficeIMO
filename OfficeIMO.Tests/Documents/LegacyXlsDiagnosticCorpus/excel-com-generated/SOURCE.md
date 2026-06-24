# Excel COM Generated Diagnostic XLS Fixtures

Generated locally with Microsoft Excel desktop automation on 2026-06-24 for
OfficeIMO legacy XLS diagnostic coverage.

The workbooks are synthetic test data authored by EvotecIT for this repository.
`encrypted-password.xls` is a BIFF8 workbook saved by Excel with password-to-open
protection. The password is `openpass`; the fixture is not sensitive and exists
only to verify that FilePass-encrypted legacy workbooks are reported as explicit
file-format blockers instead of being partially imported.

`biff5-workbook.xls` is a small Excel 5.0/95 workbook saved by Excel COM with
file format id 39. It verifies that pre-BIFF8 workbooks are reported as explicit
unsupported-version blockers before BIFF8-specific record layouts are interpreted.

| File | SHA256 |
| --- | --- |
| `biff5-workbook.xls` | `63743E8CCE1C366F57F4C79FBCFFDE4850E0AAC72DA07F15748A72B2F7911B64` |
| `encrypted-password.xls` | `0DDCE5118441DC43FD032A9AF08D45E5E4BBD13E8C88349EE7E22A2810349F46` |
