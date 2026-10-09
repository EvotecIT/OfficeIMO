# DBF fixture provenance

These small fictional tables are independent inputs to the OfficeIMO integration
tests. They contain no customer or personal records.

`plain.dbf` is produced by `generate_plain.py` using Python `dbf` 0.99.11 (BSD),
then decoded independently with `dbfread` 2.0.7 (MIT). `plain-manifest.json`
records the active rows and SHA-256 digest. It covers a dBASE III table without
memo fields, accented text, literal Markdown/HTML punctuation, blank values and
a deleted physical record.

The `db3`, `fp` and `vfp` table/memo pairs and `manifest.json` come from
`DbaClientX.Dbf.Tests/Fixtures` in the [DbaClientX repository](https://github.com/EvotecIT/DbaClientX).
Its `generate.py` is the source of truth. The manifest records producer inputs,
independent decoding, deleted records and file hashes. They cover dBASE III/DBT,
FoxPro 2/FPT and Visual FoxPro/FPT. `dbfread` does not fully interpret Visual FoxPro
nullable and binary-character flags; those assertions use the recorded producer
inputs rather than claiming independent decoder agreement.

Normal tests read the checked-in files and require no Python runtime. Regenerate
fixtures only in an isolated environment with the exact tool versions above.
They qualify these profiles and conversion scenarios, not acceptance by native
dBASE or FoxPro applications or every xBase generation.
