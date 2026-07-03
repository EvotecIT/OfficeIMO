# OfficeIMO Legacy XLS Corpus

This folder is for small, reviewable real-world `.xls` fixtures that protect
legacy binary workbook import behavior.

For each `sample.xls`, keep an approved `sample.import-report.md` generated from
`LegacyXlsImportReport.ToMarkdown()`. The corpus test compares the current import
report against the approved baseline so parser, projection, diagnostics, and
preserve-only signals cannot drift silently.

The corpus also keeps `projection-gap-summary.md`, which aggregates unsupported
projection gap counts by fixture, kind, and detail across all checked-in normal
fixtures.

To refresh approved baselines after an intentional import change:

```powershell
$env:OFFICEIMO_UPDATE_LEGACY_XLS_CORPUS_BASELINES = '1'
dotnet test .\OfficeIMO.Excel.Tests\OfficeIMO.Excel.Tests.csproj --filter "FullyQualifiedName~LegacyXls_Corpus_Fixtures_MatchApprovedImportReports"
dotnet test .\OfficeIMO.Excel.Tests\OfficeIMO.Excel.Tests.csproj --filter "FullyQualifiedName~ProjectionGapSummary"
Remove-Item Env:\OFFICEIMO_UPDATE_LEGACY_XLS_CORPUS_BASELINES
```

Keep fixtures focused and document their source or generator in a short note next
to the workbook when possible. Do not include sensitive customer data.

Fixtures that are expected to produce import errors or hard file-format blockers
belong in the sibling `LegacyXlsDiagnosticCorpus` folder instead of this normal
open/import/open corpus.

## Optional Desktop Excel Validation

When Microsoft Excel is installed on Windows, run the opt-in COM lane to generate
a real BIFF8 `.xls` workbook through Excel, verify Excel opens it, import it
through OfficeIMO, save the projected `.xlsx`, and verify Excel opens the result:

```powershell
$env:OFFICEIMO_RUN_LEGACY_XLS_COM_VALIDATION = '1'
dotnet test .\OfficeIMO.Excel.Tests\OfficeIMO.Excel.Tests.csproj --filter "FullyQualifiedName~LegacyXls_ComGeneratedWorkbook_ImportsAndOpensInDesktopExcelWhenRequested"
Remove-Item Env:\OFFICEIMO_RUN_LEGACY_XLS_COM_VALIDATION
```

The same switch also validates checked-in corpus fixtures by opening each source
`.xls` through Excel, importing it through OfficeIMO, saving the projected
`.xlsx`, and opening the result through Excel:

```powershell
$env:OFFICEIMO_RUN_LEGACY_XLS_COM_VALIDATION = '1'
dotnet test .\OfficeIMO.Excel.Tests\OfficeIMO.Excel.Tests.csproj --filter "FullyQualifiedName~LegacyXls_CorpusFixtures_OpenBeforeAndAfterImportInDesktopExcelWhenRequested"
Remove-Item Env:\OFFICEIMO_RUN_LEGACY_XLS_COM_VALIDATION
```

## Optional Public Samples

Downloaded `.xls` samples are useful for local broadening, but should not be
committed until their license and provenance are reviewed. Good starting points
are public sample-file sites and openly licensed preservation corpora such as
OpenPreserve `format-corpus`. After adding an approved fixture, generate its
`*.import-report.md` baseline with the command above and keep a short source note
beside the workbook.

The checked-in public buckets currently include:

- `openpreserve-format-corpus`: CC0 public-domain preservation fixtures.
- `apache-poi-testdata`: Apache POI Apache-2.0 spreadsheet test-data fixtures
  pinned to a specific upstream commit for formula-token, comment, hyperlink,
  AutoFilter, and style coverage.
