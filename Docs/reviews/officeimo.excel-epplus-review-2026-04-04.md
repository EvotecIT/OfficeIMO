# OfficeIMO.Excel Review - 2026-04-04

Branch: `codex/review-officeimo-excel`

## Scope

- Reviewed `OfficeIMO.Excel` with focus on workbook lifecycle, read/write parity, reader robustness, and API behavior that affects credibility against EPPlus.
- Cross-checked implementation details against the Excel-focused test suite and the package README/feature positioning.
- Ran:
  - `dotnet build OfficeIMO.Excel/OfficeIMO.Excel.csproj --no-restore`
  - `dotnet test OfficeIMO.Tests/OfficeIMO.Tests.csproj --filter "FullyQualifiedName~Excel" --no-restore`

## Validation Notes

- `OfficeIMO.Excel` builds cleanly with `0` warnings.
- Excel-focused tests passed on all target test frameworks used by `OfficeIMO.Tests`:
  - `net10.0`: passed `317`, skipped `1`
  - `net8.0`: passed `317`, skipped `1`
  - `net472`: passed `316`, skipped `1`

## Key Findings

### 1. Dispose-time autosave failures are silently swallowed

`ExcelDocument.DisposeAsync()` suppresses exceptions from both `WorkbookRoot.Save()` / `_spreadSheetDocument.Dispose()` and the later package copy-back step. That means a caller can use `autoSave: true`, dispose successfully from the caller's point of view, and still lose changes without any signal.

Relevant code:

- `OfficeIMO.Excel/ExcelDocument.cs:1689-1699`
- `OfficeIMO.Excel/ExcelDocument.cs:1723-1731`

Impact:

- Auto-save is not trustworthy under IO pressure, locked files, stream write failures, or package corruption.
- Silent failure is especially dangerous for stream-backed workflows and service code that relies on `using`/`await using`.
- This is below the bar for an EPPlus-class library, where lifecycle failures must be explicit.

Recommended fix:

- Stop swallowing dispose-time save failures for editable documents.
- Consider a split between best-effort cleanup and fail-fast persistence:
  - cleanup exceptions may be suppressed
  - save / copy-back exceptions should be surfaced
- Add regression tests that force dispose-time persistence failures for both file and stream-backed documents.

### 2. `ExcelDocumentReader.Open(...)` is less resilient than `ExcelDocument.Load(...)`

The main workbook load path normalizes package content types before opening and wraps known Open XML package failures with a more actionable message. `ExcelDocumentReader.Open(...)` skips that path and calls `SpreadsheetDocument.Open(...)` directly.

Relevant code:

- Normalized load path: `OfficeIMO.Excel/ExcelDocument.cs:641-671`
- Reader open path: `OfficeIMO.Excel/Read/ExcelDocumentReader.cs:26-28`

Impact:

- A workbook that opens successfully through `ExcelDocument.Load(...)` can still fail through the reader-only API.
- This creates inconsistent behavior across the package's own public surfaces.
- Read-only helpers become a support burden because the "lighter" API is actually less robust.

Recommended fix:

- Route `ExcelDocumentReader.Open(...)` through the same normalization/open pipeline used by `ExcelDocument.Load(...)`, or factor that pipeline into a shared internal opener.
- Add a regression test using a workbook with a repaired/normalized content-type issue and assert both APIs behave the same way.

### 3. Duplicate or normalized-colliding headers lose data in object reads and can throw in `DataTable` reads

`ReadObjects(...)` stores row values in a dictionary keyed by header text, so duplicate headers overwrite earlier columns. `ReadRangeAsDataTable(...)` adds columns directly by normalized header name, so duplicate headers can throw `DuplicateNameException`.

Relevant code:

- `OfficeIMO.Excel/Read/ExcelSheetReader.Range.cs:67-72`
- `OfficeIMO.Excel/Read/ExcelSheetReader.Range.cs:96-105`

Impact:

- Real-world spreadsheets frequently contain repeated headers such as `Value`, `Notes`, `Unnamed`, or headers that normalize to the same text after whitespace cleanup.
- Current behavior either drops columns silently or fails entirely, which is hard to debug and easy to miss in production imports.
- This weakens the package's "typed and ergonomic" read story.

Recommended fix:

- Introduce deterministic header disambiguation, for example `Header`, `Header_2`, `Header_3`.
- Apply the same policy consistently across dictionary, editable-row, typed-object, and `DataTable` readers.
- Add coverage for repeated headers, blank headers, and whitespace-normalized collisions.

### 4. Date-format detection is too naive and can misread numeric cells as `DateTime`

`StylesCache.LooksLikeDateFormat(...)` classifies a format as date-like if the raw format string contains any `d`, `y`, `h`, or `s`, or certain `m` combinations. That does not exclude quoted literals, escaped characters, or bracketed sections, so numeric formats containing words like `"days"` or `"hrs"` can be misclassified as dates.

Relevant code:

- `OfficeIMO.Excel/Read/Helpers/StylesCache.cs:33-49`

Impact:

- Reader output can be semantically wrong even when the workbook is valid.
- Bugs here are subtle because they only appear on import and only for certain custom number formats.
- Wrong typed values are more damaging than parse failures because downstream code trusts them.

Recommended fix:

- Replace the heuristic with a token-aware parser that ignores quoted text, escaped characters, and bracket sections before deciding whether date/time tokens are actually present.
- Add read tests for custom formats such as `0 "days"` and `[h]:mm` to confirm correct classification.

## Competitive Gaps vs EPPlus

The current README feature matrix is solid for common authoring/reporting workflows, but it is still narrower than what users expect from EPPlus-class libraries. The biggest gaps are not basic worksheet creation; they are trust, breadth, and proof.

### Product Gaps

- Formula engine and recalculation controls need to be first-class if the library wants to compete for server-side workbook generation beyond simple cached formulas.
- Large-workbook story still needs clearer streaming guidance and benchmarks for both reading and writing.
- Styling needs a stronger higher-level model: named styles, richer conditional formatting coverage, and clearer style reuse/composition.
- Workbook-level features such as encryption/password support, workbook protection, richer chart/pivot APIs, and import/export fidelity need a more explicit roadmap and parity matrix.

### Credibility Gaps

- There is no Excel-specific benchmark harness or public comparison data against EPPlus/ClosedXML for representative workloads.
- There is no compatibility matrix showing which Excel features are create-only, round-trip-safe, editable, or inspectable.
- The test suite is broad, but it should grow around ugly real-world workbooks rather than mainly controlled synthetic scenarios.

## Suggested Next Steps

1. Fix the four correctness findings above first, because they affect trust more than adding another feature.
2. Add an Excel competitor backlog that separates:
   - correctness bugs
   - import fidelity gaps
   - feature parity gaps
   - performance/scale gaps
3. Create an Excel benchmark project with repeatable scenarios:
   - large flat exports
   - styled reports
   - table-heavy workbooks
   - read/transform/write pipelines
   - comparison runs against EPPlus and ClosedXML
4. Publish a feature/parity matrix that is honest about current coverage:
   - supported
   - partially supported
   - round-trip only
   - not yet supported
5. Strengthen workbook corpus testing:
   - malformed but recoverable packages
   - duplicate/messy headers
   - custom number formats
   - large shared-string tables
   - chart/pivot/table round-trips created outside OfficeIMO
6. Decide where OfficeIMO.Excel wants to differentiate instead of only matching EPPlus:
   - safer APIs and deterministic saves
   - cleaner typed-read ergonomics
   - better inspection/snapshot tooling
   - more predictable cross-platform behavior

## Recommended Order of Work

1. Lifecycle trust fixes: dispose-time autosave and reader/open-path parity.
2. Read-path correctness: duplicate headers and date-format classification.
3. Competitive proof: benchmarks, corpus tests, and published parity docs.
4. Feature expansion chosen deliberately from the proof gaps rather than anecdotal demand.
