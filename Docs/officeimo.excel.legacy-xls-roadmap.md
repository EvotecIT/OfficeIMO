# OfficeIMO.Excel Legacy XLS Roadmap

This roadmap tracks native legacy `.xls` support without adding spreadsheet dependencies.
The intent is to grow support through a clean legacy workbook model instead of wiring BIFF
records directly into `ExcelDocument`.

## Architecture Boundary

Legacy `.xls` import follows this pipeline:

```text
OLE compound file
  -> BIFF record stream
  -> Legacy XLS workbook model
  -> OfficeIMO.Excel projection
  -> normal OfficeIMO .xlsx editing and save APIs
```

The parser owns binary format details. The legacy model owns workbook concepts. The
projector owns mapping decisions into OfficeIMO's Open XML workbook model.

## Phase 1 - Foundation

- Detect true OLE compound `.xls` files and locate the `Workbook` or `Book` stream.
- Read BIFF record headers safely with length and corruption diagnostics.
- Parse workbook globals needed for tabular import: `BOF`, `EOF`, `BoundSheet8`, and `SST`.
- Parse initial worksheet records: `NUMBER`, `LABELSST`, `LABEL`, `BOOLERR`, `BLANK`.
- Project parsed cells into a normal `ExcelDocument`.

Current implementation reads the workbook stream from regular FAT chains, MiniFAT
mini-stream chains below the compound-file mini stream cutoff, and extended FAT
sector lists that continue through DIFAT sectors. Corrupt container headers,
sector chains, and missing workbook streams are reported as import diagnostics
instead of escaping as raw parser failures.

## Phase 2 - Value Fidelity

- Add compact numeric records: `RK` and `MULRK`.
- Add repeated label/blank records and richer boolean/error handling.
- Support date system detection and serial date projection policies.
- Expand string parsing for rich text and phonetic/extended string payloads.

Current implementation covers `RK`, `MULRK`, `MULBLANK`, inline `LABEL`, BIFF
error-code text mapping, and `Date1904` detection. Shared string tables continued
across `Continue` records are imported when the continuation boundary falls between
complete string entries, and when a continuation splits inside the character data
of an `XLUnicodeRichExtendedString` segment. Shared strings carrying rich-text run
metadata and extended/phonetic payload blocks import their plain text while skipping
the variable metadata payload. Numeric cells with date-like BIFF
formats project as dates, including conversion from 1904-system serials to the
equivalent Open XML date serial.

## Phase 3 - Formatting And Layout

- Parse `FORMAT`, `XF`, `FONT`, palette, row heights, column widths, hidden rows/columns,
  merged ranges, freeze panes, and basic sheet visibility.
- Map supported formatting to OfficeIMO styles.
- Report unsupported formatting as diagnostics rather than silently dropping it.

Current implementation covers layout metadata for `COLINFO`, `ROW`, `MERGECELLS`,
`WINDOW2` frozen panes with `PANE`, and `BoundSheet8` visibility. These project to
OfficeIMO column width/hidden state, row height/hidden state, merged ranges, frozen
panes, hidden sheets, very-hidden sheets, worksheet gridline visibility,
right-to-left view state, and row/column heading plus zero-value view visibility.
Default empty-row height and default column width import
through the normal OfficeIMO sheet-format API, and row/column outline levels with
collapsed states import through normal OfficeIMO sheet outline APIs. It also parses style table basics from `FORMAT` and `XF`,
resolves built-in and custom number formats, detects date-like numeric cells, and
projects number formats and serial dates through normal OfficeIMO cell/style APIs.
`FONT` records are parsed for name, point size, weight, italic, underline,
strikethrough, and superscript/subscript escapement, and XF font references
project through the normal OfficeIMO style model. Palette records
and `IcvFont` values are resolved for font color projection, including custom palette
entries and the default BIFF color table. Solid and patterned XF fills project
through normal OfficeIMO style APIs. Horizontal/vertical alignment, wrap text,
rotation, indentation, shrink-to-fit, and reading order are projected from XF
alignment fields. XF side and diagonal borders project with palette-backed colors
and mapped border styles. XF locked and formula-hidden protection flags project
through normal Open XML cell style protection, and XF quote-prefix markers project
to Open XML quote-prefix styles. Cell XF records now preserve their raw parent
style index and apply-facet bits in the legacy model, and projection resolves
inherited number format, font, fill, alignment, border, and protection facets from
parent style XF records before writing normal OfficeIMO/Open XML styles. `COLINFO` default XF indexes now project as real
Open XML column style definitions, and formatted `ROW` default XF indexes project as
real Open XML row style definitions. Additional edge-case formatting remains a
later Phase 3 hardening slice.

## Phase 4 - Workbook Features

- Import formulas as formulas where token decoding is supported and as cached values otherwise.
- Add hyperlinks, comments, named ranges, print settings, protection metadata, and basic sheet settings.
- Add inspection/preflight output for features that are import-only or preserve-only.

Current implementation imports `Formula` cached results for numeric, boolean, error,
blank-string, and following-`String` text results. It also decodes a first scoped
set of BIFF formula tokens into live Open XML formulas for same-sheet references,
areas, relative `PtgRefN`/`PtgAreaN` references resolved from the formula cell,
workbook-internal 3D `PtgRef3d`/`PtgArea3d` references resolved through
`ExternSheet`, external-workbook 3D cell/range references when the supporting
`SupBook` sheet table is present, invalid references through `PtgRefErr`, `PtgAreaErr`,
`PtgRefErr3d`, and `PtgAreaErr3d`, numeric and Boolean constants, arithmetic/comparison operators,
percent, explicit parentheses, reference operators (`PtgIsect`, `PtgUnion`, and
`PtgRange`), string literal `PtgStr` operands for concatenation formulas, error
constants through `PtgErr`, missing function arguments through `PtgMissArg`,
same-workbook defined-name operands through `PtgName`, external defined-name and
add-in function-name operands through `PtgNameX`,
fixed-arity `PtgFunc` calls such as `ROUND`, `AND`, and `OR`, and scoped aggregate `PtgFuncVar`
calls such as `SUM`. Add-in user-defined functions encoded as `PtgFuncVar` with
the legacy `0x00FF` function id now decode when the leading argument resolves to
a supported external name. Conditional `IF` formulas using `PtgAttrIf`, `PtgAttrGoto`,
and `PtgFuncVar`, plus `CHOOSE` formulas using `PtgAttrChoose`, now decode to
normal Open XML formulas. Optimized `PtgAttrSum` formula attributes now decode to
normal `SUM(...)` formulas during Open XML projection, while `PtgAttrSpace` and
`PtgAttrSpaceSemi` display tokens are consumed and normalized away.
Unsupported token streams still import as cached values and now report
detail-coded formula-token diagnostics when unsupported-record reporting is enabled.
It also imports external
URL `HLink` records that use Office shared URL monikers and projects them through
normal OfficeIMO hyperlink relationships without rewriting the linked cell value.
Location-only same-workbook `HLink` records now project as normal Open XML internal
hyperlink locations without creating external relationships. Absolute local, UNC,
and relative file `HLink` records saved with Office shared `FileMoniker` data now
project as normal external file hyperlink relationships. Composite/item monikers
and other hyperlink target shapes remain diagnostic-only. Workbook- and worksheet-level
`Protect` and `Password` records now project as Open XML protection metadata,
preserving the legacy 16-bit password verifier when present. This is workbook/sheet
UI protection metadata only, not password-to-open file encryption. Basic print page
setup now imports margin records plus `Setup` scale, fit width/height, orientation,
and header/footer margins through the normal OfficeIMO page setup APIs. Print
options for row/column headings, printed gridlines, and horizontal/vertical page
centering import through the normal OfficeIMO print-options API. Manual row and
column page breaks import through the normal OfficeIMO manual page-break APIs.
Worksheet view zoom imports from `Scl` records through the normal OfficeIMO sheet
view API. `Header` and `Footer` records now import raw Excel section/token strings
into the legacy model and project them through the normal OfficeIMO header/footer
API. The first defined-name slice imports `Lbl` records backed by single 3D
cell/range references, including sheet-local names, hidden names, and built-in
`Print_Area`, through normal OfficeIMO defined-name and print-area APIs. Built-in
`Print_Titles` names backed by 3D whole-row/whole-column references and `PtgUnion`
now import through the normal OfficeIMO print-title API. Hidden built-in
`_FilterDatabase` names backed by scoped 3D areas now project through the normal
OfficeIMO AutoFilter API while retaining the imported defined-name metadata. Plain
legacy cell comments backed by `Note`, note-type `Obj`, `TxO`, and `Continue`
records now import through the normal OfficeIMO comment API. TxO formatting-run
boundaries and font indexes are preserved in the legacy comment model, and supported
font properties project through the normal OfficeIMO rich comment API. Comment object geometry remains future hardening work. Worksheet `DIMENSIONS` records
now import as explicit legacy-model declared used-range metadata, including empty
worksheet declarations. Open XML projection continues to let OfficeIMO compute the
saved `.xlsx` sheet dimension from projected cells and structures to avoid stale
dimension repair prompts. `ExcelDocument.LoadLegacyXlsWithReport(...)` now returns
the projected `ExcelDocument` together with the parsed `LegacyXlsWorkbook`,
diagnostics, and unsupported-feature report from the same parse so callers can
preflight import-only and preserve-only content without leaving the normal OfficeIMO
document path. The result exposes `EnsureNoImportErrors()` and
`EnsureNoUnsupportedFeatures()` guards for corpus and CI checks.
Simple BIFF `DVal`/`Dv` whole-number, decimal, date, time, and text-length data
validation rules with constant numeric formulas, plus inline literal, same-sheet
range-backed, workbook defined-name-backed, and sheet-local defined-name-backed
list validations, cross-sheet range-backed list validations, and decodable custom
formula validations now import into the legacy model and project through the
normal OfficeIMO data validation APIs, including prompt and error message
metadata. More complex formula shapes remain preserve-only diagnostics for later
slices. Classic BIFF `CondFmt`/`CF` conditional formatting rules now import for
simple cell-value comparison rules and formula-expression rules when the rule
formulas are decodable by the existing BIFF formula reader. These project through
the normal OfficeIMO conditional-formatting APIs. Differential formatting payloads
and richer visual rule formatting remain future hardening work. Simple BIFF
`AUTOFILTERINFO`/`AUTOFILTER` criteria import now covers decodable string
equality and numeric/string custom comparisons, stores them in the legacy model,
and projects through the normal OfficeIMO AutoFilter APIs.

## Phase 5 - Compatibility Corpus And Preservation

- Build a corpus of real-world `.xls` files with expected import diagnostics.
- Add feature reports for macros, charts, pivots, OLE objects, external links, and unsupported BIFF records.
- Keep `.xls` write support out of scope until the read/import model is mature.

Current implementation has the first diagnostics contract: `FilePass` encrypted
workbooks stop import with an explicit unsupported-encryption error, while unsupported
hyperlink, worksheet drawing/object, PivotTable, and chart records are reported with
feature-specific diagnostic codes when unsupported-record reporting is enabled.
Legacy AutoFilter control and criteria records outside the supported simple
criteria subset are likewise reported as preserve-only unsupported filter
criteria; this is separate from `_FilterDatabase` defined-name range projection,
which continues to use the normal OfficeIMO AutoFilter API.
BIFF `DVal` and `Dv` records outside the supported simple constant-formula,
inline-list, same-sheet and cross-sheet range-list, defined-name-list, and
decodable custom-formula validation subset are reported as preserve-only data
validation features so corpus tooling can flag validation rules before broader
native projection is implemented.
BIFF conditional-formatting records outside the supported simple classic
`CondFmt`/`CF` rule subset are reported as preserve-only conditional-formatting
features. Extended `CF12`, `CFEx`, and `DXF` records remain diagnostics-only until
their richer rule and differential-formatting models are implemented.
`BoundSheet8` entries for macro sheets, chart sheets, and VBA module sheets are also
reported as explicit feature diagnostics while worksheet sheets continue to import.
Unsupported sheet substreams are scanned for preserve-only feature records, so
chart-sheet-local chart and drawing records now appear with their sheet location
without projecting chart sheets as worksheets or importing chart definitions.
Unsupported sheet entries are preserved in the legacy workbook model, and `WsBool`
dialog sheets are classified as unsupported dialog sheets instead of being projected
as normal worksheets. OLE compound containers with VBA project storage are now
reported as preserve-only macro content independently of BIFF sheet entries, so
corpus tooling can flag macro-enabled legacy workbooks even when the workbook
stream itself imports cleanly. OLE compound containers with embedded object pool
storage are likewise reported as preserve-only embedded OLE object content.
`SupBook` supporting links are preserved as external-reference
metadata for external workbook, add-in, DDE/OLE, same-sheet, self, and unused links,
including parsed external/add-in name tables from supported `ExternName` records,
with external-link diagnostics for unsupported external-reference records. External
references that require full Open XML external-link package parts still remain a
projection gap; supported formulas currently project as text formulas using the
information available in the legacy model.
Unsupported and preserve-only feature occurrences now also populate a structured
`LegacyXlsWorkbook.UnsupportedFeatures` report with stable codes, feature kind,
sheet name, record type, record offset, and stable feature-detail keys such as
`Chart:Chart`, `Drawing:MsoDrawing`, `PivotTable:SxView`, and
`Compound:VbaProjectStorage` so corpus tooling can reason about unsupported
content without parsing diagnostic text or decoding raw BIFF ids by hand.
`LegacyXlsImportReport`
now exposes compact worksheet, cell, formula, comment, hyperlink, data-validation,
conditional-formatting, AutoFilter criteria, defined-name, external-reference, diagnostic-code, and unsupported-feature counts
for corpus baselines through both `LegacyXlsWorkbook.CreateImportReport()` and
`ExcelDocument.LoadLegacyXlsWithReport(...).ImportReport`.
`OfficeIMO.Tests/Documents/LegacyXlsCorpus` now defines the optional real-world
`.xls` corpus contract: every collected workbook can carry an approved
`*.import-report.md` baseline, and the corpus test compares current import reports
against those baselines with an explicit refresh environment variable for
intentional changes. The current approved `openpreserve-format-corpus` baseline
has no remaining generic formula-token blockers, which keeps future formula work
anchored to named gaps instead of a single opaque bucket. Collecting approved
non-sensitive real-world fixtures and broader preserve-only feature modeling remain
future Phase 5 work.

## Non-Goals For The Initial Work

- No external spreadsheet dependency.
- No native `.xls` save/write path.
- No hidden conversion through Excel/COM.
- No parser shortcuts that mutate `ExcelDocument` directly.
