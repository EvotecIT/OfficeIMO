# OfficeIMO.Excel Legacy XLS Read And Native Write Roadmap

This roadmap tracks what is still missing before we can confidently advertise
legacy `.xls` read support and native `.xls` write support.

## Goal

Close this roadmap by bringing legacy BIFF8 `.xls` read behavior to the same
confidence model as the `.xlsx` reader: every workbook feature family is either
projected into the normal OfficeIMO model, readable through cached values,
preserved in place, diagnosed before conversion, or blocked with a clear error.

The same confidence model applies to native `.xls` write: every normal
OfficeIMO feature family must either serialize to BIFF8, intentionally degrade
with documented diagnostics, or fail before writing so callers do not get silent
data loss.

## Honest Support Statement

Current target wording:

> OfficeIMO.Excel can import BIFF8 legacy `.xls` workbooks through
> `ExcelDocument.Load(...)`, project supported workbook and worksheet content into
> the normal OfficeIMO Excel model, expose diagnostics for unsupported or
> preserve-only legacy features, save converted output as `.xlsx`, and write a
> first native `.xls` subset for basic worksheets, scalar cell values, worksheet
> layout and views, worksheet tab colors, sheet visibility, workbook and
> worksheet calculation metadata, number formats, basic font styles including
> BIFF-backed family, character-set, underline styles, superscript/subscript,
> outline, shadow, condense, and extend metadata,
> solid/patterned fills including gray125 and resolved theme/tint colors, BIFF-backed
> alignment, borders, cell protection styles, quote-prefix markers, rich-text
> cell runs with OfficeIMO/Open XML projection for BIFF-backed run font family,
> character-set, underline styles, superscript/subscript, outline, shadow, condense, and extend
> metadata, supported cached formula cells, internal defined names and print
> names, supported hyperlinks with ScreenTips, protection metadata,
> write-reservation metadata, document properties, page setup, supported
> worksheet comments, supported worksheet AutoFilter ranges, criteria, and
> date-group criteria written as equivalent serial-date comparison ranges, and
> supported data validation, classic conditional-formatting rules, and supported
> data-consolidation settings including start-label normalization,
> same-workbook source references, and
> workbook- and sheet-scoped named sources, plus supported external workbook link
> metadata with sheet names and custom external defined names.

Do not claim:

> OfficeIMO supports all `.xls` files.

That stronger claim would be misleading today because some file types, sheet
types, and feature families are hard blockers or diagnostics-only.

## What Parity With XLSX Reader Means

`.xlsx` loading can often keep unknown Open XML package parts around during
round-trip saves. `.xls` import cannot rely on that: once we convert to `.xlsx`,
every legacy feature must be either projected into the OfficeIMO/Open XML model
or reported clearly as unsupported or preserve-only before save.

For XLS read support to be on par with the `.xlsx` reader, each feature family
needs one of these outcomes:

- **Projected**: loaded into the normal OfficeIMO model and saved as `.xlsx`.
- **Cached**: readable enough to preserve user-visible values, usually formulas
  with cached results, while reporting missing live formula projection.
- **Diagnosed**: reported through `LoadLegacyXlsWithReport(...)`,
  `InspectFeatures()`, and corpus reports so callers can reject the conversion.
- **Blocked**: failed early with a clear error when the workbook cannot be
  imported safely.

## Known Not-All-XLS Boundaries

- BIFF8 is the supported target. Older BIFF versions are explicit blockers.
- Password-to-open encrypted XLS files are explicit blockers.
- Normal loading is content-first for true OLE compound BIFF workbooks, but
  extension-only fallback is limited to `.xls`; `.xlt`, `.xla`, `.xlm`, and
  `.xlw` are blocked native save targets, not separately advertised import
  formats until fixtures prove them.
- Normal worksheets are the projection target. Chart sheets, macro sheets,
  dialog sheets, and VBA module sheets are not normal worksheet projections.
- Native `.xls` save/write has started with a first-party BIFF8 writer for
  basic workbooks, normal worksheets, scalar text/number/boolean cells,
  explicitly styled blank cells, supported rich-text cells, supported cached
  formula cells, internal defined names and print names, worksheet layout and
  views, worksheet tab colors, sheet visibility, workbook and worksheet
  calculation metadata, write-reservation metadata, document properties, and
  number formats plus basic cell font styles, solid/patterned fills including
  gray125,
  BIFF-backed alignment, basic borders, cell protection styles, and
  quote-prefix markers. Richer workbook features are explicitly blocked before
  native write unless they are listed as supported in the matrix below.
- Unsupported BIFF/OLE content is not silently preserved into the converted
  `.xlsx`; it must be projected or diagnosed.

## XLSX/XLS Support Matrix

Legend:

- **Projected**: loaded into normal OfficeIMO APIs and saved as `.xlsx`.
- **Cached**: user-visible values are readable, but live feature semantics are
  not fully projected.
- **Diagnosed**: reported before conversion or save so callers can decide.
- **Blocked**: intentionally rejected with a clear error.

| Feature family | XLSX reader today | XLS read status | Native XLS write status | What remains |
| --- | --- | --- | --- | --- |
| Package detection | Projected | Projected for BIFF8 OLE and renamed Open XML workbooks | Normal `Save("*.xls")` and stream `LegacyXls` route to the BIFF writer | Keep future corrupt/ambiguous input cases corpus-driven. |
| Older workbook formats | Not applicable | Blocked for pre-BIFF8 | `.xlt`, `.xla`, `.xlm`, and `.xlw` save targets blocked | No broader format claim until fixtures justify it. |
| Password-to-open encryption | Supported for supported OOXML encryption paths | Blocked for XLS `FilePass` | Blocked | Native XLS decryption/encryption is out of scope unless separately approved. |
| Worksheets | Projected | Projected for normal worksheets; unsupported sheet types diagnosed | Normal worksheets written | Chart, macro, dialog, and VBA module sheets remain diagnosed/blocked. |
| Scalar values and dates | Projected | Projected | Text, number, Boolean, error, blank, styled blank, numeric dates, explicit Open XML date cells, and cached formula values written | Keep expanding date-system and unusual cached-value corpus coverage. |
| Rich text and phonetics | Projected | Projected for supported rich text including BIFF-backed run font family, character-set, underline styles, superscript/subscript, outline, shadow, condense, and extend metadata, plus worksheet-level phonetic defaults; cell phonetic guide payloads and range-scoped guide text are explicit boundaries | Rich-text cells and comments plus BIFF-backed run font family, character-set, underline styles, superscript/subscript, outline, shadow, condense, and extend metadata, and worksheet-level phonetic defaults written | Cell phonetic guide payloads, range-scoped guide text, and non-BIFF rich-text metadata remain diagnosed/blocked. |
| Styles and number formats | Projected | Projected for common BIFF8 formats, fonts including BIFF family, character-set, underline styles, superscript/subscript, outline, shadow, condense, and extend metadata, fills, borders, alignment, protection, quote prefix, and selected conditional-formatting DXF payloads; Theme and style-extension records are diagnostics/preserve-only | Supported styles written, including sparse style inheritance, BIFF-backed font family, character-set, underline styles, superscript/subscript, outline, shadow, condense, and extend metadata, gray125 pattern fills, explicit general alignment, and resolved theme/tint colors | Gradient fills, style-extension payloads, and uncommon facets remain blocked. |
| Layout, views, panes, and print metadata | Projected | Projected for common row/column, merge, pane, view, page setup, printer, tab color, option metadata, multiple workbook and worksheet window/view records, ignored-error metadata, cell watches, worksheet scenarios, data-consolidation settings including start-label normalization, same-workbook and external virtual-path data-consolidation source references, and workbook-, sheet-, and external-source named data-consolidation sources; custom/named sheet views are explicit boundaries | Current BIFF-backed subset written, including worksheet protection permission exceptions, protected ranges, ignored errors, multiple workbook and worksheet window/view records, cell watches, worksheet scenarios, data-consolidation settings including start-label normalization, same-workbook and external virtual-path data-consolidation source references, and workbook-, sheet-, and external-source named data-consolidation sources | Header/footer images, custom workbook views, custom/named sheet views, richer scenario metadata, sheet-qualified external data-consolidation source references, and named data-consolidation source references with explicit ranges remain blocked. |
| Formulas | Projected as formula text; calculation remains partial | Projected when BIFF token stream is supported; cached/diagnosed otherwise | Broad supported encoder subset written, including common functions, operators, arrays, shared formulas, internal names, contiguous 3D ranges, explicit workbook-internal 3D unions, supported external-workbook sheet references, and workbook- and sheet-scoped external defined-name operands | Future unsupported token families remain blocked. |
| Named ranges and print names | Projected | Projected for supported internal references/formulas | Supported workbook/sheet names, formula names, print areas, print titles, external-workbook sheet references, and workbook- and sheet-scoped external defined-name operands written | Unsupported built-in names remain blocked. |
| AutoFilter | Projected/partially editable | Projected for supported range/dropdown/simple criteria metadata | BIFF range, dropdown, equality/custom/blank, blank-or-single-value, top-bottom, and supported date-group criteria written; date groups save as equivalent serial-date custom comparisons | Dynamic, color, icon, blank-plus-multiple-value, larger equality-list, unsupported date-group shapes, and richer control metadata remain diagnosed or blocked. |
| Sort metadata | Projected/partially editable | Projected for BIFF `Sort` records without custom-list ordering; BIFF custom-list order indexes are parsed in the legacy model and reported as unsupported projection gaps | BIFF `Sort` records written for up to three value-based A1-reference keys, descending flags, left-to-right sort, case sensitivity, and PinYin sort method | More than three keys, color/icon/custom-list sort conditions, worksheet sort maps, and BIFF custom-list index to Open XML custom-list string reconstruction remain blocked. |
| Data validation | Projected/partially editable | Projected for supported simple rules | Inline, range, sheet-range, defined-name, external-workbook sheet reference, workbook- and sheet-scoped external defined-name operand, and scalar formula validations written | Richer formulas/shapes remain blocked. |
| Conditional formatting | Projected/partially editable | Projected for supported classic rules | Classic unstyled BIFF rules, including supported external-workbook sheet reference and workbook- and sheet-scoped external defined-name operand formulas, and minimal `CfEx` metadata written | DXF/visual rules and richer rule families remain blocked. |
| Comments and notes | Projected/partially editable | Projected for supported cell comments, rich text, visible state, and anchors | Supported comment text, rich runs, colors, authors, visibility, and anchors written | Uncommon comment object metadata remains blocked. |
| Hyperlinks | Projected/editable for supported targets | Projected for supported URL/file/internal targets and ScreenTips | Supported `HLINK` and `HLinkTooltip` records written | Unsupported monikers and unsafe schemes remain diagnosed/blocked. |
| Tables and table styles | Projected/partially editable | Default table/PivotTable style names project; legacy table definitions and custom style definitions remain diagnosed/report-only | Default table/PivotTable style names and custom workbook table-style definitions written; worksheet table parts, table definition parts, and single-cell table parts are blocked before native write | Worksheet table definitions remain outside the native XLS write subset. |
| Images, drawings, charts, and PivotTables | Projected/partially editable in `.xlsx`; some package content preserved | Diagnostics/preserve-only report model with decoded metadata for BIFF drawing/object records, OfficeArt/image-store records, chart records, chart sheets, and PivotTable records | Blocked | Read parity is a documented no-silent-loss boundary, not editable Open XML projection. Native write remains a separate boundary. |
| External links and external data | Preserved/diagnosed in `.xlsx` | Supported external workbook links, external sheet names, external defined names, and external workbook formulas project; add-in, DDE/OLE, DBQueryExt, query-table tag, and external cache metadata are diagnostics/preserve-only | Supported external-workbook sheet formula references, workbook- and sheet-scoped external defined-name operands, and simple external workbook link parts with sheet names and custom external defined names write the required BIFF `SupBook`/`ExternSheet`/`ExternName` metadata | Read parity is a documented no-silent-loss boundary for non-projected external data. Native write still blocks add-in, DDE/OLE, external caches, unsupported external-link relationship types, built-in external names, connections, and query tables before save. |
| VBA, OLE packages, form controls, rich data, signatures | Preserved or diagnosed in `.xlsx` depending on package shape | Diagnosed, not executed or projected | Blocked | Keep no-silent-loss diagnostics; do not execute VBA or silently drop signed/embedded content. |
| Document properties | Projected | Projected for supported OLE core/application/custom property types | Supported OLE property streams written | Vector, clipboard, stream/storage, and other exotic custom VARTYPEs remain diagnosed. |

## Current Evidence

This branch is no longer a planning-only branch. Current evidence is:

- Every row in the XLSX/XLS support matrix has been audited to one current read
  and native-write status: projected, cached, preserved/diagnosed, blocked, or
  written with explicit boundaries.
- `LegacyXls_NormalLoad` covers path, stream, sync, async, and converted `.xlsx`
  save/reload flows.
- The checked-in normal and diagnostic corpora have approved import-report
  baselines.
- `projection-gap-summary.md` reports zero unsupported projection gaps for the
  normal corpus and keeps hard-error fixtures in the diagnostic corpus.
- Formula corpus baselines are clean across the checked-in normal and diagnostic
  reports: zero unsupported projection gaps, formula token blocker rows, and
  chart data-source formula projection failures.
- Native formula writer coverage has no current corpus-reported in-scope token
  family left open. Supported external-workbook sheet formula and defined-name
  references now write BIFF `SupBook` and `ExternSheet` metadata, and
  workbook- and sheet-scoped external defined-name operands write BIFF `ExternName`
  metadata.
- Native external-link writer coverage now consumes simple Open XML external
  workbook link parts with sheet names and custom external defined names and
  writes the matching BIFF supporting-link metadata, while unsupported external
  relationship types remain preflight blockers.
- Native `.xls` save goes through the normal OfficeIMO save APIs for file paths
  and `ExcelStreamSaveFormat.LegacyXls` streams.
- The full `LegacyXls` test lane passes across `net472`, `net8.0`, and
  `net10.0` for the current checkpoint.
- Import/projection tests prove BIFF-backed font family, character-set, underline styles, superscript/subscript,
  outline, shadow, condense, and extend metadata
  project into Open XML font styles, and native writer tests prove save/reload through the OfficeIMO legacy reader for
  the current supported subset: scalar values, styled blanks, dates, Open XML
  formula cached dates, formulas,
  shared and array formulas, styles, rich text, names, hyperlinks, comments,
  layout, views, page setup, printer settings, workbook protection password-only
  metadata, protection permission exceptions, protected ranges, write reservation,
  document properties, AutoFilter including blank-or-single-value criteria and
  date-group criteria written as serial-date comparison ranges,
  sort metadata, BIFF-backed font family, character-set, underline styles, superscript/subscript,
  outline, shadow, condense, and extend metadata, data validation,
  conditional formatting, ignored-error metadata, cell watches, worksheet
  scenarios, data-consolidation settings including start-label normalization,
  same-workbook and external virtual-path data-consolidation
  source references, workbook-, sheet-, and external-source named data-consolidation sources,
  default and custom workbook table-style definitions, and sparse style
  inheritance from parent style formats.
- AutoFilter, sort, data-validation, and conditional-formatting read coverage is
  split into a tested supported subset plus explicit boundaries: simple and
  common Excel-authored shapes project into the OfficeIMO/Open XML model, while
  custom-list sort indexes, unsupported data-validation records/formula tokens,
  and unsupported conditional-formatting records/formula tokens are reported as
  unsupported projection gaps or preserve-only diagnostics.
- Legacy worksheet table/list definition records are a closed native-write
  boundary and documented read boundary:
  `FeatHdr11`, `Feature11`, `List12`, and `Feature12` are classified as
  table-definition preserve-only diagnostics, while default table and PivotTable
  style names continue to project from `TableStyles`. Native write keeps
  worksheet table parts, table definition parts, and single-cell table parts as
  explicit `tables` preflight blockers instead of silently dropping them.
- Drawings, embedded images, charts, chart sheets, and PivotTables have a closed
  native-write boundary and explicit XLS read contract: BIFF drawing/object
  records, OfficeArt/image-store records, chart records, chart-sheet metadata,
  and PivotTable records are decoded into diagnostics/report models and
  preserved as no-silent-loss boundaries instead of being projected as editable
  Open XML drawing, chart, or PivotTable parts. Native write blocks worksheet
  drawing/image/chart parts, direct image and 3D model relationships, chart
  sheets, and PivotTable markers before save.
- External links and external data have a closed native-write boundary and
  explicit XLS read contract: supported external workbook links, external sheet
  names, external defined names, and external workbook formulas project into Open
  XML external-link parts and formula text. Native write supports simple external
  workbook link metadata, external workbook formula references, external defined
  name operands, external virtual-path data-consolidation source references, and
  the workbook refresh marker. Add-in, DDE/OLE, DBQueryExt connection records,
  PivotTable query-table tags, XCT/CRN external cell caches, workbook
  connections, and worksheet query tables remain diagnostics, preserve-only
  report models, or `connections or query tables` preflight blockers.
- Style, theme, rich-text, phonetic, layout, and view edge cases now have an
  explicit XLS read contract: common BIFF8 styles, fonts, fills, borders,
  alignment, rich text, panes, row/column layout, worksheet views, and workbook
  windows project; `XFCRC`, `XfExt`, `StyleExt`, `Theme`, phonetic guide
  payloads, custom/named views, and other richer XLSX-only metadata are
  diagnostics/preserve-only or native-write blockers.
- Preflight tests prove unsupported workbook and worksheet feature families block
  before native write instead of being silently dropped, including digital
  signature package metadata, unsupported workbook calculation metadata,
  unsupported AutoFilter sort, extension, dropdown-control, column, criteria metadata,
  unsupported date-group shapes, dynamic, color, icon, larger equality-list, and
  blank-plus-multiple-value criteria,
  unsupported sort-state metadata,
  duplicated AutoFilter criteria containers,
  unsupported data-validation collection, extension, and IME metadata,
  duplicated data-validation formula elements and unsupported formula metadata,
  data-validation formulas outside the native XLS formula subset, missing
  required data-validation formulas, invalid data-validation ranges, too many
  data-validation ranges, and data-validation ranges outside BIFF8 limits,
  unsupported conditional-formatting pivot, extension, collection, rule, and formula metadata,
  conditional-formatting differential formats, visual payloads, unsupported
  operators and rule types, formulas outside the native XLS formula subset,
  formula-backed multi-range rules, invalid ranges, too many ranges, and ranges
  outside BIFF8 limits,
  worksheet table parts, table definition parts, single-cell table parts,
  drawing/image/chart/chart-sheet/PivotTable parts, unsupported external workbook link metadata,
  query tables/connections, OLE objects/packages, form controls,
  slicers/timelines, digital signatures, and non-comment VML shapes,
  unsupported comment object shape, fill, line, shadow, textbox, path, and client metadata,
  unsupported comment rich-text run metadata,
  oversized defined-name formula payloads,
  oversized formula token payloads,
  oversized data-validation and conditional-formatting formula token payloads,
  oversized data-validation text payloads,
  oversized data-consolidation source reference payloads,
  unsupported data-consolidation source-reference collection metadata,
  oversized worksheet scenario payloads,
  oversized protected-range payloads,
  unsupported protected-range metadata,
  unsupported ignored-error collection metadata,
  unsupported cell-watch collection metadata,
  unsupported hyperlink collection and item metadata,
  oversized comment text and author payloads,
  invalid legacy protection and write-reservation hashes,
  modern workbook, workbook-revision, and worksheet protection hashes,
  invalid Open XML date cells and formula cached dates,
  gradient fills,
  style-extension payloads,
  oversized cell and cached formula text,
  oversized custom number formats,
  oversized hyperlink payloads and tooltips,
  oversized header/footer text,
  header/footer images,
  custom workbook views,
  custom sheet views,
  named sheet views,
  worksheet sort maps,
  worksheet custom properties,
  oversized worksheet printer settings payloads,
  multiple worksheet printer settings parts,
  unsupported rich-text cell run and font metadata beyond BIFF-backed family,
  character-set, underline styles, superscript/subscript, outline, shadow, condense, and extend metadata,
  unsupported cell font metadata beyond BIFF-backed family, character-set,
  underline styles, superscript/subscript, outline, shadow, condense, and extend metadata,
  duplicated workbook singleton metadata elements,
  duplicated worksheet singleton metadata, layout, print, and feature collection elements,
  duplicated font style properties,
  duplicated fill and border style child elements,
  duplicated cell-format alignment and protection child elements,
  manual page breaks outside BIFF8 worksheet limits,
  VBA projects and macro sheet parts,
  unsupported chart/dialog/macro sheet types,
  PivotTable cache markers,
  unsupported external workbook link metadata,
  external workbook formula references outside the supported sheet-reference subset,
  external defined-name references outside the supported workbook- and sheet-scoped operand subset,
  workbook connections and worksheet query tables,
  embedded OLE objects and packages,
  form controls,
  worksheet slicers and timelines,
  digital signature package metadata,
  direct worksheet image and 3D model relationships,
  worksheet drawings/images/charts, and
  unsupported worksheet phonetic settings,
  phonetic cell text,
  workbook and worksheet data-part relationships,
  workbook metadata extensions,
  rich data features,
  rich styles,
  volatile dependency metadata,
  revision or user data,
  attached toolbars, and
  sparklines,
  non-comment VML drawing shapes, and
  custom workbook table-style metadata.
- Load-policy tests prove the public failure choices: normal
  `ExcelDocument.Load(...)` fails only hard XLS import errors such as unsupported
  password-to-open encryption, non-fatal unsupported BIFF features flow into
  `InspectFeatures()`, and advanced callers can reject preserve-only imports
  with `EnsureNoImportErrors()`, `EnsureNoUnsupportedFeatures()`, and
  `InspectFeatures().EnsureNoAdvancedFeatures()`.
- Unsupported sheet tests and corpus/report coverage expose skipped sheet
  substreams by kind, name, visibility, diagnostics, projection-gap state, and
  chart-sheet metadata such as `PrintSize`, text-object counts, chart record
  categories, and chart type families.
- Compound-feature and object-shape tests prove macro and embedded-content
  boundaries are explicit instead of silent: VBA project storage is preserve-only
  with module/name/size summaries, embedded OLE object storage is preserve-only,
  digital signature streams are diagnosed, and BIFF drawing/form-control object
  records such as pictures, buttons, checkboxes, and dropdown lists are reported
  through drawing-object diagnostics and metadata buckets.

## Remaining Work Checklist

This is the active end-to-end queue. Completed implementation detail belongs in
the support matrix and evidence summary above, not in this list.

The implementation target is now complete for the current release claim: BIFF8
normal workbook reads with no silent loss, plus native BIFF8 write for the
supported OfficeIMO Excel feature surface. The remaining work is release
closeout.

- [ ] Let the current PR #2002 GitHub Actions run finish and inspect only jobs
  that fail. At the 2026-06-28 checkpoint, these jobs were still pending:
  `Analyze (csharp)`, `Cross-platform build (windows-latest)`, `Windows`,
  `Ubuntu net8.0`, and `Ubuntu net10.0`.
- [ ] Fix any new branch-caused or XLS-contract regression found in those logs.
  Do not expand the XLS feature scope unless a failure proves the current read
  or native-write contract is wrong.
- [ ] If code changes again, rerun focused local validation before pushing:
  `dotnet build OfficeIMO.Excel\OfficeIMO.Excel.csproj -f netstandard2.0
  --no-restore -v minimal /clp:ErrorsOnly`, the `LegacyXls` test lane, and the
  rich-text/comment smoke filter.
- [ ] If the support claim changes while fixing a real issue, update this
  roadmap, `OfficeIMO.Excel/README.md`, `OfficeIMO.Excel/COMPATIBILITY.md`, and
  the PR body in the same pass.
- [ ] Recheck PR #2002 after the final CI result: checks, full paginated review
  threads, raw review comments, reviews, PR comments, and PR-body reactions.
- [ ] Resolve any new addressed or outdated-but-fixed inline review threads.
- [ ] Wait for the delayed reviewer settlement on the current head. Treat a
  Codex/IX/Copilot `EYES` reaction as pending until it disappears or produces
  review output.
- [ ] Declare the roadmap closed only when CI is green or every remaining red
  job is explicitly documented as non-XLS release-state context, no actionable
  review feedback remains, and the PR text still matches the final support
  shape.
- [ ] Merge PR #2002 when the gates above are satisfied.
- [ ] After merge or explicit closure, remove the Codex-created worktree/branch
  and any temporary validation artifacts.

## Confidence Statement

The confidence bar for this branch is no silent data loss. A feature may be
fully projected, cached with diagnostics, preserve-only/diagnosed, or blocked
before write, but it must not disappear without a report or preflight failure.

## Guardrails

- Keep parser, legacy model, and Open XML projection separated.
- Keep normal OfficeIMO APIs as the integration path.
- Do not add Excel/COM conversion to production code.
- Do not add external spreadsheet dependencies.
- Do not disguise Open XML `.xlsx` bytes with a legacy `.xls` extension.
