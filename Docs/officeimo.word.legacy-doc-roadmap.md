# OfficeIMO.Word Legacy DOC Read And Native Write Roadmap

This roadmap tracks how to add dependency-free legacy `.doc` support to
`OfficeIMO.Word` while following the same confidence model used by the legacy
`.xls` work.

## Goal

Add first-party support for classic binary Word documents without introducing a
new runtime dependency and without making callers learn a separate conversion
API for normal use.

The target user experience is intentionally quiet:

- `WordDocument.Load("file.doc")` detects a legacy binary Word document and
  projects supported content into a normal `WordDocument`.
- `document.Save("file.docx")` saves the projected document through the normal
  Open XML path.
- `document.Save("file.doc")` eventually writes a native binary `.doc` file
  when the document is within the supported native-write subset.
- explicit `LoadLegacyDoc(...)` and `LoadLegacyDocWithReport(...)` methods exist
  for callers that want diagnostics or format-aware control.
- unsupported, preserve-only, encrypted, or unsafe legacy content is diagnosed
  or blocked before silent data loss.

## Settled Decisions

- [x] Keep this dependency-free at runtime. No NPOI, Word automation, LibreOffice,
  or other conversion engine in `OfficeIMO.Word`.
- [x] Use `OfficeIMO.Shared` as linked source, not a new package. The shared
  source files are already included by libraries such as Word and Excel, and the
  compound-file owner should follow that pattern.
- [x] Use Word COM and NPOI freely for fixture generation, comparison, and local
  validation when useful. Generated `.doc` fixtures and compact reports should be
  checked into the corpus; the generators themselves must not become production
  dependencies.
- [x] Avoid temporary compatibility paths, old/new API probes, or local
  workarounds. If the correct shared owner is not ready, fix the owner first.
- [x] Treat the roadmap as a working load/save ledger. It should show the next
  actionable proof step, not a historical pile of completed experiments.
- [x] PR #2002 is merged into `master`; start implementation from current
  `origin/master` so DOC can build on the merged XLS pattern directly.

## Pattern To Mirror From XLS

The legacy `.xls` implementation provides the pattern to reuse:

- keep the normal public API simple and route by file signature/extension.
- keep legacy code under a clear format-specific folder instead of mixing parser
  details into the main object model.
- parse into a legacy model first, then project into the normal OfficeIMO model.
- expose import diagnostics and unsupported feature lists on the loaded document.
- provide a report-returning load method for callers that need the full import
  story.
- make native legacy saves opt-in through path/stream format routing, with
  preflight blockers for unsupported content.
- use corpus fixtures and checked-in Markdown reports to keep claims honest.

## Ownership Decision

DOC and XLS both need OLE compound file support. That should not become two
independent format-specific implementations.

Current repo state already has:

- `OfficeIMO.Shared\OfficeEncryption.CompoundFile.cs`, a private compound-file
  helper used by encryption.
- `OfficeIMO.Excel\LegacyXls\Compound`, an Excel-owned compound reader/locator
  for legacy workbook streams.

Before adding DOC parsing, promote the reusable compound-file capability into a
shared internal owner under `OfficeIMO.Shared`, then consume it from both
`LegacyXls` and `LegacyDoc`. This should remain linked source, not an additional
NuGet package or project that every OfficeIMO library must reference.

The reusable layer owns only container mechanics: signatures, directories,
FAT/mini-FAT streams, stream lookup, property-set streams, and writer basics.
Word-specific `WordDocument` stream parsing belongs in `OfficeIMO.Word\LegacyDoc`.

## Proposed Project Shape

Add focused files instead of growing the main Word document partials:

```text
OfficeIMO.Shared/
  Compound/
    OfficeCompoundFile.cs
    OfficeCompoundFileEntry.cs
    OfficeCompoundFileReader.cs
    OfficeCompoundFileWriter.cs
    OfficeOlePropertySetReader.cs
    OfficeOlePropertySetWriter.cs

OfficeIMO.Word/
  WordDocument.LoadRouting.cs
  WordDocument.LegacyDoc.cs
  WordDocument.LegacyDocState.cs
  LegacyDoc/
    LegacyDocImportOptions.cs
    LegacyDocLoadResult.cs
    LegacyDocImportReport.cs
    Diagnostics/
    Fib/
    Model/
    Projection/
    Write/
```

PR #2002 has merged, so implementation should stay rebased on current
`origin/master` and use the merged `LegacyXls` source as the reference point. Do
not copy its compound code into Word permanently; extract the reusable owner
first, then consume that owner from both `LegacyXls` and `LegacyDoc`.

## Read Support Scope

Start with Word 97-2003 binary `.doc` files that use the standard OLE compound
container and a `WordDocument` stream. Treat older Word formats as explicit
diagnostic blockers until fixtures prove they are worth supporting.

The first readable slice should project:

- body paragraphs and runs
- plain text and common character formatting
- paragraph alignment, spacing, indentation, and basic styles when reliable
- sections and page setup where the binary structures are straightforward
- tables in the common Word 97-2003 shape
- headers, footers, footnotes, and endnotes after body text is stable
- core/application/custom properties through the shared OLE property-set reader

The first diagnostic slice should detect and report:

- encrypted/password-to-open documents
- macros/VBA project storage
- embedded OLE objects and packages
- ActiveX controls
- tracked changes and comments when not projected
- images, drawings, text boxes, charts, and equations until each has a supported
  projection story
- fast-save or damaged stream shapes that cannot be safely imported

The first implementation should not claim "all `.doc` files". The honest claim
should be closer to: supported `.doc` files can be loaded through normal
`WordDocument.Load(...)`, projected into the normal OfficeIMO Word model, and
saved as `.docx`, with diagnostics for unsupported legacy content.

## Native Write Scope

Native `.doc` writing should come after read/report support is useful. The first
write target should be deliberately smaller than the read target:

- normal document stream and table stream for simple documents
- body paragraphs, runs, and common formatting
- basic page setup and section breaks
- simple tables after paragraph/run output is stable
- core/application/custom properties

Native write should block before saving when the current `WordDocument` contains
features outside the supported binary writer subset. Use the existing
`InspectFeatures()` model as the first preflight source, then add DOC-specific
preflight details for binary-only limits.

## Public API Shape

Keep the normal API invisible for common use:

```csharp
using WordDocument document = WordDocument.Load("input.doc");
document.Save("output.docx");
```

Add explicit diagnostics when callers want control:

```csharp
LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport("input.doc");
if (result.Document != null) {
    result.Document.Save("output.docx");
}
```

Use a save option rather than a separate writer method for stream saves, matching
the XLS direction:

```csharp
document.Save(stream, new WordSaveOptions {
    StreamFormat = WordStreamSaveFormat.LegacyDoc
});
```

The implementation can introduce `WordSaveOptions` and `WordStreamSaveFormat`
only when native `.doc` stream writing is ready. Do not add public options early
just to reserve names.

## Implementation Slices

- [x] Create a dedicated DOC worktree and roadmap branch.
- [x] Confirm XLS PR #2002 is merged and refresh the DOC branch onto current
  `origin/master`.
- [x] Inventory merged `LegacyXls` compound/read/write code and current
  `OfficeIMO.Shared` encryption compound helper; write down the exact extraction
  boundaries before moving code.
- [x] Shared compound-file extraction first slice: promote the XLS-owned OLE
  compound read/write helpers into linked `OfficeIMO.Shared` source and keep XLS
  behavior unchanged.
- [x] Validate the XLS shared-compound slice with `OfficeIMO.Excel` build,
  `OfficeIMO.Tests` build, and the full focused `LegacyXls` test sweep.
- [x] Promote the OLE property-set reader into the shared owner and consume it
  from XLS and DOC import paths.
- [x] Fold the reusable OLE property-set writer into the shared owner when
  native DOC write needs that surface; encryption's private compound helper
  remains separate until encryption needs the shared container surface.
- [x] Add `WordDocument` load routing for `.doc` signature/extension detection
  before `WordprocessingDocument.Open(...)`, with `.docx` and encrypted OOXML
  staying on the existing paths.
- [x] Build a minimal `LegacyDocDocument` model from the FIB and document/table
  stream boundaries, with diagnostics for unsupported or unsafe streams.
  - [x] Reject pre-Word97 FIB versions as explicit import errors before
    projection, while continuing to allow Word 97+ binary FIB streams.
- [x] Project plain body text into normal `WordDocument` paragraphs/runs and
  prove save-to-`.docx` reload through normal OfficeIMO APIs.
- [x] Validate the first DOC reader against real Word COM/NPOI-generated corpus
  fixtures, then keep expanding only where the fixture proves the contract.
- [x] Project OLE SummaryInformation and DocumentSummaryInformation metadata
  into normal Word built-in, application, and scalar custom properties.
- [ ] Add formatting projection for run/paragraph styles only when each mapping
  has a fixture and observable OfficeIMO contract.
  - [x] Project direct CHPX bold/italic character runs into OfficeIMO runs.
  - [x] Project direct CHPX underline, size, and color runs into OfficeIMO runs.
  - [x] Extend character projection to font family through the DOC font table.
  - [x] Project direct PAPX paragraph alignment into OfficeIMO paragraphs.
  - [x] Project direct PAPX paragraph spacing and indentation into OfficeIMO paragraphs.
  - [x] Project fixed built-in paragraph styles from PAPX base style indexes and
    `sprmPIstd` into OfficeIMO paragraph styles.
  - [x] Parse STSH paragraph style records and project stylesheet-defined
    custom paragraph styles as deterministic DOCX paragraph style IDs with
    local style definitions.
  - [x] Validate stylesheet-defined custom paragraph style projection against a
    Word COM-generated legacy `.doc` in the opt-in desktop validation lane.
  - [x] Project stylesheet base-style inheritance plus common style-level
    paragraph/run formatting into custom DOCX style definitions, including
    alignment, spacing, indentation, bold, size, and color, with both synthetic
    and Word COM-generated legacy `.doc` proof.
  - [x] Resolve stylesheet-level font-family run formatting through the DOC font
    table, including styles that carry character formatting after an empty
    paragraph UPX, with synthetic fixture proof and opt-in Word COM validation.
  - [x] Project direct CHPX strikethrough runs plus stylesheet-level
    strikethrough run formatting into normal OfficeIMO runs and custom DOCX
    style definitions with synthetic legacy DOC fixture proof.
  - [x] Project direct CHPX outline runs plus stylesheet-level outline run
    formatting into normal OfficeIMO runs, custom DOCX style definitions, and
    built-in DOCX style definitions with synthetic legacy DOC fixture proof.
  - [x] Project direct CHPX shadow runs plus stylesheet-level shadow run
    formatting into normal OfficeIMO runs, custom DOCX style definitions, and
    built-in DOCX style definitions with synthetic legacy DOC fixture proof.
  - [x] Project direct CHPX emboss runs plus stylesheet-level emboss run
    formatting into normal OfficeIMO runs, custom DOCX style definitions, and
    built-in DOCX style definitions with synthetic legacy DOC fixture proof.
  - [x] Project direct CHPX imprint runs plus stylesheet-level imprint run
    formatting into normal OfficeIMO runs, custom DOCX style definitions, and
    built-in DOCX style definitions with synthetic legacy DOC fixture proof.
  - [x] Project direct CHPX hidden-text runs plus stylesheet-level hidden-text
    run formatting into normal OfficeIMO runs, custom DOCX style definitions,
    and built-in DOCX style definitions with synthetic legacy DOC fixture proof.
  - [x] Project direct CHPX double-strikethrough runs plus stylesheet-level
    double-strikethrough run formatting into normal OfficeIMO runs and custom
    DOCX style definitions with synthetic legacy DOC fixture proof.
  - [x] Project direct CHPX all-caps and small-caps runs plus stylesheet-level
    caps/small-caps run formatting into normal OfficeIMO runs and custom DOCX
    style definitions with synthetic legacy DOC fixture proof.
  - [x] Project direct CHPX superscript/subscript runs plus stylesheet-level
    superscript/subscript run formatting into normal OfficeIMO runs and custom
    DOCX style definitions with synthetic legacy DOC fixture proof.
  - [x] Project direct CHPX highlight runs plus stylesheet-level highlight run
    formatting into normal OfficeIMO runs and custom DOCX style definitions
    with synthetic legacy DOC fixture proof.
  - [x] Project direct PAPX paragraph pagination flags plus stylesheet-level
    keep-lines, keep-next, page-break-before, and widow-control formatting into
    normal OfficeIMO paragraphs and custom DOCX style definitions with
    synthetic legacy DOC fixture proof.
  - [x] Project simple palette-backed direct PAPX paragraph shading from DOC
    `sprmPShd80` into normal OfficeIMO paragraphs with synthetic legacy DOC
    fixture proof.
  - [x] Project simple palette-backed stylesheet paragraph shading from DOC
    `sprmPShd80` into custom DOCX style definitions with synthetic legacy DOC
    fixture proof.
  - [x] Project direct PAPX paragraph tab-stop changes, including clear stops,
    into normal OfficeIMO paragraph `TabStops` with synthetic legacy DOC fixture
    proof.
  - [x] Project stylesheet-level paragraph tab-stop changes, including clear
    stops, into custom and built-in DOCX style definitions with synthetic
    legacy DOC fixture proof.
  - [x] Merge supported DOC stylesheet paragraph/run formatting into built-in
    DOCX style definitions, starting with a Heading 1 fixture that proves
    alignment, spacing, bold, underline, highlight, color, and size projection.
  - [x] Preserve DOC stylesheet inheritance across built-in and custom styles,
    with a custom style inheriting from a formatted Heading 1 fixture.
  - [ ] Expand stylesheet projection to style inheritance and style-level
    paragraph/run formatting beyond the first supported mapping set once each
    additional mapping has a fixture and observable OfficeIMO contract.
- [ ] Add common table projection after paragraph/run projection is stable.
  - [x] Project simple DOC cell/row marker tables into `WordTable` instances
    with plain cell text.
  - [x] Preserve direct run formatting inside projected simple table cells.
  - [x] Prefer explicit PAPX table cell and end-row markers when present,
    including trailing empty cells, while retaining the simple marker heuristic
    for existing minimal fixtures.
  - [x] Project simple row-level table cell widths from DOC `sprmTDefTable`
    row definitions into normal OfficeIMO table cell width properties.
  - [x] Report merged table cell descriptors from DOC `sprmTDefTable`/`TC80`
    row definitions before merged-cell projection exists, so native DOC re-save
    is blocked instead of flattening merged table structure.
  - [x] Project simple horizontal merged cells from DOC `sprmTDefTable`/`TC80`
    row definitions into normal OfficeIMO horizontal merge properties.
  - [x] Project simple vertical merged cells from DOC `sprmTDefTable`/`TC80`
    row definitions into normal OfficeIMO vertical merge properties, while
    keeping invalid/conflicting merge descriptors diagnosed instead of silently
    flattening table structure.
  - [x] Project simple row-level table heights from DOC `sprmTDyaRowHeight`
    row definitions into normal OfficeIMO table row height properties,
    preserving exact versus at-least height rules in the Open XML row.
  - [x] Project simple row-level repeat-header and no-split flags from DOC
    `sprmTTableHeader` and `sprmTFCantSplit*` row definitions into normal
    OfficeIMO table row properties.
  - [x] Project simple table alignment from DOC `sprmTJc` row definitions into
    normal OfficeIMO table alignment, and preserve it through native DOC
    save/reload.
  - [x] Project simple table indentation from the DOC `sprmTDefTable` first
    row edge into normal OfficeIMO table indentation, and preserve it through
    native DOC save/reload.
  - [x] Project simple table preferred width and autofit layout from DOC
    `sprmTTableWidth` and `sprmTFAutofit` row definitions into normal
    OfficeIMO table width/layout properties, and preserve them through native
    DOC save/reload.
  - [x] Project simple table cell vertical alignment from DOC `sprmTDefTable`
    / `TC80` row definitions into normal OfficeIMO table cell vertical
    alignment properties, and preserve it through native DOC save/reload.
  - [x] Project simple table cell fit-text and no-wrap flags from DOC
    `sprmTDefTable` / `TC80` row definitions into normal OfficeIMO table cell
    text layout properties, and preserve them through native DOC save/reload.
  - [x] Project simple table cell hide-mark flags from DOC `sprmTDefTable` /
    `TC80` row definitions into normal OfficeIMO table cell properties, and
    preserve them through native DOC save/reload.
  - [x] Project simple table cell text direction from DOC `sprmTDefTable`
    / `TC80` `textFlow` values into normal OfficeIMO table cell text
    direction properties, and preserve them through native DOC save/reload.
  - [x] Project simple table cell margins from DOC `sprmTCellPadding` and
    `sprmTCellPaddingDefault` row definitions into normal OfficeIMO table cell
    margin properties, and preserve them through native DOC save/reload.
  - [x] Project simple table-level cell spacing from DOC
    `sprmTCellSpacingDefault` row definitions into normal OfficeIMO table cell
    spacing, and preserve it through native DOC save/reload.
  - [x] Project simple palette-backed table cell shading from DOC
    `sprmTDefTableShd80`/`Shd80` row definitions into normal OfficeIMO table
    cell fill colors, and preserve it through native DOC save/reload.
  - [x] Project simple palette-backed table cell borders from DOC `TC80`
    `Brc80` values in `sprmTDefTable` row definitions into normal OfficeIMO
    table cell border properties, and preserve them through native DOC
    save/reload.
  - [x] Preserve simple palette-backed table-level borders by expanding normal
    `tblBorders` outer and inside edges into DOC `TC80` `Brc80` cell borders
    during native DOC save/reload.
  - [ ] Add table formatting, merged cells, and nested tables as separate
    fixture-backed slices.
- [ ] Add section/page setup, headers, footers, footnotes, and endnotes as
  separate fixture-backed slices.
  - [x] Project single-section page size, orientation, margins, header/footer
    distance, and gutter from DOC `PlcfSed`/`Sepx` records into normal
    OfficeIMO section properties.
  - [x] Report multiple section descriptor records as unsupported/preserve-only
    before multi-section projection exists, so native DOC re-save is blocked
    instead of flattening section boundaries.
  - [x] Project paragraph-boundary multi-section breaks with per-section page
    setup from DOC `PlcfSed`/`Sepx` records, and preserve that simple shape
    through native DOC save/reload.
  - [x] Project paragraph-boundary section break kinds from DOC `sprmSBkc`
    records and preserve continuous section breaks through native DOC
    save/reload.
  - [x] Preserve paragraph-boundary section breaks after simple table body
    blocks through native DOC save/reload, including per-section page setup.
  - [ ] Add section breaks inside richer body shapes, headers, footers,
    footnotes, and endnotes as separate fixture-backed slices.
- [x] Wire unsupported/preserve-only DOC features into `LegacyDocImportReport`
  and loaded `WordDocument` state.
  - [x] Report unsupported header/footer, footnote, endnote, comment, and text
    box stories from FIB story counts before those stories are projected.
  - [x] Keep body text projection clipped to the main `Fib:CcpText` story so
    diagnosed header/footer story text is not flattened into body paragraphs.
  - [x] Report container-level VBA project storage before macros have a
    projection or preservation story.
  - [x] Report container-level ActiveX controls and embedded package payloads
    before those features have a projection story.
  - [x] Report non-empty compound `Data` streams as binary payloads before
    pictures, drawings, form fields, or related payloads are projected.
  - [x] Report FIB fast-save/quick-save and picture-present flags before those
    preserve-only states have a projection or rewrite story.
  - [x] Report DOC revision-tracking DOP flags before tracked revisions have a
    projection or rewrite story.
- [x] Add normal fixture folders under
  `OfficeIMO.Tests\Documents\LegacyDocCorpus` and
  keep `OfficeIMO.Tests\Documents\LegacyDocDiagnosticCorpus` for the first
  diagnostic corpus slice.
- [x] Add a Word COM fixture-generation helper for simple paragraph, character
  formatting, and paragraph formatting fixtures in test/support tooling only;
  checked-in fixtures are the source of CI proof.
- [ ] Add optional NPOI fixture-generation tooling only if a feature slice needs
  deterministic `.doc` shapes that are awkward to produce through Word COM.
- [x] Add an opt-in desktop Word COM validation lane that checks generated DOC
  import, native DOC save, and checked-in corpus openability without running in
  normal CI.
- [x] Add corpus report approval tests with short Markdown baselines.
- [x] Define the first native `.doc` writer preflight for paragraph-only output,
  including body element, document part, paragraph, and run blockers before any
  target file bytes are committed.
  - [x] Block native `.doc` save for documents imported from legacy DOC when
    the import reported unsupported or preserve-only features, before file or
    stream bytes are committed.
  - [x] Block native `.doc` save for revision tracking settings, tracked
    revision markup, and comments before file bytes are committed.
- [x] Implement native writer first slice for simple documents and prove
  OfficeIMO can reload written `.doc` output through the legacy reader.
- [x] Introduce `WordSaveOptions` and `WordStreamSaveFormat` for explicit native
  `.doc` stream saves once the native writer first slice was ready.
- [ ] Expand native writer slices for formatting, tables, and simple
  sections only after preflight blocks all unsupported content.
  - [x] Write direct bold/italic CHPX runs and reload them through the legacy
    reader.
  - [x] Write direct underline, size, and color CHPX runs and reload them
    through the legacy reader.
  - [x] Write direct strikethrough CHPX runs and reload them through the legacy
    reader.
  - [x] Write direct outline CHPX runs and reload them through the legacy reader.
  - [x] Write direct shadow CHPX runs and reload them through the legacy reader.
  - [x] Write direct emboss CHPX runs and reload them through the legacy reader.
  - [x] Write direct imprint CHPX runs and reload them through the legacy reader.
  - [x] Write direct hidden-text CHPX runs and reload them through the legacy reader.
  - [x] Write direct double-strikethrough CHPX runs and reload them through
    the legacy reader.
  - [x] Write direct all-caps and small-caps CHPX runs and reload them through
    the legacy reader.
  - [x] Write direct superscript/subscript CHPX runs and reload them through
    the legacy reader.
  - [x] Write direct highlight CHPX runs and reload them through the legacy
    reader.
  - [x] Extend native character writing to font family through the DOC font
    table.
  - [x] Write direct paragraph alignment PAPX records and reload them through
    the legacy reader.
  - [x] Write direct paragraph spacing and indentation PAPX records and reload
    them through the legacy reader.
  - [x] Write direct paragraph pagination flags and reload keep-lines,
    keep-next, page-break-before, and widow-control formatting through the
    legacy reader.
  - [x] Write simple palette-backed paragraph shading with `sprmPShd80`, then
    reload it through the legacy reader while blocking non-palette fill colors
    before bytes are committed.
  - [x] Write built-in paragraph style PAPX records and reload them through
    the legacy reader.
  - [x] Write direct paragraph tab-stop PAPX records, including clear stops,
    and reload them through the legacy reader.
  - [x] Project imported tab characters and native-written tabs as real Word
    tab runs after legacy DOC reload.
  - [x] Project imported line/page break characters and native-written
    text-wrapping/page breaks as real Word break runs after legacy DOC reload.
  - [x] Stop stamping quick-save count bits into native-written DOC FIB flags,
    and prove OfficeIMO-authored DOC output reloads without preserve-only FIB
    flag diagnostics.
  - [x] Write simple body tables with one paragraph per cell and supported run
    content, then reload them as `WordTable` instances through the legacy reader.
  - [x] Preserve supported paragraph formatting inside simple native-written
    table cells and project fixture-backed table-cell paragraph formatting
    during legacy DOC reload.
  - [x] Write explicit PAPX table cell and end-row marker flags for native
    simple tables, so saved DOC output carries table structure metadata instead
    of relying only on marker-character fallback.
  - [x] Write simple row-level table cell width definitions with
    `sprmTDefTable` and reload them through the legacy reader.
  - [x] Use explicit `tblGrid` column widths as the native `.doc` table row
    definition fallback when cells do not carry direct `tcW` widths, and block
    grid widths outside the Word 97-2003 signed twip range before bytes are
    committed.
  - [x] Block multi-paragraph table cells before native `.doc` bytes are
    committed until TAP-backed table parsing can disambiguate table-internal
    paragraph marks from normal paragraphs before a table.
  - [x] Block nested table cells before native `.doc` bytes are committed until
    nested table projection has TAP-backed read/write coverage.
  - [x] Write simple table row heights with `sprmTDyaRowHeight`, reload them
    through the legacy reader, and block unsupported row-level table properties
    before native `.doc` bytes are committed.
  - [x] Write simple table row repeat-header and no-split flags with
    `sprmTTableHeader` and `sprmTFCantSplit90`, then reload them through the
    legacy reader.
  - [x] Write simple table alignment with `sprmTJc`, then reload it through
    the legacy reader.
  - [x] Write simple table indentation by carrying `tblInd` into the
    `sprmTDefTable` first row edge, then reload it through the legacy reader.
  - [x] Write simple table preferred width and autofit layout with
    `sprmTTableWidth` and `sprmTFAutofit`, then reload them through the legacy
    reader.
  - [x] Write simple horizontal table cell merges with `sprmTDefTable`/`TC80`,
    including the normal OfficeIMO `gridSpan` save shape, and reload them
    through the legacy reader while blocking vertical merges before bytes are
    committed.
  - [x] Write simple vertical table cell merges with `sprmTDefTable`/`TC80`,
    reload them through the legacy reader, and keep invalid/conflicting imported
    merge descriptors blocked before native `.doc` bytes are committed.
  - [x] Write simple table cell fit-text and no-wrap flags with
    `sprmTDefTable`/`TC80`, then reload them through the legacy reader.
  - [x] Write simple table cell hide-mark flags with `sprmTDefTable`/`TC80`,
    then reload them through the legacy reader.
  - [x] Write simple table cell text direction with `sprmTDefTable`/`TC80`
    `textFlow`, then reload it through the legacy reader.
  - [x] Write simple table cell margins with `sprmTCellPadding`, then reload
    them through the legacy reader.
  - [x] Write table-level default cell margins with `sprmTCellPaddingDefault`,
    then reload inherited defaults and per-cell overrides through the legacy
    reader.
  - [x] Write table-level cell spacing with `sprmTCellSpacingDefault`, then
    reload it through the legacy reader.
  - [x] Write simple palette-backed table cell shading with
    `sprmTDefTableShd80`, then reload it through the legacy reader while
    blocking non-palette fill colors before bytes are committed.
  - [x] Write simple palette-backed table cell borders with `TC80` `Brc80`
    values in `sprmTDefTable`, then reload them through the legacy reader while
    blocking unsupported border styles and non-palette colors before bytes are
    committed.
  - [x] Write simple table-level `tblBorders` by expanding outer and inside
    table edges to per-cell `TC80` `Brc80` values, then reload them through the
    legacy reader while keeping direct cell borders as explicit overrides.
  - [x] Write simple final-section page size, orientation, margins,
    header/footer distance, and gutter, then reload them through the legacy
    reader.
  - [x] Keep blocking unsupported final-section properties before native `.doc`
    bytes are committed so unimplemented section features are not silently
    dropped.
  - [x] Write paragraph-boundary next-page section breaks with simple
    per-section page setup, then reload them through the legacy reader.
  - [x] Write paragraph-boundary section break kinds, including continuous
    breaks, then reload them through the legacy reader.
  - [x] Write paragraph-boundary section breaks after simple table body blocks,
    then reload the table and following section page setup through the legacy
    reader.
  - [ ] Add table formatting, merged/nested tables, section breaks inside richer
    body shapes, and richer section writing as separate preflight-backed slices.
- [x] Update `OfficeIMO.Word\COMPATIBILITY.md` and README wording only after tests
  prove the support statement.
- [ ] Before PR handoff or merge, rerun the focused DOC lane, the shared compound
  lane, full `LegacyXls` sweeps, `OfficeIMO.Word` builds across supported target
  frameworks, and `git diff --check`.

## Roadmap Operating Rules

- Keep exactly one or a small number of slices active at a time.
- Before repeating an investigation, note what changed since the last pass and
  what new evidence the rerun should produce.
- When a slice lands, replace speculative wording with the current tested
  contract and leave the next unchecked item visible.
- Do not carry completed temporary notes forever. Collapse completed evidence
  into the current-state text once it stops helping the next implementation pass.
- If a slice uncovers a shared-owner bug, update the shared-owner item instead of
  adding a Word-local workaround.

## Fixture Generation Policy

Generated fixtures are welcome; runtime dependencies are not.

- Word COM may be used on Windows to generate real `.doc` samples, compare Word's
  interpretation of generated files, or produce difficult layout/feature cases.
- NPOI may be used in test-support tooling to generate targeted `.doc` fixtures
  when that is faster or more deterministic than COM.
- COM/NPOI generation should be repeatable from scripts or test utilities, but
  CI should rely on checked-in fixture files and reports unless a specific
  Windows-only validation lane is intentionally added.
- Do not add NPOI to `OfficeIMO.Word` or to production projects for this feature.
- Do not use generated fixtures to justify broad claims. Every fixture should map
  to one support statement or one explicit unsupported boundary.

## Quality Bar

This feature should be built as if OfficeIMO is aiming to be the best Word
automation library in its class:

- no hidden converter dependency
- no silent loss on legacy import or native save
- no temporary shims waiting for another branch
- no separate public API required for normal `.doc` loading
- fixtures for every advertised capability
- diagnostics for every known boundary
- public docs that say what works, what is blocked, and what remains partial

## Test Plan

Use contract-focused tests rather than a giant incidental matrix:

- routing tests for `.doc`, renamed `.doc`, `.docx`, encrypted OOXML, and invalid
  binary input.
- normal load tests for path, stream, async path, and converted `.docx` save/reload.
- report tests proving unsupported features are visible before conversion.
- corpus report approval tests with a short, reviewable Markdown report per
  fixture.
- native writer tests that write `.doc`, reload through the legacy reader, and
  compare the projected OfficeIMO document contract.
- preflight tests proving unsupported content blocks native `.doc` save before
  bytes are committed.

Run at least:

```powershell
dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj --configuration Release --filter "FullyQualifiedName~LegacyDoc|FullyQualifiedName~Word.Save|FullyQualifiedName~Word.Load"
dotnet build OfficeIMO.Word\OfficeIMO.Word.csproj --configuration Release
```

When the shared compound layer changes, also run the legacy XLS tests because the
container owner is shared:

```powershell
dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj --configuration Release --filter FullyQualifiedName~LegacyXls
```

## Remaining Design Choices

- `WordSaveOptions` and `WordStreamSaveFormat` now exist for native `.doc`
  stream saves. Keep future save options tied to real implemented behavior
  rather than reserving placeholder names.
- Use `AllowLossyLegacyDocSave` only if native writer work reaches a real
  preserve-only import state where an explicit caller override is safer than a
  blanket block.
- Decide whether COM-generated, NPOI-generated, or external corpus fixtures form
  the first baseline per feature slice. The answer can vary by slice as long as
  each fixture is checked in and explained.

## Current Worktree

This plan was started in the dedicated worktree:

```text
C:\Support\GitHub\_worktrees\OfficeIMO-legacy-doc-support
```

Branch:

```text
codex/legacy-doc-support-plan
```
