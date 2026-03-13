# OfficeIMO => Google Workspace Analysis

## Goal

Describe how OfficeIMO could support native Google Workspace generation and synchronization, not just Drive-based file conversion.

This document focuses on:

- `OfficeIMO.Word` => Google Docs
- `OfficeIMO.Excel` => Google Sheets
- package shape, architecture, feature mapping, gaps, and delivery phases

It does not assume this can be built quickly. The point is to show how we would build it if we chose to invest in it.

## Executive Summary

The short version:

- `OfficeIMO.Word` can be mapped to Google Docs, but this is not a "save as Google Doc" feature. It is a document compiler that must walk the OfficeIMO object model and emit ordered `documents.batchUpdate` requests against Google Docs.
- `OfficeIMO.Excel` maps better to Google Sheets than Word maps to Docs, because the Sheets API exposes a large spreadsheet object model: sheets, ranges, merges, conditional formatting, named ranges, filters, charts, pivot tables, protected ranges, developer metadata, tables, and grid properties.
- Word and Excel should not share the same implementation path. They need separate translators and separate feature matrices.
- A thin Drive import/export bridge is useful, but it is not the main design if the real goal is native Google Workspace authoring.
- The cleanest long-term shape is a separate package family, not changes inside the core OpenXML packages.

My inference from the current repo is that this work fits best as extension packages in the same style as `OfficeIMO.Word.Html` and `OfficeIMO.Word.Markdown`, not as changes to the core `OfficeIMO.Word` or `OfficeIMO.Excel` packages.

## Why This Is A Different Class Of Project

OfficeIMO is fundamentally an OpenXML family of libraries today:

- root repo positioning: `README.md`
- Word package: `OfficeIMO.Word/OfficeIMO.Word.csproj`
- Excel package: `OfficeIMO.Excel/OfficeIMO.Excel.csproj`

The existing extension pattern is:

- `OfficeIMO.Word.Html`
- `OfficeIMO.Word.Markdown`

Those packages convert around the Word model. They do not provide a backend-neutral authoring model. That matters because native Google support is not just "another serializer". It requires mapping to a different editing model and a different set of API rules.

## Current OfficeIMO Surfaces We Can Build On

### Word

OfficeIMO already exposes enough traversal surface to build a translator:

- sections: `OfficeIMO.Word/WordDocument.cs`
- ordered section traversal: `OfficeIMO.Word/Converters/DocumentTraversal.cs`
- paragraphs, tables, lists, images, charts, footnotes, headers, footers, content controls: `OfficeIMO.Word/WordDocument.cs`, `OfficeIMO.Word/WordSection.cs`, `OfficeIMO.Word/WordHeaderFooter.Properties.cs`
- table operations and normalization helpers: `OfficeIMO.Word/WordTable*.cs`
- paragraph/run operations: `OfficeIMO.Word/WordParagraph*.cs`

There is even explicit Google Docs compatibility work already, but only for `.docx` rendering compatibility, not native Docs authoring:

- `OfficeIMO.Word/WordDocument.WebCompat.cs`
- `OfficeIMO.Word/WordTable.Properties.cs`
- `OfficeIMO.Word/WordTable.Methods.cs`
- `OfficeIMO.Tests/Word.WebCompat.cs`

### Excel

OfficeIMO.Excel is even more promising for a native Google target:

- workbook and sheets: `OfficeIMO.Excel/ExcelDocument.cs`
- sheet list: `OfficeIMO.Excel/ExcelDocument.cs`
- values and formulas: `OfficeIMO.Excel/ExcelSheet.CellValue.cs`
- named ranges: `OfficeIMO.Excel/ExcelDocument.NamedRanges.cs`
- tables: `OfficeIMO.Excel/ExcelSheet.Tables.cs`
- charts: `OfficeIMO.Excel/ExcelSheet.Charts.cs`
- validation: `OfficeIMO.Excel/ExcelSheet.DataOperations.cs`
- conditional formatting: `OfficeIMO.Excel/ExcelSheet.ConditionalFormatting.cs`
- typed reads / editable rows: `OfficeIMO.Excel/ExcelSheet.ReadBridge.cs`

## Current Google Workspace Capability Relevant To This Design

### Drive

Drive is still useful, but only as infrastructure:

- upload/create files
- convert uploaded Office files into Google-native files
- export Google-native Docs/Sheets to `.docx` / `.xlsx`
- comments and replies live in Drive, not in Docs API

Useful, but not sufficient for native authoring.

### Google Docs API

The Docs API supports:

- create document
- get full document structure as JSON
- atomic `documents.batchUpdate`
- insert and format text
- create bullets
- insert tables, rows, and columns
- update table row / cell / column properties
- merge and unmerge table cells
- insert inline images
- insert page breaks
- create named ranges and replace named range content
- create headers, footers, and footnotes
- insert section breaks and update section style
- pin header rows in tables

Important constraints:

- edits are index-based inside body/header/footer/footnote segments
- section breaks cannot be inserted inside tables, equations, headers, or footers
- inline images need a public URI and must be inline
- many operations must happen inside an existing paragraph
- comments are handled through the Drive API, and Drive's anchored comments are treated by Google Workspace editors as effectively unanchored comments

### Google Sheets API

The Sheets API supports:

- spreadsheet-level `spreadsheets.batchUpdate`
- `spreadsheets.values` for efficient values writes
- sheet add/update/delete and ordering
- row/column insert/delete/move/autoresize
- cell updates and formula writes
- merges
- borders and formatting
- conditional formatting
- data validation
- basic filters and filter views
- protected ranges
- named ranges
- charts
- pivot tables
- developer metadata
- row/column groups
- grid properties like frozen rows/columns and hidden gridlines
- tables, including table metadata and update/delete requests

Important constraints:

- the API has explicit support for many spreadsheet concepts, so Excel=>Sheets can be high fidelity for core tabular scenarios
- images are visible in the request surface as embedded objects for move/delete, but I did not find a first-class "add image" request in the current request surface; this likely needs either a different insertion mechanism or a deferred implementation
- print/header-footer style features are not prominent in the current request surface; those should be treated as potential gaps unless confirmed separately

## Package Architecture Options

### Option A: Direct Translators

Packages:

- `OfficeIMO.Word.GoogleDocs`
- `OfficeIMO.Excel.GoogleSheets`
- optional shared auth/Drive package: `OfficeIMO.GoogleWorkspace`

How it works:

- walk current OfficeIMO objects directly
- emit Google API requests
- keep state locally while building Google-side objects

Pros:

- fastest route to a working implementation
- leverages current OfficeIMO APIs immediately
- lowest up-front architecture cost

Cons:

- translation logic becomes coupled to OfficeIMO internals
- harder to support round-trip or future targets later
- Word and Excel translators will likely diverge heavily

### Option B: Intermediate Representation

Packages:

- `OfficeIMO.Documents`
- `OfficeIMO.Word.GoogleDocs`
- `OfficeIMO.Excel.GoogleSheets`
- maybe future `OfficeIMO.PowerPoint.GoogleSlides`

How it works:

- map OfficeIMO to an internal canonical document/workbook model
- map that internal model to Google Workspace

Pros:

- better long-term architecture
- easier multi-target support
- easier testing

Cons:

- much more initial work
- likely a new platform inside the repo

### Recommended Direction

Hybrid approach:

1. Build direct translators first.
2. Internally structure them around small mapping models and operation plans.
3. Promote those internal models into a formal intermediate representation only if the feature line proves worth continuing.

This lowers risk while keeping a path to a cleaner future architecture.

## Recommended Package Layout

Recommended first pass:

- `OfficeIMO.GoogleWorkspace`
  - auth
  - credential loading
  - Drive helpers
  - common error model
  - retry / rate limit policy
  - shared HTTP / Google client setup
- `OfficeIMO.Word.GoogleDocs`
  - `WordDocument` => Google Docs translator
  - optional Google Docs => lightweight OfficeIMO import helpers later
- `OfficeIMO.Excel.GoogleSheets`
  - `ExcelDocument` => Google Sheets translator
  - optional Google Sheets => lightweight OfficeIMO import helpers later

Public API shape could look like:

```csharp
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word.GoogleDocs;
using OfficeIMO.Excel.GoogleSheets;

var session = GoogleWorkspaceSession.FromOAuthClient(...);

using var word = WordDocument.Create(...);
var googleDoc = await word.SaveAsGoogleDocAsync(session, new GoogleDocsSaveOptions());

using var excel = ExcelDocument.Create(...);
var googleSheet = await excel.SaveAsGoogleSpreadsheetAsync(session, new GoogleSheetsSaveOptions());
```

## Authentication And Tenancy

This should be explicit in the design from day one.

Supported auth modes:

- OAuth client flow for user-owned files
- service account for app-owned files
- service account + domain-wide delegation for org scenarios

Recommended abstraction:

- `IGoogleWorkspaceCredentialSource`
- `GoogleWorkspaceSession`
- token cache abstraction
- file ownership / drive location options

Why this matters:

- native Docs/Sheets creation is multi-tenant and identity-sensitive
- service account behavior differs from user flow
- comments, permissions, and sharing are Drive-level concerns

## Word => Google Docs Design

### Compiler Model

The Word translator should work like a compiler, not a serializer.

Recommended pipeline:

1. Traverse `WordDocument` in document order.
2. Build a stream of logical operations:
   - start section
   - paragraph
   - list item
   - table
   - inline image
   - footnote reference
   - header/footer content
3. Convert logical operations into ordered Docs API requests.
4. Track document indexes as requests are planned and applied.
5. Post-process references that require IDs returned by earlier requests:
   - header/footer IDs
   - named range IDs
   - maybe comment anchors or bookmarks

This should be implemented with a `GoogleDocsOperationPlan`, not ad hoc request generation.

### Word Feature Mapping

| OfficeIMO.Word feature | Google Docs target | Strategy |
| --- | --- | --- |
| Plain paragraphs | Paragraph + text runs | Direct mapping |
| Heading styles | Paragraph style / named heading | Direct mapping |
| Bold/italic/underline/strike | Text style | Direct mapping |
| Text color / highlight | Text style | Direct mapping |
| Paragraph alignment / indentation / spacing | Paragraph style | Direct mapping, but expect some differences |
| Bulleted / numbered lists | Create paragraph bullets | Direct mapping, careful with nesting and restart behavior |
| Hyperlinks | Text style link | Direct mapping |
| Tables | Insert table + cell text + style updates | Direct mapping |
| Merge cells | MergeTableCellsRequest | Direct mapping |
| Repeat header row | PinTableHeaderRowsRequest | Direct mapping |
| Images | InsertInlineImageRequest | Direct mapping for inline images only |
| Page breaks | InsertPageBreakRequest | Direct mapping |
| Sections | InsertSectionBreakRequest + UpdateSectionStyleRequest | Direct mapping, but sensitive to Docs index rules |
| Headers / footers | CreateHeaderRequest / CreateFooterRequest | Direct mapping |
| Footnotes | CreateFootnoteRequest | Direct mapping |
| Bookmarks / anchors | Named ranges | Partial semantic mapping |
| Comments | Drive comments API | Out-of-band mapping |
| TOC field | Docs table of contents structure | Likely indirect / limited |
| Equations | Docs equation structure exists in document model | Treat as unsupported until a safe write strategy is verified |
| Content controls (checkbox/date/dropdown/etc.) | No real equivalent | Flatten to text or checklist style |
| SmartArt | No real equivalent | Render as image or flatten to paragraphs |
| Word charts | No direct Word-chart equivalent in Docs API | Render as image, or generate chart in Sheets and embed later if we choose to support that path |
| Text boxes / shapes / floating layout | Weak fit | Flatten or rasterize |
| Watermarks | Weak fit | Ignore, flatten to header text, or rasterize |
| Embedded objects / OLE | No fit | Ignore or attach externally |
| Track changes / revisions | No parity | Ignore in v1 |
| Macros / protection | No parity | Ignore in v1 |

### Word Hard Problems

#### 1. Index-sensitive editing

Docs writes are index-based. This means request order is part of the implementation, not a convenience.

We will need:

- an index allocator / planner
- segment-aware insertion helpers for body/header/footer/footnote
- request coalescing so we do not spam tiny API calls

#### 2. Sections are not Word sections

Docs has sections, but Word's section semantics are richer in practice. We should treat section mapping as:

- page size / margins / columns where possible
- headers / footers where possible
- expect edge-case mismatches

#### 3. Floating content

Word can do far more with floating images, shapes, text boxes, and positioning than the Docs request model comfortably exposes in a deterministic way.

Recommendation:

- v1 supports inline content only
- floating objects are flattened or rasterized

#### 4. Comments and anchors

Docs comments are not part of the Docs API write surface. Drive comments exist, but Drive says anchored comments are treated by Google Workspace editors as unanchored comments.

So:

- comment text can be preserved
- exact anchor fidelity should not be promised

### Recommended Word Delivery Scope

#### V1

- paragraphs
- runs / styling
- headings
- hyperlinks
- lists
- tables including merge and pinned header rows
- inline images
- page breaks
- sections
- headers / footers
- footnotes
- named ranges / bookmarks

#### V2

- richer paragraph and table formatting
- comments via Drive API
- better section/page setup fidelity
- optional chart-as-image fallback

#### V3

- partial round-trip support
- imported Docs => OfficeIMO document model
- richer document diagnostics / fidelity report

## Excel => Google Sheets Design

### Compiler Model

Sheets should use a split pipeline:

1. Create spreadsheet.
2. Create sheets and base sheet properties.
3. Write cell values and formulas efficiently.
4. Apply formatting and structural operations.
5. Add higher-level objects:
   - named ranges
   - filters
   - tables
   - charts
   - pivots
   - conditional formatting
   - protected ranges
   - metadata

Unlike Docs, Sheets has a strong grid object model. This is much closer to what OfficeIMO.Excel already represents.

### Excel Feature Mapping

| OfficeIMO.Excel feature | Google Sheets target | Strategy |
| --- | --- | --- |
| Workbook | Spreadsheet | Direct mapping |
| Worksheet | Sheet | Direct mapping |
| Sheet order / visibility / RTL | SheetProperties | Direct mapping |
| Cell values | values/updateCells | Direct mapping |
| Formulas | updateCells / userEnteredValue.formulaValue | Direct mapping with formula normalization rules |
| Number/date formats | CellFormat.numberFormat | Direct mapping |
| Fonts / fills / borders / alignment | CellFormat | Direct mapping |
| Column width / row height | UpdateDimensionProperties | Direct mapping |
| Freeze panes | GridProperties frozen rows/cols | Direct mapping |
| Hide gridlines | GridProperties.hideGridlines | Direct mapping |
| Merged cells | MergeCellsRequest | Direct mapping |
| Named ranges | AddNamedRangeRequest | Direct mapping |
| Data validation | SetDataValidationRequest | Direct mapping |
| Basic filters | SetBasicFilterRequest | Direct mapping |
| Filter views | FilterView requests | Direct mapping if desired |
| Conditional formatting | AddConditionalFormatRuleRequest | Direct mapping |
| Protected ranges | AddProtectedRangeRequest | Direct mapping |
| Charts | AddChartRequest / UpdateChartSpecRequest | Direct mapping for supported chart classes |
| Pivot tables | UpdateCells with PivotTable definition | Direct mapping |
| Row/column groups | Dimension group requests | Direct mapping |
| Tables | AddTableRequest / UpdateTableRequest / DeleteTableRequest | Direct mapping |
| Developer metadata | CreateDeveloperMetadataRequest | Direct mapping |
| Header/footer print text and images | No obvious first-class match in request surface reviewed | Treat as gap until proven |
| Print area / print titles | No obvious first-class match in request surface reviewed | Treat as gap until proven |
| In-sheet images | Unclear first-class create path in reviewed request surface | Treat as gap / deferred |

### Excel Hard Problems

#### 1. Formula semantics

Google Sheets formulas are not Excel formulas in every case.

We will need:

- function name compatibility map
- unsupported function detection
- optional "preserve formula as text" mode
- fidelity report for changed formulas

#### 2. Table semantics

This area now looks better than it used to because the Sheets API exposes tables directly.

Still needed:

- name normalization
- header mapping
- dropdown column types
- table range maintenance

#### 3. Chart spec translation

OfficeIMO has a broad chart surface. Google Sheets charts are powerful, but not identical.

We should define:

- supported chart types
- partial support chart types
- unsupported style properties
- chart fallback behavior

#### 4. Images and print settings

These are the least clear pieces in the currently reviewed Sheets request surface.

Recommendation:

- do not promise them in v1
- implement explicit diagnostics saying they were skipped

### Recommended Excel Delivery Scope

#### V1

- workbook and sheets
- cell values and formulas
- core formatting
- merges
- named ranges
- filters
- conditional formatting
- validation
- protected ranges
- freeze panes / grid properties
- tables

#### V2

- charts
- pivot tables
- developer metadata
- row/column grouping
- better formula translation diagnostics

#### V3

- import path from Google Sheets back to OfficeIMO
- richer chart parity
- remaining print/image gaps if API support is confirmed

## Diagnostics And Fidelity Reporting

This project should include a first-class fidelity report from day one.

Suggested types:

```csharp
public sealed class TranslationReport {
    public List<TranslationNotice> Notices { get; } = new();
    public bool HasLossyMappings => Notices.Any(n => n.Severity >= TranslationSeverity.Warning);
}

public sealed class TranslationNotice {
    public string Path { get; init; } = "";
    public string Feature { get; init; } = "";
    public TranslationSeverity Severity { get; init; }
    public string Message { get; init; } = "";
}
```

Why this matters:

- Word and Excel parity will never be perfect
- users need to know what was flattened, skipped, rasterized, or downgraded

## Testing Strategy

### Word

- unit tests for operation plan generation
- golden tests on batch request payloads
- integration tests against disposable test documents
- round-trip smoke tests for:
  - OfficeIMO Word => Google Doc => export `.docx`
  - compare text, headings, lists, tables, images count, headers/footers, footnotes

### Excel

- unit tests for range, style, and request planning
- integration tests against disposable spreadsheets
- OfficeIMO Excel => Google Sheet => export `.xlsx`
- compare:
  - sheet names
  - dimensions
  - values
  - formulas
  - named ranges
  - merges
  - tables
  - conditional formatting count

## Suggested Build Plan

### Phase 0: Foundations

- add `OfficeIMO.GoogleWorkspace`
- auth/session abstractions
- test project infrastructure
- Google client bootstrapping

### Phase 1: Excel First

Do Excel first if the goal is earliest native success.

Why:

- Sheets API is richer and more explicit than Docs API
- OfficeIMO.Excel concepts align better with Sheets concepts
- more direct mapping, fewer layout traps

Deliver:

- create spreadsheet
- create sheets
- values, formulas, formats
- merges
- named ranges
- validation
- conditional formatting
- filters
- tables

### Phase 2: Word Core

Deliver:

- create doc
- paragraphs and runs
- headings
- lists
- hyperlinks
- tables
- inline images
- page breaks
- headers/footers
- footnotes

### Phase 3: Advanced Features

- Excel charts and pivots
- Word sections and richer table styling
- comments, diagnostics, optional fallback rendering

## Final Recommendation

If we decide to build native Google Workspace support, the best engineering path is:

1. Create separate Google Workspace extension packages.
2. Build Excel=>Sheets first for faster native parity wins.
3. Build Word=>Docs as an operation-planned compiler, not a serializer.
4. Treat comments, floating layout, Word charts, SmartArt, and some print/image scenarios as explicit gap areas.
5. Ship a fidelity report so unsupported semantics are visible instead of silently lost.

This is a large but real project. It is not "impossible", and it is not a 30-minute feature either.

## Sources

Official Google documentation used for this analysis:

- Drive uploads and conversion: https://developers.google.com/workspace/drive/api/guides/manage-uploads
- Drive downloads and export: https://developers.google.com/workspace/drive/api/guides/manage-downloads
- Drive export formats: https://developers.google.com/workspace/drive/api/guides/ref-export-formats
- Drive MIME types: https://developers.google.com/workspace/drive/api/guides/mime-types
- Drive comments: https://developers.google.com/workspace/drive/api/guides/manage-comments
- Docs API requests: https://developers.google.com/workspace/docs/api/reference/rest/v1/documents/request
- Docs API structure: https://developers.google.com/workspace/docs/api/concepts/structure
- Docs API request/response model: https://developers.google.com/workspace/docs/api/concepts/request-response
- Docs tables guide: https://developers.google.com/workspace/docs/api/how-tos/tables
- Docs image insertion guide: https://developers.google.com/workspace/docs/api/how-tos/images
- Sheets batchUpdate guide: https://developers.google.com/workspace/sheets/api/guides/batchupdate
- Sheets request reference: https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/request
- Sheets sheet/resource reference: https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/sheets
- Sheets charts guide: https://developers.google.com/workspace/sheets/api/samples/charts
- Sheets pivot tables guide: https://developers.google.com/workspace/sheets/api/guides/pivot-tables
- Google Workspace credentials: https://developers.google.com/workspace/guides/create-credentials
- Service account auth: https://developers.google.com/identity/protocols/oauth2/service-account
