# OfficeIMO and PSWriteOffice API design

This document is the design contract for the unified OfficeIMO and PSWriteOffice public API: the rules every C# type and PowerShell command follows, the shared vocabulary between them, and what is removed. Package READMEs describe current behavior once it ships; remaining delivery work belongs in the [roadmap](ROADMAP.md#unified-public-api).

Status: proposal for review. The redesign ships as OfficeIMO 3.6 and PSWriteOffice 3.6, released together from one pull request. Breaking changes are accepted throughout; nothing keeps a compatibility wrapper unless a section says so.

## Why

OfficeIMO's engines are broad and capable, but the public surface grew one feature at a time:

- PSWriteOffice has 594 cmdlets, 383 command aliases and 1,720 distinct parameter names. The same idea is named differently from area to area: input as `-Path`, `-InputPath` or `-SourcePath`; output as `-Path` or `-OutputPath`; a sheet as `-Sheet` plus `-SheetIndex` on 87 cmdlets; sizes as `-HeightPixels`, `-SpacingBeforePoints`, `-WidthCm` or `-WidthCentimeters`.
- 123 on/off parameters are typed `bool` or `bool?`, so scripts read `-KeepWithNext $true` instead of `-KeepWithNext`.
- 113 cmdlets have no declared output, and no format views ship, so results print as walls of properties, including raw `Bytes`. `-OutputPath .\deck.png` silently becomes a folder.
- The C# libraries expose about 7,000 public types. Word and Excel each have an object model and a separate `AsFluent()` API; Excel has a third, `SheetComposer`. There are three or four ways to convert to PDF. Word and PowerPoint carry the 9.5 MB PDF engine whether or not their users export PDF, while Excel does not.
- PowerPoint geometry defaults to raw EMUs. Excel `AddTable` returns `void` while Word and PowerPoint return the table.
- Ordinary output looks plain. Content slides are a small grey title and unstyled bullets; the designer deck looks designed only until an ordinary slide is added; charts use serif fonts. There is no cross-format theme.

The goal is an API we can show on the first screen of the README and be proud of: small jobs take one line, large jobs stay fast, and the output looks good without styling every element by hand.

## Principles

1. **One vocabulary, two languages.** A concept has one name in C# and PowerShell: `Heading`, `Paragraph`, `Text`, `List`, `Table`, `Chart`, `Image`, `Theme`, `Sheet`, `Slide`, `Page`. A C# developer reading a PowerShell script, or the reverse, recognizes every word.
2. **Designed by default.** Every new document starts from a theme. A plain `Export-OfficeExcel`, `WordTable` or `PptSlide -Title` result is presentable without extra parameters. Styling is for brand and intent, not for repairing defaults.
3. **One way to do each thing.** Duplicate entry points, twin aliases and parallel builder APIs are removed rather than kept for compatibility.
4. **Small and large jobs share one API.** The one-liner and the 100,000-file batch use the same commands. Scale comes from streaming, parallelism and resumable batches, not from a separate expert API.
5. **No silent surprises.** Failures throw or write errors. Conversions report what they could not carry. File-changing commands support `-WhatIf`. Nothing silently becomes a folder or overwrites a file the user did not ask to replace.
6. **Public means supported.** A public C# type or PowerShell command exists because users need it and it is documented. Implementation detail is internal. PSWriteOffice uses only the public C# API; if it needs something the library lacks, that becomes a library feature.
7. **Pay only for what you use.** Installing a format package brings that format, not other format engines. Conversions between formats live in adapter packages; the `OfficeIMO` package bundles the common set for people who want everything.

## Scope

| Tier | Areas | Treatment |
| --- | --- | --- |
| 1 | Word, Excel, PowerPoint, PDF, Markdown, Core (theme, units, results), Reader and search, conversion | Full redesign in C# and PowerShell, visual review of output |
| 2 | Visio, HTML, RTF, OpenDocument, CSV, email, Project, Access, EPUB, OneNote, iWork, AsciiDoc, LaTeX, Confluence, OCR | Shared lifecycle, naming, parameters, units, results and conversion routes. Every capability stays; no feature redesign |
| Out of PSWriteOffice | Electronic invoices (11 cmdlets and the PDF invoice alias) | Removed from the module. The C# `OfficeIMO.Invoicing*` packages are unaffected |

Consolidation removes duplicate commands, never capabilities: every job possible today remains possible, usually through fewer commands.

## Shared foundations (OfficeIMO.Core)

These types are the common language between every format and both languages.

### Lengths

`OfficeLength` is a value type for public geometry: page margins, spacing, image sizes, shape positions and slide coordinates.

```csharp
OfficeLength.Points(12)
OfficeLength.Centimeters(2.5)
OfficeLength.Inches(1)
OfficeLength.Millimeters(10)
OfficeLength.Pixels(96)            // at 96 DPI
OfficeLength.Parse("2.5cm")        // invariant culture
```

- It stores whole EMUs (`long`), which represent points, inches, centimeters, millimeters and 96-DPI pixels exactly. Formats with coarser units (twips, half-points) round to nearest at write time.
- It implements `IEquatable`, `IComparable`, `IFormattable` and `TryParse`. Parsing is culture-invariant (`2.5cm`, never `2,5cm`).
- C# has no implicit conversion from numbers; the unit is always visible in code.
- Public methods never take raw EMUs, twips or half-points, and parameter names never carry a unit suffix (`Width`, not `WidthCm` or `HeightPixels`).
- Excel column widths and row heights are the exception: Excel measures columns in characters of the default font. They take an `ExcelColumnWidth` (characters, or a length converted through the workbook's default font).

In PowerShell a length parameter accepts a number, meaning points, or a value with a unit. PowerShell passes unit values unquoted as strings, so `-Width 5cm`, `-SpacingBefore 12pt`, `-Margin 1in` and `-Width 140` all work as written.

### Colors

The existing `OfficeColor` becomes the only public color type. It accepts `#RRGGBB`, `#RRGGBBAA`, CSS names and theme roles.

- `Accent1`–`Accent6`, `Text`, `TextMuted` and `Background` map to the document's theme slots and keep following the theme when someone changes it in Office.
- `Success`, `Warning` and `Danger` resolve from the theme when written. Office has no slot for them, so they are stored as literal colors.

### Themes

`OfficeTheme` describes a visual identity that every format applies:

- palette: background, text, muted text, six accents and `Success`, `Warning`, `Danger`;
- fonts: heading and body, with fallbacks that render on Windows, macOS and Linux;
- type scale: title, heading levels, body, caption;
- table style: header, banding, borders, padding;
- chart style: series palette, sans-serif axis and title text, light gridlines, label rules;
- spacing rhythm for paragraphs, lists and slide margins.

Built-in themes: `Modern` (the default), `Classic`, `Minimal`, `Boardroom`, `Dark`. `OfficeTheme.FromDocument(path)` reads a corporate theme from an existing `.potx`, `.pptx`, `.dotx`, `.docx`, `.xltx` or `.xlsx`, so a company template drives generated output.

Each format maps the theme to its native constructs: Word styles and theme part; Excel cell styles, table styles and chart formats; PowerPoint master, layouts and theme part; PDF builder defaults; HTML CSS variables. The HTML-only `OfficeVisualThemeKind` is replaced by `OfficeTheme`.

A theme applies when a document is created. Loading or opening an existing document never restyles it; `ApplyTheme` does that on request.

### Shared enums and options

- Concepts that exist in several formats have one enum in Core, prefixed `Office`: `OfficeChartType`, `OfficeAlignment`, `OfficeImageFormat`, `OfficePageOrientation`.
- Format-only concepts keep a format prefix: `ExcelTableStyle`, `WordBreakType`.
- Shared options follow the same rule: `OfficeTableOptions`, `OfficeChartOptions`.

### Results and diagnostics

- Failures (unreadable input, I/O, invalid arguments) always throw in C# and write errors in PowerShell.
- Converting to another format reports loss and continues; `FailOnLoss` turns loss into an error.
- Saving a document in its own format fails if content would be lost, as it does today; `AllowLoss` accepts the loss explicitly.
- `OfficeConversionReport` is the single report type: source and target, output path, pages or items written, elapsed time, and a list of `OfficeDiagnostic`.
- `OfficeDiagnostic` has `Code`, `Severity`, `Message`, `Location` and `Format`. Codes come from one catalog in Core, so a flattened 3-D chart is `chart-flattened` in Word, Excel and PowerPoint alike.
- `report.HasLoss`, `report.ThrowIfLoss()`, `FailOnLoss` for conversions and `AllowLoss` for native saves are the only loss controls.
- PowerShell errors use the diagnostic code as the error ID, so `FullyQualifiedErrorId` starts with the same code the C# report carries; PowerShell appends the implementing cmdlet type after the comma, as it does for every binary cmdlet.

## C# rules

### Lifecycle

Every document type, including PDF, has the same shape:

```csharp
using var doc = WordDocument.Create();                     // in memory
using var doc = WordDocument.Create("Report.docx", options);
using var doc = WordDocument.Load("Report.docx", options);
await using var doc = await WordDocument.LoadAsync(stream, options);

doc.Save();                  // to the path it was created or loaded with
doc.Save("Copy.docx");
await doc.SaveAsync(stream);
```

- Options are `{Format}CreateOptions`, `{Format}LoadOptions` and `{Format}SaveOptions`. Each load options type has `UntrustedDefaults` for files from outside.
- `PdfDocument` joins this shape: `Create`, `Load`, `Save`, `IDisposable`, `PdfCreateOptions` and `PdfLoadOptions.UntrustedDefaults`. Its immutable internal pipeline stays an implementation detail.
- Static inspection helpers (`PdfInspector.Inspect` and similar) move onto the document as `Inspect()`, plus `OfficeDocument.Inspect(path)` for any format.

### Authoring

There is one authoring model per format: the object model. The `AsFluent()` APIs, `WordFluentDocument`, `ExcelFluentWorkbook` and their builders are removed. Their useful capabilities become methods or options on the object model.

- `Add*` methods return the created object: `ExcelSheet.AddTable` returns `ExcelTable`, `AddImage` returns the image, `AddChart` returns the chart.
- One verb per concept: `AddImage` everywhere. Word's `InsertImage` and PowerPoint's `AddPicture` are removed.
- Typed data in, styled output out, the same in Word, Excel and PowerPoint:

```csharp
doc.AddTable(rows);                                         // IEnumerable<T>, columns from properties
doc.AddChart(OfficeChartType.Line, trend, t => t.Week,
    OfficeSeries.From(trend, t => t.Complete, "Complete"));
doc.AddChart(OfficeChartType.Column, trend, nameof(Trend.Week), nameof(Trend.Revenue), nameof(Trend.Margin));
```

  Series values accept any numeric type, and series names come from the property name or are given explicitly.
- Excel's `SheetComposer` stays as the report builder, reached through `sheet.Report()`, because it is a real feature (titled sections, KPI rows, legends) rather than a second way to write cells.

### Large workbooks

The object model keeps the whole workbook editable until `Save()`. Very large exports use an explicit streaming writer instead, which writes rows as they arrive and returns no live table:

```csharp
using var writer = ExcelDocument.CreateWriter("Big.xlsx", theme: OfficeTheme.Modern);
writer.WriteTable("Events", events);   // IEnumerable<T> or IAsyncEnumerable<T>, bounded memory
```

`Export-OfficeExcel` uses the streaming writer whenever it creates the whole file, and the object model when it appends to an existing workbook.

### Conversion

Two forms, the same in every format, always called on the source document:

| Form | Meaning | Example |
| --- | --- | --- |
| `SaveAs{Target}(path or stream, options)` | Write the target file; returns `OfficeConversionReport` | `doc.SaveAsPdf("Report.pdf")` |
| `To{Target}(options)` | Return the target document model in memory | `MarkdownDocument md = doc.ToMarkdown();` |


- Converting into a format is a conversion from the source: `MarkdownDocument.Parse(text).ToWord()`, `HtmlDocument.Parse(html).ToWord()`. There are no static `From…` factories, because adapters would need C# 14 static extension members to add them.
- There is one method per form and target. The throwing and `…Result` pairs merge, for conversions (`SaveAsPdf` / `SaveAsPdfResult`) and for other operations that come in pairs today, such as the PDF editing operations.
- `To{Target}` has an overload with `out OfficeConversionReport report` for callers who want the loss report. Async forms (`To{Target}Async`) return `OfficeConversionResult<T>`, which carries the value and the report, because async methods cannot have `out` parameters.
- `ToHtml()` returns a string. HTML is the one target whose model users rarely want; `HtmlDocument.Parse` gives the model when needed.
- `OfficeConverter` is the universal entry point for file-to-file and batch work. It absorbs `OfficeWorkflow.Convert`, the `OfficeDocumentPdfConverter` façades and the `*ToPdfBytes` helpers. Studio and the CLI save jobs as the same `OfficeBatchOptions`.

```csharp
OfficeConversionReport report = OfficeConverter.Convert("In.docx", "Out.pdf");

OfficeBatchReport batch = await OfficeConverter.ConvertAsync(
    Directory.EnumerateFiles("Reports", "*.docx"),
    outputDirectory: "Pdf",
    target: OfficeDocumentFormat.Pdf,
    new OfficeBatchOptions { MaxParallelism = 8, ResumeFrom = "Pdf/.officeimo-batch.json" });
```

### Packages

Package boundaries follow principle 7.

- A format package (`OfficeIMO.Word`, `OfficeIMO.Excel`, `OfficeIMO.PowerPoint`, `OfficeIMO.Pdf`, …) depends only on Core and the format's own third-party library, such as the Open XML SDK. It never depends on another format engine.
- Conversions live in adapter packages named `OfficeIMO.{Source}.{Target}`: `OfficeIMO.Word.Pdf`, `OfficeIMO.Excel.Pdf`, `OfficeIMO.PowerPoint.Pdf`, `OfficeIMO.Word.Markdown`, `OfficeIMO.Word.Html`. Each brings only the two engines it joins. Word and PowerPoint PDF export, which moved into the format packages during 3.x, move back out, so a Word user no longer carries the PDF engine.
- Adapters do not chain through a third format. `OfficeIMO.Word.Markdown` currently pulls in `OfficeIMO.Word.Html` and with it the HTML engine and AngleSharp; Markdown to Word should not need HTML.
- Reconstruction from PDF (PDF to Word, Excel or PowerPoint) lives in the matching `OfficeIMO.Pdf.{Target}` adapter.
- Image export stays in the Excel and PowerPoint packages, whose renderers draw with Core's drawing types. Word paginates through the PDF layout engine, so Word page images live in `OfficeIMO.Word.Pdf` together with PDF export.
- `OfficeIMO.Word.Pdf` and `OfficeIMO.PowerPoint.Pdf` change direction in 3.6: today they convert PDF into Word and PowerPoint; afterwards they convert Word and PowerPoint into PDF, and the reverse lives in `OfficeIMO.Pdf.Word` and `OfficeIMO.Pdf.PowerPoint`. `MIGRATION.md` states this first, because the package names stay the same while their contents change.
- The `OfficeIMO` package is the front door. It references Word, Excel, PowerPoint, PDF and Markdown, the adapters between them, `OfficeConverter` and Reader with its format handlers. `dotnet add package OfficeIMO` is the first line of the README. The NuGet ID is already registered to this project (two unlisted 0.1.x versions from 2022).
- Namespaces stay per format (`OfficeIMO.Word`, `OfficeIMO.Excel`). The `OfficeIMO` namespace already holds Core's shared types. All shared foundations live there, including `OfficeColor` and `OfficeColorRole`, which move from `OfficeIMO.Drawing`; `OfficeConverter` and the batch types join them, and the existing `OfficeDocumentFormat` enum names formats.
- Every library keeps the frameworks it supports today, and 3.6 drops none of them: `netstandard2.0`, `net472`, `net8.0` and `net10.0`. Windows builds add `net472`, as they do today. This includes the `OfficeIMO` front door, `OfficeConverter`, the batch engine and `OfficeIMO.Workflows`, so .NET Framework applications and Windows PowerShell 5.1 get the same conversion routes as modern .NET. The batch engine runs on every framework and uses newer runtime APIs where they exist. `OfficeIMO.Workflows` and `OfficeIMO.Workflows.IWork`, which target only modern .NET today, gain the same targets. A route whose engine is modern-only (for example invoice validation) is registered only on the frameworks that engine supports. Other packages that target only modern .NET today (Tool, AI, Odbc, Studio, Data.Arrow, the KSeF and browser-PDF adapters) keep their targets.
- `net8.0` is dropped in a later release, after .NET 8 reaches end of support on 10 November 2026, not as part of 3.6.
- `IsAotCompatible` and the trimming analyzers are on for the `net8.0` and `net10.0` builds of every library. The `netstandard2.0` and `net472` builds keep a test run on .NET Framework on Windows.
- Packages that pass trimming and NativeAOT analysis declare `IsAotCompatible`.

### Naming

- The Markdown model is `MarkdownDocument`, loaded with `MarkdownDocument.Load` or `Parse`, like every other format.
- Format types start with the format: `WordTable`, `ExcelTable`, `PowerPointTable`, `PdfTable`. Markdown syntax-tree node types (`HeadingBlock`, `ParagraphBlock`) are the exception; the `OfficeIMO.Markdown` namespace already qualifies them.
- Every document type implements `IDisposable`, doing nothing where there is nothing to release, so `using` works the same for every format.
- Options end in `Options`, results in `Report`, read-only snapshots in `Info`.

### Visibility

A type or member stays public only when it is part of the documented user API: it appears in a README, guide or example, or it is an options, result or model type of something that does. Cross-package plumbing between OfficeIMO's own assemblies uses `InternalsVisibleTo` or shared source, never public API.

The candidate list is generated by analysis, not by hand, and reviewed per package. After cleanup, Tier 1 packages adopt the public API analyzer with a shipped baseline, so new public surface is a deliberate, reviewed change.

## PowerShell rules

### Command names

- `Verb-Office{Format}{Noun}` for format commands, with approved verbs only. Format names are `Word`, `Excel`, `PowerPoint`, `Pdf`, `Markdown`, `Visio`, `Html`, `Rtf`, `OpenDocument`, `Csv`, `Email`, `Project`, `Access`, `Epub`, `OneNote`, `IWork`, `AsciiDoc`, `Latex`, `Confluence`.
- `Verb-OfficeDocument{Noun}` for commands that work on any supported format.
- A `Test-` command never changes a file. Changes go through `Repair-`, `Set-` or `Update-`.
- Internal or diagnostic concepts are not separate commands. `Get-OfficeExcelStreamingContract`, `Get-OfficeExcelWriteReservation`, `Set-OfficeExcelExecutionPolicy` and similar become parameters or properties of the commands they serve.

### Cross-format commands

These cover most everyday jobs for every supported format:

| Command | Job |
| --- | --- |
| `Convert-OfficeDocument -Path -To -OutputPath` | Convert one file or a batch between document formats. Replaces the file-to-file `ConvertTo-*`/`ConvertFrom-*` commands and `Export-OfficeDocumentPdf` |
| `Read-OfficeDocument -Path` | Text, tables, chunks and structure for search and AI ingestion |
| `Search-OfficeDocument -Path -Pattern` | Find content across mixed files and folders. `-Pattern` is a regular expression, with `-SimpleMatch` and `-CaseSensitive`, as in `Select-String` |
| `Get-OfficeDocumentInfo -Path` | Metadata, page and sheet counts, protection, features in use |
| `Export-OfficeDocumentImage -Path -OutputDirectory` | Render pages, slides or sheets to PNG, JPEG or SVG. Replaces eight `Export-Office*Image` commands |
| `Test-OfficeDocument` / `Repair-OfficeDocument` | Validate, then fix |
| `Protect-OfficeDocument` / `Unprotect-OfficeDocument` | Passwords and restrictions |
| `Compare-OfficeDocument`, `Merge-OfficeDocument`, `Split-OfficeDocument` | Version comparison, combining and splitting for the formats that support them |

`Convert-OfficeDocument` writes one report per file to the pipeline as each finishes, runs in parallel with `-ThrottleLimit`, and continues an interrupted batch with `-Resume`.

Some current `ConvertTo-`/`ConvertFrom-` commands are not file conversions, and they keep their own commands:

- **Object-to-text conversions** follow PowerShell's `ConvertTo-Html` / `ConvertTo-Json` pattern and stay as `ConvertTo-` commands: `ConvertTo-OfficeMarkdown` (objects to a Markdown table) and `ConvertTo-OfficeCsv`.
- **PDF transforms** become commands with matching verbs: `Optimize-OfficePdf` (size and sanitization), `Remove-OfficePdfContent` (permanent redaction of text, patterns or areas), and `Update-OfficePdf` (flatten forms and annotations, add a searchable OCR text layer). Passwords, encryption and permissions for PDF go through `Protect-OfficeDocument` like every other format.

### Document lifecycle

Every authoring format has the same four commands:

```powershell
New-Office{Format}   -Path [-Theme] [-Force] [{ content }]
Open-Office{Format}  -Path [{ content }]
Save-Office{Format}  -Document [-Path]
Close-Office{Format} -Document
```

- With a content block, `New-` and `Open-` save and close when the block ends. They return nothing, or the file with `-PassThru`. If the block throws, nothing is saved, a partially written new file is removed, and the error propagates.
- Without a content block, they return the live document, which the script saves and closes. `-NoSave` is removed; leaving out the block is how to keep a document open. `New-` without a block does not need `-Path`; `Save-Office{Format} -Path` names the file later.
- `New-` refuses to replace an existing file unless `-Force` is given, like `New-Item`.
- `Open-` replaces the `Get-Office{Format}` commands that open files. `Get-` is for reading content out of a document.
- `-Show` opens the saved file in its default application.

### Parameters

One name, one meaning, everywhere. The set below is fixed; a new command does not invent synonyms.

| Parameter | Meaning |
| --- | --- |
| `-Path` | The input file when the command reads one (`Open-`, `Import-`, `Read-`, `Search-`, `Convert-`, `Export-…Image`). Otherwise the file the command creates (`New-`, `Export-OfficeExcel`, `Save-` to a new name). Supports wildcards when it is an input |
| `-LiteralPath` | Input without wildcard expansion, in its own parameter set. Binds from the pipeline by property name through `PSPath` and `FullName`, as in `Get-Content`, so `Get-ChildItem … | Convert-OfficeDocument` works for any file name |
| `-OutputPath` | The output file when the command also has an input file. Never silently treated as a folder |
| `-OutputDirectory` | The output folder when a command writes several files |
| `-Document` | A live document from `New-`/`Open-` or the pipeline. Replaces `-Presentation`, `-Workbook` and similar |
| `-InputObject` | Data rows for tables, charts and exports. An array is enumerated as rows, from the pipeline or as an argument, unlike `Export-Csv -InputObject` |
| `-Sheet` | A worksheet by name (string), 1-based position (integer) or object. Replaces the `-Sheet` + `-SheetIndex` pair. A sheet named `1` is selected with `-Sheet '1'` |
| `-Theme` | A built-in theme name, a theme object or a template path |
| `-Password` | `SecureString`. Argument transformation also accepts a plain string (with a warning) or a `PSCredential`, whose password is used; there is no separate `-Credential` parameter. `SecureString` is not encrypted outside Windows; it keeps passwords out of history, logs and transcripts |
| `-Culture` | Culture for parsing and formatting numbers and dates in data commands. Defaults to the current culture; lengths and other parameters always parse invariantly |
| `-PassThru` | Return what was created or changed |
| `-Force` | Allow `New-` to replace an existing file |
| `-NoClobber` | Make `Export-` and `Convert-` refuse to replace an existing output. By default they replace it, like `Export-Csv` |
| `-Show` | Open the result in its default application |
| `-Category` / `-Value` | The properties a chart reads from `-InputObject`. Replaces `-CategoryProperty` / `-SeriesProperty` |

Other rules:

- **Positions are 1-based** everywhere: pages, slides, sheets, sections, table rows and columns, list items and insert positions, matching how Office numbers them.
- **Parameter sets instead of runtime checks.** Mutually exclusive inputs (`-Path` / `-LiteralPath` / `-Document`, `-Text` / content block, `-Url` / `-Sheet` targets) are parameter sets with a sensible default, so tab completion and `Get-Help` show only valid combinations. "Use either -X or -Y" errors disappear.
- **Switches for on/off.** The 123 `bool` and `bool?` parameters become switches. A shared base class reads `BoundParameters` to tell "not given" from `-Bold:$false`, so formatting can still be turned off explicitly. Defaults that are on are turned off with `-No{Feature}` switches (`-NoFreeze`, `-NoAutoFit`). Numeric and enum parameters keep their types.
- **Options objects** come from `New-Office{Format}{Purpose}Options` commands, which give tab completion and help. Every `-Options` parameter also accepts a hashtable of option property names, converted to the options type the command needs; unknown keys are errors.
- **Lengths, colors and themes** use the shared Core types with argument transformation and tab completion. Every enum and style completes.
- **`-WhatIf` and `-Confirm`** cover file-system effects only: `New`, `Save`, `Export`, `Convert`, `Repair`, `Protect`, `Remove`. Commands that add content inside a document block do not support them, so `New-OfficeWord -WhatIf { … }` builds nothing and writes nothing, without asking once per paragraph.

### The authoring DSL

The DSL is the same tree in every format, and shares names with C#. Inside a block, the commands write to the enclosing document, section, sheet or slide.

| Concept | Word | PowerPoint (inside a slide) | PDF | Markdown | Excel |
| --- | --- | --- | --- | --- | --- |
| Heading | `WordHeading` | `PptSlide -Title` | `PdfHeading` | `MarkdownHeading` | — |
| Paragraph | `WordParagraph` | `PptParagraph` | `PdfParagraph` | `MarkdownParagraph` | — |
| Styled text | `WordText` | `PptText` | `PdfText` | `MarkdownText` | `ExcelText` (rich text inside a cell) |
| List | `WordList` / `WordListItem` | `PptList` / `PptListItem` | `PdfList` / `PdfListItem` | `MarkdownList` / `MarkdownListItem` | — |
| Table | `WordTable` | `PptTable` | `PdfTable` | `MarkdownTable` | `ExcelTable` |
| Chart | `WordChart` | `PptChart` | `PdfChart` | — | `ExcelChart` |
| Image | `WordImage` | `PptImage` | `PdfImage` | `MarkdownImage` | `ExcelImage` |
| Header and footer | `WordHeader` / `WordFooter` | `PptFooter` | `PdfHeader` / `PdfFooter` | — | `ExcelHeader` / `ExcelFooter` |
| Page number | `WordPageNumber` | `PptSlideNumber` | `PdfPageNumber` | — | — |
| Speaker notes | — | `PptNotes` | — | — | — |

- The same concept takes the same parameters in every format. `WordText`, `PptText`, `PdfText` and `MarkdownText` share `-Text`, `-Bold`, `-Italic`, `-Strikethrough`, `-Color`, `-Font`, `-Size` and `-Link`; `-Underline` takes the shared underline style where the format supports it. A list item takes `-Text` or a content block.
- Styled runs are written as child commands. The `-Run @{ Text = …; Bold = $true, $true, $false }` hashtables of parallel arrays are removed.
- Every DSL command has exactly one alias: the format prefix plus the noun, singular unless the Office term itself is plural (`PptNotes`). `Ppt` is the only PowerPoint prefix.
- Aliases exist only for commands that add content inside a content block. Commands that change the current element (`Set-OfficeWordParagraph`) and deck- or document-level settings are called by their full name inside a block; their natural alias would collide with the command that creates the element. Plural twins (`WordCheckBoxes`), spelling twins (`PdfHr` / `PdfHorizontalRule`), verb-last forms (`ExcelAutoFilterSet`), and unprefixed or verb-noun aliases (`ConvertTo-Rtf`, `TextRun`, `Edit-ExcelRow`, `Replace-OfficeWordText`) are removed.
- A test rejects aliases that break these rules or collide with commands in PowerShell itself.

### Consolidation

Families of near-duplicate commands become one command with parameter sets. The full old-to-new map is generated as machine-readable data in M0. The main families:

| Today | After |
| --- | --- |
| `Import-OfficeExcel`, `Get-OfficeExcelData`, `Get-OfficeExcelRange`, `Get-OfficeDocumentTable` for reading rows | `Import-OfficeExcel` with `-Sheet`, `-Range` and `-Table` |
| Seven `Set-OfficeExcel*Link*` commands | `Set-OfficeExcelHyperlink` with URL, internal-sheet and by-header sets |
| Seven `Add-OfficeExcelValidation*` commands | `Add-OfficeExcelValidation -Type List | WholeNumber | Decimal | Date | Time | TextLength | Formula` |
| Four `Add-OfficeExcelConditional*` commands | `Add-OfficeExcelConditionalFormat` with rule, color-scale, data-bar and icon-set sets |
| `Get-OfficePdf`, `Read-OfficePdf`, `Get-OfficePdfText`, `Get-OfficePdfInfo` | `Open-OfficePdf` for editing; `Read-OfficeDocument` and `Get-OfficeDocumentInfo` for reading |
| File-to-file `ConvertTo-*` and `ConvertFrom-*` commands, `Export-OfficeDocumentPdf` | `Convert-OfficeDocument`, with format-specific options through `-Options` |
| Eight `Export-Office*Image` commands | `Export-OfficeDocumentImage` |
| Five CSV commands | `Import-OfficeCsv`, `Export-OfficeCsv` and `ConvertTo-OfficeCsv`, kept for formula-injection protection, typed columns and compressed files |
| `New-OfficePowerPointDeckPlan`, `Add-OfficePowerPointPlan*`, `Add-OfficePowerPointDesignerDeck` | Theme layouts and automatic layout selection (see [PowerPoint](#powerpoint)) |

The goal is fewer, stronger commands. The exact count comes out of the map; roughly half of today's commands is the expected order.

The same map drives three things:
- the 3.6 section of `MIGRATION.md`;
- a test that every old command name has a recorded replacement;
- a hint when a script calls a removed command. The module registers a command-not-found handler that prints the replacement, for example "`Get-OfficeExcelData` was replaced by `Import-OfficeExcel -Range`".

### Output

- Every command declares `[OutputType]`.
- The module ships format views for every returned type: documents, paragraphs, tables, slides, sheets, reports, diagnostics, search hits and Reader chunks. Binary content such as `Bytes` never appears in default output.
- Batch commands write each result to the pipeline as it finishes, with `Write-Progress`. Parallel work runs on background tasks and hands results back to the pipeline thread, because PowerShell only accepts output and progress from there.
- Errors in a batch are non-terminating, carry the input as their target object and use proper error categories. Invalid arguments are terminating.

### Help

- Every command has at least two working examples: the smallest useful call and a realistic one. Examples are executed in tests, not only parsed.
- `about_PSWriteOffice` and one `about_PSWriteOffice_{Format}` topic per Tier 1 format explain the model, one page each.

### Platform support

- PSWriteOffice keeps its current runtime reach: the PowerShell 7 payload is built for `net8.0` and runs on PowerShell 7.4 and later, including 7.6 on .NET 10, and Windows PowerShell 5.1 loads its own `net472` payload. Import checks the edition and picks the payload. The PowerShell 7 payload moves to `net10.0` (PowerShell 7.6 and later) in the same later release that drops `net8.0` from the libraries.
- The module loads its third-party dependencies (Open XML SDK, AngleSharp and others) in an isolated assembly load context, so they never clash with other modules. OfficeIMO's own assemblies load so that type names such as `[OfficeIMO.Word.WordDocument]` work in user scripts. M3 qualifies the exact mechanism.
- Windows PowerShell 5.1 stays supported for the commands that work there. The manifest description and the README say what is missing on 5.1, and assembly resolution problems are reported through PowerShell's error stream, not the console.

## Visual quality

Visual quality is a release criterion, not polish.

### Defaults in every format

- The `Modern` theme is applied to new documents unless another is chosen.
- Charts use the theme palette, sans-serif text, light gridlines and no decorative borders.
- Tables use the theme table style: a header row, light banding, comfortable padding and right-aligned numbers.
- Text uses the theme type scale and spacing rhythm.

### PowerPoint

Creating a deck has to be as easy as creating a Word document, and the result has to look designed. Nobody should need coordinates to make a good slide.

There are three ways to make a deck, from least to most effort:

1. **From data, one line.** `$rows | Export-OfficePowerPoint -Path .\Review.pptx -Title 'Service review'` makes a title slide, then table slides (splitting long tables across slides) and optional chart slides from `-Category`/`-Value`.
2. **From an outline.** `Convert-OfficeDocument -Path .\Review.md -To PowerPoint` treats each `---` or top-level heading as a slide: headings become titles, lists become bullets, tables and images flow into place, and blockquotes become speaker notes.
3. **With the DSL**, for full control, still without coordinates.

Slides are built from theme layouts with designed placeholders:

| Layout | Use |
| --- | --- |
| `Title` | Cover slide: eyebrow, title, subtitle, date |
| `Section` | Section divider |
| `TitleAndContent` | Title with a list, table, chart or image |
| `TwoContent` / `Comparison` | Two side-by-side blocks, optionally with headings |
| `Kpi` | Two to six metrics with labels and change indicators |
| `Cards` | Three to six cards (today's card grid) |
| `Process` | Three to six steps (today's process slide) |
| `Quote`, `ImageFull`, `Closing` | Statement, full-bleed image and closing slides |
| `CaseStudy` | Challenge, approach and result with two or three metrics |
| `Logos` | A wall of four to twelve logos or product images |
| `Coverage` | A map or diagram with pinned locations |

- `-Layout` is optional. By default the layout is chosen from the content: a title only gives a section slide; one list, table, chart or image gives `TitleAndContent`; two blocks give `TwoContent`; `PptKpi` items give `Kpi`. The first slide of a deck uses `Title`.
- Content that is too long for a slide is handled rather than overflowing: lists shrink to the theme's minimum size, then continue on a new slide; tables split across slides with the header repeated.
- `-X`, `-Y`, `-Width` and `-Height` remain for deliberate placement. They take lengths and can snap to the theme grid (`-Column 1 -Span 6`).
- The designer-deck engine becomes the implementation behind these layouts, so every slide in a deck shares one design. Corporate `.potx` templates supply their own layouts through `-Theme`.

### Word, Excel and PDF

- Word: theme heading and body styles, a consistent heading hierarchy, a footer whose `Page N of M` stays on one line (it currently splits onto two), and table and chart defaults from the theme.
- Excel: `Export-OfficeExcel` and `AddTable` produce a styled table with a frozen header, fitted columns and number formats inferred from the data. `-NoFreeze`, `-NoAutoFit` and `-NoTableStyle` turn each one off.
- PDF: the composition builder takes its fonts, colors, table style and spacing from the theme.

### Acceptance

Each built-in theme and layout is checked by eye on macOS and Windows in Microsoft Office, LibreOffice and OfficeIMO's own renderer, using the hero examples. Reference renders are kept as evidence.

## Scale and performance

- `Export-OfficeExcel` and `ExcelDocument.CreateWriter` stream rows with bounded memory, so exporting a million rows does not hold the pipeline in memory.
- `Convert-OfficeDocument`, `Read-OfficeDocument` and `Search-OfficeDocument` process folders in parallel with `-ThrottleLimit`. `Convert-OfficeDocument -Resume` continues an interrupted batch.
- Module import stays under one second on PowerShell 7. It measured 0.9 seconds on 26 October 2026 on an Apple silicon Mac.
- Timing and memory budgets are measured in the opt-in performance runs described in `AGENTS.md`, not in pull-request gates.

## What the result looks like

### PowerShell

```powershell
# Small: one line each
Get-Process | Export-OfficeExcel -Path .\Processes.xlsx -Show
Get-ChildItem .\Reports -Filter *.docx | Convert-OfficeDocument -To Pdf -OutputDirectory .\Pdf
Search-OfficeDocument -Path .\Contracts -Pattern 'termination'
Get-Service | Where-Object Status -eq Stopped | Export-OfficePowerPoint -Path .\Stopped.pptx -Title 'Stopped services'

# A themed report
New-OfficeWord -Path .\Status.docx -Theme Modern {
    WordHeading 'Northwind migration'
    WordParagraph {
        WordText 'Overall status: ' -Bold
        WordText 'On track' -Bold -Color Success
    }
    WordTable -InputObject $milestones {
        WordTableCondition { $_.Status -eq 'At risk' } -BackgroundColor Danger
    }
    WordChart -Type Line -InputObject $trend -Category Week -Value Complete -Title 'Completion by week'
    WordFooter { WordText 'Internal · Page '; WordPageNumber -IncludeTotal }
}

# A deck without coordinates or layout names
New-OfficePowerPoint -Path .\Review.pptx -Theme Boardroom {
    PptSlide -Title 'Quarterly business review' -Subtitle 'Performance, decisions and priorities'
    PptSlide -Title 'Performance trend' {
        PptChart -InputObject $trend -Category Quarter -Value Revenue -Type Column -Title 'Revenue (USD m)'
        PptChart -InputObject $trend -Category Quarter -Value Margin -Type Line -Title 'Margin (%)'
        PptNotes 'Explain the margin before revenue.'
    }
    PptSlide -Title 'Next-quarter priorities' { PptTable -InputObject $priorities }
}
```

### C\#

```csharp
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

// Small
OfficeConverter.Convert("Status.docx", "Status.pdf");

// A themed workbook
using var book = ExcelDocument.Create("Orders.xlsx", new() { Theme = OfficeTheme.Modern });
book.AddSheet("Orders").AddTable(orders);   // styled, frozen header, fitted columns
book.Save();

// A report, saved twice
using var doc = WordDocument.Create("Status.docx");
doc.AddHeading("Northwind migration", level: 1);
doc.AddTable(milestones);
doc.AddChart(OfficeChartType.Line, trend, nameof(Trend.Week), nameof(Trend.Complete));
doc.Save();
OfficeConversionReport pdf = doc.SaveAsPdf("Status.pdf");
```

## Optional console experience

PSWriteOffice ships its own format views and depends on no other module. When [ConsoleMonster](https://github.com/EvotecIT/ConsoleMonster) is installed, PSWriteOffice can use it for richer console output: document outlines as trees, inspection and conversion reports as panels, and live progress for large batches. Detection happens at run time and is opt-in. This comes after the core redesign.

## Delivery

The milestones and their order are in the [roadmap](ROADMAP.md#unified-public-api).
