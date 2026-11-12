# OfficeIMO.Word

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word)](https://www.nuget.org/packages/OfficeIMO.Word)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word)

Create, edit, read and save Word documents (`.docx`, and the supported subset of Word 97–2003 `.doc`) in .NET, without Microsoft Word or COM. New documents use the `Modern` theme for headings, body text, tables and charts. Open supported [Word 6 and Word 95 files](../Docs/officeimo.word.legacy-doc-compatibility.md#import-word-6-and-word-95-documents) to migrate editable content to DOCX.

## Install

```sh
dotnet add package OfficeIMO.Word
```

Add an adapter only for the conversions you need:

```sh
dotnet add package OfficeIMO.Word.Pdf        # SaveAsPdf, page images
dotnet add package OfficeIMO.Word.Markdown   # ToMarkdown, Markdown to Word
dotnet add package OfficeIMO.Word.Html       # ToHtml, SaveAsHtml
```

Targets `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.

## Create a report

<!-- powerforge:code source="Build/DocumentationExamples/Word/Program.cs" region="QuickStart" -->
```csharp
using OfficeIMO;
using OfficeIMO.Word;

Milestone[] milestones = [new("Discovery", "Product", "Done"), new("Pilot", "Operations", "At risk"), new("Rollout", "IT", "Planned")];

using var document = WordDocument.Create("Status.docx");
document.AddHeading("Northwind migration", level: 1);

WordParagraph status = document.AddParagraph("Overall status: ");
status.AddText("On track").Bold = true;

document.AddHeading("Milestones", level: 2);
document.AddTable(milestones);

document.Save();

record Milestone(string Name, string Owner, string Status);
```
<!-- /powerforge:code -->

`AddTable` takes any `IEnumerable<T>`: columns come from public properties, the header row repeats on every page and numbers are right-aligned. `OfficeTableOptions` picks, orders and renames columns:

```csharp
using OfficeIMO;
using OfficeIMO.Word;

Milestone[] milestones = [new("Discovery", "Product", "Done"), new("Pilot", "Operations", "At risk")];

using var document = WordDocument.Create("Milestones.docx");
var options = new OfficeTableOptions { Name = "Milestones" };
options.Order.AddRange(["Status", "Name"]);
options.Exclude.Add("Owner");
options.HeaderRenames["Name"] = "Milestone";
document.AddTable(milestones, options);
document.Save();

record Milestone(string Name, string Owner, string Status);
```

## Charts

Charts are native, editable Word charts with their data in an embedded workbook. Name the category and value properties, or pass selectors with `OfficeSeries`:

```csharp
using OfficeIMO;
using OfficeIMO.Word;

Week[] trend = [new("W1", 20, 18), new("W2", 45, 40), new("W3", 70, 66), new("W4", 92, 85)];

using var document = WordDocument.Create("Trend.docx");
document.AddHeading("Completion by week", level: 1);
document.AddChart(OfficeChartType.Line, trend, nameof(Week.Label), nameof(Week.Planned), nameof(Week.Complete));
document.AddChart(OfficeChartType.Column, trend, t => t.Label,
    new OfficeChartOptions { Title = "Complete (%)", Width = OfficeLength.Centimeters(14) },
    OfficeSeries.From(trend, t => t.Complete, "Complete"));
document.Save();

record Week(string Label, int Planned, int Complete);
```

## Headers, footers and images

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Create("Letter.docx");
document.PageSetup.Header.AddParagraph("Northwind Traders");
document.PageSetup.Footer.AddParagraph("Page ").AddPageNumber(includeTotalPages: true);

document.AddImage("logo.png", width: OfficeLength.Centimeters(4));
document.AddParagraph("Dear customer,");
document.Save();
```

## Edit an existing document

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Load("Status.docx");
int replaced = document.ReplaceAll("Northwind", "Contoso");
document.AddParagraph($"Updated on {DateTime.Today:d}; {replaced} replacements.");
document.Save("Status-updated.docx");
```

`Find` and `ReplaceAll` match ordinally and ignore case by default; pass `new OfficeTextSearchOptions { MatchCase = true }` or `WholeWords = true` to change that. `Find` returns a `WordSearchResult` with the matching paragraphs per document part, the match count and each match as a `WordRange`, for both literal text and a `Regex`.

### Text ranges

A `WordRange` is a span of text that can start and end inside runs and cross paragraphs, in the body, tables, headers, footers, footnotes and endnotes. Get one from a search hit, a bookmark (`bookmark.GetRange()`) or a paragraph (`paragraph.GetRange(start, length)`, extended with `ExtendTo`). Read `Text`, then replace, format, bookmark, comment on or link it:

```csharp
using System.Text.RegularExpressions;
using OfficeIMO.Word;

using var document = WordDocument.Load("Contract.docx");

// Regex replace with group substitution, a callback, or formatted content.
document.ReplaceAll(new Regex(@"(\d{2})/(\d{2})/(\d{4})"), "$3-$2-$1");
document.ReplaceAll(new Regex(@"\bTBD\b"), match => "to be confirmed");
document.ReplaceAll(new Regex(@"\[(.+?)\]\((.+?)\)"),
    (match, p) => p.AddHyperLink(match.Groups[1].Value, new Uri(match.Groups[2].Value), addStyle: true));

WordRange amount = document.Find("1,000 EUR").Ranges[0];
amount.AddBookmark("Amount");                   // spans exactly "1,000 EUR"; REF fields show it
amount.AddComment("Przemysław", "PK", c => c.AddText("Confirm with ").AddText("finance").SetBold());
document.Bookmarks.First(b => b.Name == "Amount").Text = "2,500 EUR";

foreach (WordRange hit in document.Find("OfficeIMO").Ranges) hit.Format(run => run.SetBold());
document.Find("officeimo.com").Ranges[0].AddHyperLink(new Uri("https://officeimo.com"));
document.Save();
```

`Replace(string)` gives the new text the formatting of the first replaced character; `Replace(p => ...)` inserts formatted text, hyperlinks, images or fields. Replacing a range moves the ranges after it, so search hits can be replaced in order. Edits that could damage the document throw `InvalidOperationException` and change nothing: field results, tracked changes, equations, drawings, footnote references, part of a content control, paragraph joins across a table or section break, and comments in headers or footers. Regex `ReplaceAll` checks every match before replacing any.

`Save()` writes to the path the document was created or loaded with; `Save(path)` writes a copy and keeps working on it. For files from outside your organization, load with `WordLoadOptions.UntrustedDefaults`, which rejects macros, embedded objects, ActiveX and external relationships before parsing.

For notes in `FootNotes` and `EndNotes`, `TextParagraphs` returns visible text per
top-level paragraph, retaining blank paragraphs, whitespace, tabs and cached
field results without field instructions. It returns a fresh text snapshot,
or `null` when the referenced ordinary note body cannot be resolved. `Paragraphs`
returns live run handles for formatting and editing: several handles can belong
to one paragraph, so its `Count` is not the note's paragraph count. Editing a
handle updates the note; `document.Inspect()` provides rich semantic snapshots.

Loading and conversion keep the source appearance. To restyle a document deliberately, apply a shared theme before saving or exporting:

```csharp
using OfficeIMO;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("Notes.docx");
document.ApplyTheme(OfficeTheme.Boardroom);
document.Save("Notes-styled.docx");
document.SaveAsPdf("Notes-styled.pdf");
```

`ApplyTheme` changes typography, paragraph rhythm and table decoration on the loaded model. It preserves content, semantic emphasis, code and symbol fonts, section geometry, table widths and merged cells. Table headers follow declared header rows. Word uses light-page theme defaults, including for `Dark`: an existing solid page background becomes white, while background images remain attached. Changed fonts, sizes and padding can change wrapping and pagination. The same themed model can use the existing PDF, HTML and image exporters.

## Styles and theme

`document.Styles` lists, creates, changes and removes the document's own paragraph, character and table styles. Changes edit the style definition in place, so everything that uses the style follows:

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Create("Report.docx");
document.Theme.SetColor(WordThemeColor.Accent1, OfficeColor.Parse("#0B5394"));
document.Theme.MajorFont = "Georgia";

WordStyle heading = document.Styles.GetOrAdd(WordParagraphStyles.Heading1);
heading.ThemeColor = WordThemeColor.Accent1;   // follows later theme changes
heading.ThemeFont = WordThemeFont.Major;

WordStyle note = document.Styles.AddParagraphStyle("Note", "Note");
note.IndentationBefore = OfficeLength.Centimeters(1);
note.ShadingFillColor = OfficeColor.Parse("#FFF2CC");
WordStyle term = document.Styles.AddCharacterStyle("Term", "Key Term");
term.Bold = true;

document.AddParagraph("Results").SetStyle(heading);
document.AddParagraph("Unaudited. ").SetStyle(note).AddText("Draft").SetCharacterStyle(term);
document.Save();
```

New styles are based on the default style of their type; set `BasedOnStyleId`, `NextStyleId` and `LinkedStyleId` to build chains. Formatting properties return what the style itself declares, and `null` inherits. `AddTableStyle` creates table styles with `SetTableBorders` and per-region formatting from `GetTableRegion` (header row, banding and so on). `Remove` falls content back to the default style and rebases dependent styles.

`Styles.Import(source, conflictMode)` copies styles from another loaded document or `.dotx` template, together with the styles they depend on and the lists they number with. A style matches by type and name, then by identifier, so localized built-in styles map onto their counterparts. `WordStyleConflictMode` keeps the target's style, overwrites it or imports a renamed copy; the returned `WordStyleImportReport` lists each outcome. `document.Theme.Import(source)` copies a document's theme. Text, borders and shading that select a theme colour or theme font follow theme changes in Word, PDF and HTML output, and the RGB and font names Word caches beside them are updated too. Text that names a font or colour directly, such as the Aptos body font of a new document, keeps it.

## Merge and split documents

`AppendDocument` adds another document's content at the end, with the styles, lists, images, hyperlinks, headers, footers, footnotes, endnotes and comments it uses. List, note, comment, bookmark, drawing and content-control identifiers are renumbered so nothing collides with the destination, and a list never continues one of the destination's lists. The appended content starts in its own section with its own page setup and headers.

```csharp
using var pack = WordDocument.Load("Cover.docx");
using var chapter = WordDocument.Load("Chapter.docx");
pack.AppendDocument(chapter, new WordDocumentMergeOptions {
    StyleConflictMode = WordStyleConflictMode.Rename   // the default: appended text keeps its look
});
pack.Save("Pack.docx");
```

When a source style has the same name as a destination style but a different definition, `WordStyleConflictMode.Rename` copies it as `Heading1_0` / `heading 1_0`, `KeepTarget` formats the appended text with the destination's style, and `Overwrite` replaces the destination's definition. `WordMergeSectionMode.UseDestinationSections` continues the appended content in the destination's last section instead.

`SplitBySection` and `SplitByHeading` cut a document into new documents. Each piece keeps the page setup and headers in effect where its content sits, and only the styles, lists, images, notes and comments it uses:

```csharp
using var report = WordDocument.Load("Report.docx");
IReadOnlyList<WordDocument> chapters = report.SplitByHeading(level: 1);
for (int i = 0; i < chapters.Count; i++) {
    using WordDocument chapter = chapters[i];
    chapter.Save($"Chapter{i + 1}.docx");
}
```

## Before sharing

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Load("Proposal.docx");
OfficeHiddenDataReport removed = document.RemoveHiddenData(new OfficeHiddenDataOptions {
    TrackedChanges = OfficeTrackedChangesAction.AcceptAll
});
WordAccessibilityReport accessibility = document.FixAccessibility(new WordAccessibilityFixOptions {
    Title = "Proposal", Language = "en-US", MarkTableHeaderRows = true
});
foreach (WordAccessibilityFinding finding in accessibility.Findings) Console.WriteLine($"{finding.Severity} {finding.Code} at {finding.Location}");
document.Save("Proposal-shared.docx");
```

`InspectHiddenData()` reports properties, comments, tracked changes, reviewer names, hidden text, variables, custom XML, image GPS and other metadata without changing anything; `RemoveHiddenData` removes the selected categories and reports what happened to each item. Hidden text and tracked changes are only changed when you ask. `InspectAccessibility()` checks alternative text, headings, tables, links, contrast, floating objects, blank-paragraph spacing, title and language. The [hidden data matrix](../Docs/officeimo.hidden-data-support-matrix.md) and the [Word guide](../Docs/officeimo.word-guide.md#accessibility-checks) list every rule.

## Convert

Conversions are called on the document and return an `OfficeConversionReport` that lists anything the target could not carry:

```csharp
using OfficeIMO;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;

using var document = WordDocument.Load("Status.docx");

OfficeConversionReport pdf = document.SaveAsPdf("Status.pdf");
if (pdf.HasLoss) {
    foreach (OfficeDiagnostic diagnostic in pdf.Diagnostics) Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}

document.SaveAsHtml("Status.html");
MarkdownDocument markdown = document.ToMarkdown();
File.WriteAllText("Status.md", markdown.ToMarkdown());

using WordDocument fromMarkdown = MarkdownDocument.Parse("# Notes\n\n- First\n- Second").ToWord();
fromMarkdown.Save("Notes.docx");
```

## More

- [Word guide](../Docs/officeimo.word-guide.md): lists, styles, sections and columns, fields and tables of contents, comments and revisions, content controls, mail merge, protection and signatures, equations, legacy `.doc`, image export and the full chart API.
- [Compatibility matrix](COMPATIBILITY.md) and the [generated operation contract](../Docs/Compatibility/generated/package-operations.md): exactly what is supported, preserved or rejected.
- [Word and HTML support](../Docs/officeimo.word-html-support-matrix.md), [Word advanced contracts](../Docs/officeimo.word-advanced-contracts.md), [templates and mail merge](../Docs/officeimo.word-template-mail-merge-scenarios.md) and [legacy DOC compatibility](../Docs/officeimo.word.legacy-doc-compatibility.md).
- [Examples](../OfficeIMO.Examples/Word) and the [migration guide](../MIGRATION.md).
- Adapters: [OfficeIMO.Word.Pdf](../OfficeIMO.Word.Pdf/README.md), [Markdown](../OfficeIMO.Word.Markdown/README.md), [Html](../OfficeIMO.Word.Html/README.md), [Rtf](../OfficeIMO.Word.Rtf/README.md), [OpenDocument](../OfficeIMO.Word.OpenDocument/README.md), [GoogleDocs](../OfficeIMO.Word.GoogleDocs/README.md) and [IWork](../OfficeIMO.Word.IWork/README.md). PDF back to Word lives in [OfficeIMO.Pdf.Word](../OfficeIMO.Pdf.Word/README.md).
- PowerShell: [PSWriteOffice](../PowerShell/PSWriteOffice/README.MD) (`New-OfficeWord`).

Dependencies: the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) and `OfficeIMO.Core`. Install `OfficeIMO.Security` only to create or validate document and VBA signatures. MIT licensed.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 14 | 12 | 0 | 8 | 2 | 2 |
| Read | 24 | 0 | 9 | 0 | 2 | 3 |
| Edit | 12 | 0 | 22 | 3 | 0 | 1 |
| Preserve | 11 | 0 | 22 | 0 | 0 | 1 |
| Inspect | 5 | 1 | 0 | 0 | 0 | 0 |
| Validate | 4 | 0 | 0 | 0 | 0 | 2 |
| Remove | 4 | 0 | 0 | 0 | 1 | 0 |
| Convert | 31 | 27 | 0 | 8 | 0 | 1 |

The complete rows for `OfficeIMO.Word` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
