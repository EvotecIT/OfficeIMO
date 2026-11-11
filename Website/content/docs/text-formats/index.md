---
title: "Text and Markup Formats"
description: "Read, edit and convert AsciiDoc, LaTeX, DocBook, CHM, OPML, DjVu, ADF, OfficeIMO Markup and bibliography data in .NET with loss reports."
order: 42
meta.seo_title: "AsciiDoc, LaTeX, DocBook and CHM formats in .NET"
---

These engines cover documentation, publishing, and reference formats that sit beside Word, Excel, and PowerPoint. Each package parses with explicit input limits, does not execute embedded code or fetch remote resources during parsing, and returns diagnostics when a conversion approximates or omits content.

| Format | Package(s) | Read | Write | Converts to |
|---|---|---|---|---|
| AsciiDoc | `OfficeIMO.AsciiDoc`, `.Markdown`, `.Pdf` | Yes | Yes, source-preserving | Markdown (both directions), PDF |
| LaTeX | `OfficeIMO.Latex`, `.Markdown`, `.Pdf` | Bounded LaTeX2e profile | Yes, source-preserving | Markdown (both directions), PDF |
| DocBook | `OfficeIMO.DocBook` | DocBook 4 and 5 | DocBook 4.5 and 5.2 | Shared OfficeIMO document model |
| Compiled HTML Help | `OfficeIMO.Chm`, `.Markdown`, `.Epub`, `.Pdf` | Yes | No | HTML, Markdown, EPUB, PDF |
| OPML | `OfficeIMO.Opml` | OPML 1.0, 1.1, 2.0 | OPML 1.0 and 2.0 | Shared OfficeIMO document model |
| DjVu | `OfficeIMO.DjVu`, `OfficeIMO.DjVu.Pdf` | Yes | No | Page rasters, searchable PDF |
| Atlassian Document Format | `OfficeIMO.Adf` | ADF JSON | ADF JSON | Markdown and HTML (both directions) |
| Confluence Cloud | `OfficeIMO.Confluence` | Pages and attachments | Pages and attachments | Markdown, HTML, ADF page bodies |
| OfficeIMO Markup | `OfficeIMO.Markup`, `.Word`, `.Excel`, `.PowerPoint` | `.omd` and Markdown-like text | Starter C# and PowerShell | Word, Excel, PowerPoint |
| Bibliography | `OfficeIMO.Bibliography` | BibTeX, BibLaTeX, CSL JSON, RIS, NBIB/MEDLINE, EndNote XML | Same formats | Between those formats; CSL citations as text or HTML |

## AsciiDoc

`OfficeIMO.AsciiDoc` parses AsciiDoc into a semantic model while keeping the original source tree. Unchanged input writes back character for character, and unsupported constructs stay in the source tree instead of disappearing. Includes, conditionals, and registered directives run only through the separate `AsciiDocProcessor`, and the built-in include resolver is confined to a root directory. `OfficeIMO.AsciiDoc.Markdown` converts to and from Markdown, and `OfficeIMO.AsciiDoc.Pdf` renders through the Markdown projection and the OfficeIMO PDF engine.

```shell
dotnet add package OfficeIMO.AsciiDoc
dotnet add package OfficeIMO.AsciiDoc.Markdown
dotnet add package OfficeIMO.AsciiDoc.Pdf
```

```csharp
using OfficeIMO;
using OfficeIMO.AsciiDoc;
using OfficeIMO.AsciiDoc.Markdown;
using OfficeIMO.AsciiDoc.Pdf;

AsciiDocDocument guide = AsciiDocDocument.Create()
    .AddHeading(1, "Getting started")
    .AddParagraph("Install the package and keep the source *exactly* as written.")
    .Add("* Parse\n* Edit\n* Write");
File.WriteAllText("guide.adoc", guide.ToAsciiDoc());

OfficeConversionReport markdown = guide.SaveAsMarkdown("guide.md");
OfficeConversionReport pdf = guide.SaveAsPdf("guide.pdf");
foreach (OfficeDiagnostic diagnostic in pdf.Diagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

Limits: the engine does not claim every Asciidoctor substitution, macro, extension, advanced table layout, or rendering behavior. File encoding and BOM bytes are not retained by `Load`/`Save`. See the [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.AsciiDoc/README.md) and the [AsciiDoc support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.asciidoc-support-matrix.md).

## LaTeX

`OfficeIMO.Latex` is a source-preserving LaTeX2e parser and writer for a bounded document profile: article, report, and book structure, lists, figures, tabular data, labels and references, citations, theorem-like environments, footnotes, and inline or display math. It is not a TeX compiler. It never executes commands, loads packages, or starts an external runtime. Unknown commands and environments remain in the source, and simple document-local macros expand only when you opt in.

```shell
dotnet add package OfficeIMO.Latex
dotnet add package OfficeIMO.Latex.Markdown
dotnet add package OfficeIMO.Latex.Pdf
```

```csharp
using OfficeIMO;
using OfficeIMO.Latex;
using OfficeIMO.Latex.Markdown;
using OfficeIMO.Latex.Pdf;

LatexDocument article = LatexDocument.Load("article.tex");
article.Headings[0].Title = "Updated section";
File.WriteAllText("article-edited.tex", article.ToLatex());

OfficeConversionReport markdown = article.SaveAsMarkdown("article.md");
OfficeConversionReport pdf = article.SaveAsPdf("article.pdf");
Console.WriteLine($"PDF conversion reported loss: {pdf.HasLoss}");
```

Limits: PDF output is a projection of the bounded profile through the Markdown and PDF engines, not TeX typesetting. Math layout, package-specific commands, bibliography formatting, and unknown environments are reported as source fallbacks or simplifications. See the [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Latex/README.md) and the [LaTeX support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.latex-support-matrix.md).

## DocBook

`OfficeIMO.DocBook` creates, reads, edits, validates, and writes a common-structure profile of DocBook 4 and 5: articles, books, metadata, sections, lists, CALS-shaped tables, code, links, cross-references, admonitions, media, and indexes. The typed API is a view over the underlying XML, so unknown elements, comments, and processing instructions are preserved, and an unchanged file is written byte for byte. External entity resolution is disabled.

```shell
dotnet add package OfficeIMO.DocBook
```

```csharp
using OfficeIMO.DocBook;

DocBookDocument document = DocBookDocument.CreateArticle(DocBookProfile.DocBook52);
document.Title = "Deployment guide";

DocBookNode section = document.AddSection("Install");
section.AddParagraph("Install the package from NuGet.");
section.AddProgramListing("dotnet add package OfficeIMO.DocBook", "shell");
section.AddAdmonition(DocBookNodeKind.Note, "Validate before publishing.");

DocBookValidationResult validation = document.Validate();
Console.WriteLine($"Valid: {validation.IsValid}");
document.Save("guide.docbook");
```

Limits: `Validate()` checks the OfficeIMO common-structure contract only. It does not run the official DTD, RELAX NG, or Schematron rules; the package exposes their identifiers so you can run them in your own validator. Shared-model pages, forms, and structured visuals have no equivalent in this profile and are reported as loss. See the [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.DocBook/README.md).

## Compiled HTML Help (CHM)

`OfficeIMO.Chm` reads `.chm` files in process on Windows, macOS, and Linux. It exposes topics, exact entry bytes, embedded resources, the contents tree, and the keyword index. Loading does not open the Windows help viewer, run scripts, or fetch external resources. The adapter packages convert a book to linked Markdown, a reflowable EPUB, or PDF, and each conversion returns a fidelity report.

```shell
dotnet add package OfficeIMO.Chm
dotnet add package OfficeIMO.Chm.Markdown
dotnet add package OfficeIMO.Chm.Epub
dotnet add package OfficeIMO.Chm.Pdf
```

```csharp
using OfficeIMO;
using OfficeIMO.Chm;

ChmDocument book = ChmDocument.Load("manual.chm");
Console.WriteLine(book.Title);
foreach (ChmTopic topic in book.Topics)
    Console.WriteLine($"{topic.Title}: {topic.Path}");

OfficeConversionReport markdown = book.SaveAsMarkdown("manual.md");
OfficeConversionReport epub = book.SaveAsEpub("manual.epub");
OfficeConversionReport pdf = book.SaveAsPdf("manual.pdf");
foreach (OfficeDiagnostic diagnostic in pdf.Diagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

Limits: there is no CHM writer. ActiveX, scripts, compiled full-text search, merged books, and help-viewer window behavior are outside the conversion contract. EPUB and PDF output do not rebuild the keyword index, and multi-topic PDFs are untagged. See the [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Chm/README.md) and [CHM support](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Chm/SUPPORT.md).

## OPML

`OfficeIMO.Opml` creates, reads, edits, validates, and writes OPML 1.0 and 2.0 without a third-party parser; OPML 1.1 is read with the 1.0 profile. Standard subscription attributes are typed, and unknown attributes and elements are preserved. Nested outlines map to and from the shared OfficeIMO document model with loss diagnostics.

```shell
dotnet add package OfficeIMO.Opml
```

```csharp
using OfficeIMO.Opml;

OpmlDocument document = OpmlDocument.Create(OpmlVersion.Opml20);
document.Head.Title = "Subscriptions";

OpmlOutline folder = document.AddOutline("Technology");
OpmlOutline feed = folder.AddChild("OfficeIMO");
feed.Type = "rss";
feed.XmlUrl = "https://example.com/feed.xml";
feed.HtmlUrl = "https://example.com/";

document.Validate();
document.Save("subscriptions.opml");
```

Limits: DTDs and external entities are disabled. Tables, assets, pages, and forms have no OPML outline equivalent and are omitted with a diagnostic when converting from the shared model. See the [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Opml/README.md).

## DjVu

`OfficeIMO.DjVu` reads `.djvu` and `.djv` scanned documents with managed codecs: no external decoder, native library, or executable is required. It exposes pages in source order, stored text with its status (absent, empty, present, or corrupt), outlines, and page rendering to RGBA rasters. `OfficeIMO.DjVu.Pdf` produces PDFs that keep the scanned pixels, the source page size, and a searchable layer from the stored text.

```shell
dotnet add package OfficeIMO.DjVu
dotnet add package OfficeIMO.DjVu.Pdf
```

```csharp
using OfficeIMO;
using OfficeIMO.DjVu;
using OfficeIMO.DjVu.Pdf;

DjVuDocument document = DjVuDocument.Load("book.djvu");
DjVuTextResult text = document.Pages[0].GetText();
Console.WriteLine($"Page 1 text: {text.Status}");

OfficeConversionReport report = document.SaveAsPdf("book.pdf", new DjVuPdfExportOptions {
    PageNumbers = new[] { 1, 2 }
});
```

Limits: DjVu authoring is not supported. JPEG 2000 layers fail explicitly, and the PDF keeps scanned pixels rather than rebuilding editable text or typography. OCR for pages without stored text runs only through the asynchronous PDF methods with an `IOcrEngine` you supply; see [Charts, OCR and provenance](/docs/foundations/). See the [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.DjVu/README.md) and [DjVu support](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.DjVu/SUPPORT.md).

## Atlassian Document Format and Confluence

`OfficeIMO.Adf` parses and writes Atlassian Document Format JSON and converts it to and from Markdown and HTML through the OfficeIMO Markdown engine. Unknown nodes, marks, and attributes survive JSON round trips, and projections report what they could not carry. Validation can check recognized structure or the complete embedded ADF schema, and an optional destination policy records the node and mark types your target accepts.

`OfficeIMO.Confluence` is a Confluence Cloud client built on the same conversions. It reads and writes pages and attachments, returns a reviewable plan for a create, update, or delete without executing it, and updates managed sections between stable markers. Credentials stay with the caller; the library does not store tokens.

```shell
dotnet add package OfficeIMO.Adf
dotnet add package OfficeIMO.Confluence
```

```csharp
using OfficeIMO.Adf;

AdfConversionResult<AdfDocument> generated = AdfConverter.FromMarkdown("## Status\n\nReady for **review**.");
string adfJson = generated.Value.ToJson(indented: true);

AdfDocument document = AdfDocument.Parse(adfJson);
AdfConversionResult<string> markdown = AdfConverter.ToMarkdown(document);
AdfConversionResult<string> html = AdfConverter.ToHtml(document);
Console.WriteLine($"Markdown projection reported loss: {markdown.Report.HasLoss}");
```

```csharp
using OfficeIMO.Confluence;

string email = Environment.GetEnvironmentVariable("CONFLUENCE_EMAIL")!;
string apiToken = Environment.GetEnvironmentVariable("CONFLUENCE_API_TOKEN")!;

var session = new ConfluenceSession(
    new ConfluenceBasicCredentialSource(email, apiToken),
    new ConfluenceSessionOptions { SiteUri = new Uri("https://example.atlassian.net/") });

using ConfluenceClient client = session.CreateClient();
ConfluencePage page = await client.GetPageAsync("123", ConfluenceBodyFormat.AtlasDocFormat);
ConfluenceContentConversionResult<string> pageMarkdown = ConfluenceContentConverter.ToMarkdown(page);
```

Limits: schema validity is not the same as acceptance by a specific Jira or Confluence API; product settings can add restrictions. Cards, media, and mentions are projected from local labels and metadata, never fetched. Creates, updates, and uploads are not retried automatically after an ambiguous write. See the [ADF README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Adf/README.md) and the [Confluence README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Confluence/README.md).

## OfficeIMO Markup

`OfficeIMO.Markup` parses `.omd` or Markdown-like text into a semantic document for one of three profiles: document, workbook, or presentation. It adds container directives such as `@slide`, `@section`, and `@sheet`, and Office-aware blocks such as `::notes`, `::chart`, `::range`, and `::formula`. You can validate the result, emit starter C# or PowerShell, or export editable Office files through `OfficeIMO.Markup.Word`, `OfficeIMO.Markup.Excel`, and `OfficeIMO.Markup.PowerPoint`. The `officeimo markup` command in `OfficeIMO.Tool` runs the same pipeline, and editor support for VS Code is in `OfficeIMO.Markup.VSCode`.

```shell
dotnet add package OfficeIMO.Markup
dotnet add package OfficeIMO.Markup.PowerPoint
```

`OfficeIMO.Markup.Word` and `OfficeIMO.Markup.Excel` are built from the OfficeIMO source tree and are not published as standalone NuGet packages; reference their projects from a checkout.

```csharp
using OfficeIMO;
using OfficeIMO.Markup;
using OfficeIMO.Markup.PowerPoint;

string markup = @"@slide title=""Quarterly review""

- Revenue grew
- Churn improved

::notes
Open with the top-line result.
";

OfficeMarkupParseResult parsed = OfficeMarkupParser.Parse(markup, new OfficeMarkupParserOptions {
    Profile = OfficeMarkupProfile.Presentation
});
foreach (OfficeMarkupDiagnostic diagnostic in parsed.Diagnostics)
    Console.WriteLine($"{diagnostic.Severity}: {diagnostic.Message}");

if (!parsed.HasErrors) {
    OfficeConversionReport report = parsed.Document.SaveAsPowerPoint("quarterly-review.pptx");
    string starterCode = new OfficeMarkupCSharpEmitter().Emit(parsed.Document);
}
```

Limits: Mermaid diagrams in PowerPoint export need a caller-installed Mermaid CLI, or the in-process renderer from `OfficeIMO.ChartForgeX.Markdown`. See the [Markup README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Markup/README.md) and the exporter READMEs for [Word](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Markup.Word/README.md), [Excel](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Markup.Excel/README.md), and [PowerPoint](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Markup.PowerPoint/README.md).

## Bibliography

`OfficeIMO.Bibliography` holds citation data in one editable model with codecs for BibTeX, BibLaTeX, CSL JSON, RIS, PubMed NBIB/MEDLINE, and EndNote XML. An unchanged document writes back its original bytes; edited documents are written in deterministic canonical syntax. Every write returns a conversion report that marks fields as approximated or omitted, and `FailOnLoss` rejects lossy output. A managed CSL processor renders citations and bibliographies as text or HTML from styles you supply.

```shell
dotnet add package OfficeIMO.Bibliography
```

```csharp
using OfficeIMO.Bibliography;

BibliographyReadResult read = BibliographyDocument.Load("library.bib", BibliographyFormat.BibLatex);
BibliographyItem item = read.Document.Items[0];
item.Title = "A corrected title";
item.SetIdentifier("DOI", "10.1000/example");

BibliographyWriteResult csl = read.Document.Save("library.json", new BibliographyWriteOptions {
    Format = BibliographyFormat.CslJson,
    Mode = BibliographyWriterMode.Canonical
});
Console.WriteLine($"CSL JSON conversion reported loss: {csl.Report.HasLoss}");
```

Limits: the package does not run TeX, fetch DOI or PubMed metadata, or resolve remote resources. The CSL renderer implements a qualified subset of CSL 1.0.2, and style resolution performs no network or file access. See the [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Bibliography/README.md) and the [bibliography support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.bibliography-support-matrix.md).
