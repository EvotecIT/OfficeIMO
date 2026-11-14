---
title: Create and Edit EPUB Files
description: "Create EPUB 3 or EPUB 2 e-books in .NET with chapters, styles, images and a table of contents, import an HTML manuscript, and run a preflight check."
order: 42
search_terms:
  - create epub
  - make ebook
  - html to epub
  - edit epub
  - epub writer
packages:
  - OfficeIMO.Epub
cmdlets:
  - New-OfficeEpub
  - Add-OfficeEpubChapter
  - Add-OfficeEpubResource
  - Set-OfficeEpubMetadata
  - Open-OfficeEpub
  - Save-OfficeEpub
limits: Preflight runs targeted native checks only; full EPUB schema conformance and accessibility review still need an independent validator and reading-system testing.
verified: 2026-11-14
---

`EpubPublication` holds a complete EPUB package: metadata, manifest, spine, navigation and every file in the ZIP. Create one from scratch or load an existing book, change it, and `Save` it. Loaded books keep entries OfficeIMO does not model, and an unedited save writes the original package back unchanged. `Save` returns an `EpubWriteReport`; `ThrowIfLoss()` stops on anything that could not be written as requested.

## Create an EPUB with chapters and a table of contents

`AddChapter` takes a well-formed XHTML body fragment. Stylesheet arguments are manifest ids, and navigation targets are container paths with optional fragments.

### C#

```csharp
using OfficeIMO.Epub;

EpubPublication publication = EpubPublication.Create("Operations handbook", "en");
publication.Creator = "Platform team";

publication.AddStylesheet("style", "EPUB/style.css",
    "body { font-family: sans-serif; } p { line-height: 1.4; }");

publication.AddChapter("opening", "EPUB/opening.xhtml", "Opening",
    "<h1 id='opening'>Opening</h1><p>Welcome to the <em>operations</em> handbook.</p>",
    new[] { "style" });
publication.AddChapter("on-call", "EPUB/on-call.xhtml", "On call",
    "<h1 id='on-call'>On call</h1><p>Escalate after fifteen minutes.</p>",
    new[] { "style" });

publication.SetNavigation(new[] {
    new EpubNavigationEntry("Opening", "EPUB/opening.xhtml#opening"),
    new EpubNavigationEntry("On call", "EPUB/on-call.xhtml#on-call")
});

publication.AddMetadataProperty("schema:accessMode", "textual");
publication.AddMetadataProperty("schema:accessibilityFeature", "structuralNavigation");

EpubWriteReport report = publication.Save("handbook.epub");
report.ThrowIfLoss();
```

To create EPUB 2.0.1 with NCX navigation, pass `version: EpubVersion.Epub2` to `Create`.

### PowerShell

```powershell
New-OfficeEpub -Path .\Operations.epub -Title 'Operations handbook' -Author 'Platform team' {
    Add-OfficeEpubChapter -Title 'Opening' -Text 'Welcome to the operations handbook.'
    Add-OfficeEpubChapter -Title 'On call' -Markdown "Escalate after **fifteen** minutes."
}
```

## Add a cover image

`AddResource` copies image, font or media bytes into the package. `SetCoverImage` marks an existing image as the cover.

### C#

```csharp
using OfficeIMO.Epub;

byte[] cover = File.ReadAllBytes("cover.png");

EpubPublication publication = EpubPublication.Create("Guide", "en");
publication.AddResource("cover", "EPUB/images/cover.png", "image/png", cover);
publication.SetCoverImage("cover");
publication.AddChapter("intro", "EPUB/intro.xhtml", "Introduction",
    "<h1>Introduction</h1><p><img src='images/cover.png' alt='Guide cover'/></p>");
publication.SetNavigation(new[] { new EpubNavigationEntry("Introduction", "EPUB/intro.xhtml") });

publication.Save("guide.epub").ThrowIfLoss();
```

### PowerShell

```powershell
$book = New-OfficeEpub -Path .\Guide.epub -Title 'Guide'
$book | Add-OfficeEpubResource -Path .\cover.png -Id cover -MediaType image/png
$book | Add-OfficeEpubChapter -Title 'Introduction' -Text 'Start here.'
Save-OfficeEpub -Document $book
```

## Turn an HTML manuscript into an EPUB

`EpubManuscript.ImportHtml` splits chapters at level-1 headings, keeps nested headings in the navigation and rewrites links between chapters. Scripts and event attributes are removed and reported.

### C#

```csharp
using OfficeIMO.Epub;
using OfficeIMO.Html;

HtmlConversionDocument source = HtmlConversionDocument.Parse(
    "<title>A short book</title>" +
    "<h1 id='opening'>Opening</h1><p>First chapter.</p>" +
    "<h1 id='closing'>Closing</h1><p><a href='#opening'>Back to the start</a></p>");

EpubManuscriptResult imported = EpubManuscript.ImportHtml(source,
    new EpubManuscriptOptions { Language = "en", Creator = "A writer" });

foreach (var diagnostic in imported.Report.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}

imported.ThrowIfLoss();
imported.Publication.Save("short-book.epub");
```

Set `ChapterHeadingLevel` to `0` for a single chapter, or 1–6 to split at another heading level.

## Edit an existing EPUB

Change metadata through the typed properties, and chapter content through `GetContentXml` and `SetContentXml`. Unchanged entries keep their bytes.

### C#

```csharp
using System.Xml.Linq;
using OfficeIMO.Epub;

EpubPublication publication = EpubPublication.Load("handbook.epub");
publication.Title = "Operations handbook, 2nd edition";

string chapterId = publication.Spine[0].ManifestId;
XDocument chapter = publication.GetContentXml(chapterId);
XNamespace xhtml = "http://www.w3.org/1999/xhtml";
chapter.Root!.Element(xhtml + "body")!.Add(new XElement(xhtml + "p", "Updated for the new rota."));
publication.SetContentXml(chapterId, chapter);

publication.Save("handbook-2nd.epub").ThrowIfLoss();
```

### PowerShell

```powershell
Open-OfficeEpub -Path .\Guide.epub {
    Set-OfficeEpubMetadata -Title 'Operations handbook, 2nd edition'
    Add-OfficeEpubChapter -Title 'Review' -Text 'Reviewed and approved.'
}
```

## Check a book before publishing

`Preflight` runs native checks (identifiers, links, image alternatives, language tags, accessibility metadata, media overlays) without writing anything.

### C#

```csharp
using OfficeIMO.Epub;

EpubPublication publication = EpubPublication.Load("handbook.epub");
EpubPreflightReport review = publication.Preflight();

foreach (EpubPreflightCheck check in review.Checks) {
    Console.WriteLine($"{check.Code}: {check.Status}");
    foreach (EpubDiagnostic finding in check.Diagnostics) {
        Console.WriteLine($"  {finding.Path}: {finding.Code}: {finding.Message}");
    }
}

Console.WriteLine($"Errors: {review.HasErrors}");
```
