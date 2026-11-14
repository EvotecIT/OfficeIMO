---
title: Convert RTF to Word, PDF, Markdown and HTML
description: Convert RTF files to DOCX, PDF, Markdown or HTML, and convert Word, Markdown or PDF back to RTF, with a report of anything that was flattened or omitted.
order: 42
search_terms:
  - rtf to docx
  - rtf to word
  - rtf to pdf
  - rtf to markdown
  - rtf to html
  - docx to rtf
packages:
  - OfficeIMO.Rtf
  - OfficeIMO.Word.Rtf
  - OfficeIMO.Rtf.Pdf
  - OfficeIMO.Rtf.Markdown
  - OfficeIMO.Html.Rtf
cmdlets:
  - Convert-OfficeDocument
  - New-OfficeRtfPdfOptions
commands:
  - officeimo convert
related:
  - /docs/rtf/
  - /docs/rtf/read/
  - /docs/html/render-and-convert/
limits: Paragraph frames become ordinary text flow in Word and PDF, and Markdown output omits headers, footers and annotations; each case is reported, and FailOnLoss turns it into an error.
verified: 2026-11-14
---

Each destination has its own small adapter package, and every adapter reports what it preserved, flattened, omitted or blocked. Set `FailOnLoss = true` on the options, or call `ThrowIfLoss()` on the report, when a conversion has to be exact. PDF export is a layout conversion; Word, Markdown and HTML are semantic conversions that keep the document editable.

| Destination | Package | Call |
|---|---|---|
| Word (.docx) | `OfficeIMO.Word.Rtf` | `ToWord`, `SaveAsWord` |
| PDF | `OfficeIMO.Rtf.Pdf` | `ToPdf`, `SaveAsPdf` |
| Markdown | `OfficeIMO.Rtf.Markdown` | `ToMarkdown`, `SaveAsMarkdown` |
| HTML | `OfficeIMO.Html.Rtf` | `ToHtml`, `SaveAsHtml` |

## Convert RTF to Word

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;

RtfDocument rtf = RtfDocument.Create();
rtf.AddParagraph().AddText("Service agreement").SetBold();
rtf.AddParagraph("Effective from January 1.");

using WordDocument word = rtf.ToWord(new RtfWordExportOptions(), out OfficeConversionReport report);
word.Save("agreement.docx");

foreach (OfficeDiagnostic diagnostic in report.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\agreement.rtf -To Word -OutputPath .\agreement.docx
```

### CLI

```sh
officeimo convert agreement.rtf agreement.docx
```

## Convert RTF to PDF

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;

RtfDocument rtf = RtfDocument.Create();
rtf.AddParagraph("Invoice 2026-041");
rtf.AddParagraph("Total due: 1,250.00");

OfficeConversionReport report = rtf.SaveAsPdf("invoice.pdf", new RtfPdfExportOptions());
Console.WriteLine($"Loss reported: {report.HasLoss}");
```

### PowerShell

```powershell
Convert-OfficeDocument -Path .\invoice.rtf -To Pdf -OutputPath .\invoice.pdf
```

### CLI

```sh
officeimo convert invoice.rtf invoice.pdf
```

## Convert RTF to Markdown

Footnotes and endnotes become Markdown footnotes. Supply an image exporter to write pictures next to the Markdown file.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;

RtfDocument rtf = RtfDocument.Create();
rtf.AddParagraph().AddText("Release notes").SetBold();
rtf.AddParagraph("Fixed the export timeout.");

var options = new RtfMarkdownExportOptions {
    ImagePathFactory = (_, index) => $"media/image-{index + 1}.png",
    ImageExporter = (image, _, path) => {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, image.Data);
    }
};

OfficeConversionReport report = rtf.SaveAsMarkdown("release-notes.md", options);
Console.WriteLine(File.ReadAllText("release-notes.md"));
```

### CLI

```sh
officeimo convert release-notes.rtf release-notes.md
```

## Convert RTF to HTML

The web-safe profile writes bounded semantic HTML; the print-review profile writes a complete styled page.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Rtf;

RtfDocument rtf = RtfDocument.Create();
rtf.AddParagraph("Hello from RTF");

string fragment = rtf.ToHtml(RtfHtmlExportOptions.CreateWebSafeProfile(), out OfficeConversionReport report);
string page = rtf.ToHtml(RtfHtmlExportOptions.CreatePrintReviewProfile(OfficeVisualThemeKind.WordLike));

File.WriteAllText("hello.html", page);
Console.WriteLine(fragment);
```

## Convert Word, Markdown or PDF to RTF

The same adapters work in the other direction.

### C#

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;

using (WordDocument word = WordDocument.Create("letter.docx")) {
    word.AddParagraph("Dear customer,");
    word.Save();
}

using (WordDocument word = WordDocument.Load("letter.docx")) {
    word.SaveAsRtf("letter.rtf");
}

MarkdownDocument markdown = MarkdownDocument.Parse("# Notes\n\n- First\n- Second\n");
RtfDocument fromMarkdown = markdown.ToRtf();
File.WriteAllText("notes.rtf", fromMarkdown.ToRtf());
```

PDF import (`pdf.ToRtf(...)` from `OfficeIMO.Rtf.Pdf`) recovers metadata, headings, paragraphs and list items as editable RTF. Tables, images and form fields are reported rather than reconstructed.
