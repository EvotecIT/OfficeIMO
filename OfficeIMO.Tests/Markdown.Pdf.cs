using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class MarkdownPdfTests {
    [Fact]
    public void Markdown_SaveAsPdf_ExportsCoreDocumentStructure() {
        var options = new MarkdownPdfSaveOptions();
        string markdown = """
# Release Notes

This is **important**, _portable_, and links to [OfficeIMO](https://github.com/EvotecIT/OfficeIMO).

## Details

- [x] Native PDF engine
- Table and code support

| Area | State |
| --- | --- |
| Markdown | PDF |
| Word | PDF |

```csharp
Console.WriteLine("OfficeIMO");
```
""";

        byte[] pdf = markdown.SaveAsPdf(options);

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.True(pdf.Length > 0);
        Assert.True(info.PageCount >= 1);
        Assert.Empty(options.Warnings);
        Assert.Contains("Release Notes", text);
        Assert.Contains("Native PDF engine", text);
        Assert.Contains("Markdown", text);
        Assert.Contains("Console.WriteLine", text);
        Assert.Contains(info.Outlines, outline => outline.Title == "Release Notes");
    }

    [Fact]
    public void Markdown_SaveAsPdf_RecordsWarningsForRemoteImages() {
        var options = new MarkdownPdfSaveOptions();
        string markdown = """
# Remote Asset

![OfficeIMO logo](https://example.com/logo.png)
""";

        byte[] pdf = markdown.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.True(pdf.Length > 0);
        MarkdownPdfExportWarning warning = Assert.Single(options.Warnings);
        Assert.Equal("UnsupportedImage", warning.Code);
        Assert.Contains("OfficeIMO logo", text);
    }

    [Fact]
    public void Markdown_SaveAsPdf_EmbedsRemoteImagesThroughExplicitResolver() {
        Uri? requestedUri = null;
        var options = new MarkdownPdfSaveOptions {
            RemoteImageResolver = uri => {
                requestedUri = uri;
                return Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=");
            }
        };
        string markdown = """
# Remote Asset

![OfficeIMO logo](https://example.com/logo.png){width=24 height=24}
""";

        byte[] pdf = markdown.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
        IReadOnlyList<PdfCore.PdfExtractedImage> images = PdfCore.PdfImageExtractor.ExtractImages(pdf);

        Assert.Equal(new Uri("https://example.com/logo.png"), requestedUri);
        Assert.Empty(options.Warnings);
        Assert.DoesNotContain("[Image:", text, StringComparison.Ordinal);
        Assert.Single(images);
        Assert.Equal(1, images[0].Width);
        Assert.Equal(1, images[0].Height);
    }

    [Fact]
    public void Markdown_SaveAsPdf_WarnsWhenResolvedRemoteImageExceedsLimit() {
        var options = new MarkdownPdfSaveOptions {
            MaximumRemoteImageBytes = 3,
            RemoteImageResolver = _ => new byte[] { 1, 2, 3, 4 }
        };
        string markdown = """
# Remote Asset

![OfficeIMO logo](https://example.com/logo.png)
""";

        byte[] pdf = markdown.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        MarkdownPdfExportWarning warning = Assert.Single(options.Warnings);
        Assert.Equal("ImageTooLarge", warning.Code);
        Assert.Contains("OfficeIMO logo", text);
        Assert.Contains("[Image:", text, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownPdfConverter_SaveFileAsPdf_ResolvesRelativeLocalImages() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Markdown.Pdf", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string markdownPath = Path.Combine(directory, "README.md");
            string pdfPath = Path.Combine(directory, "README.pdf");
            string imagePath = Path.Combine(directory, "pixel.png");

            File.WriteAllBytes(imagePath, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="));
            File.WriteAllText(markdownPath, """
# Asset Report

![Local pixel](pixel.png){width=32 height=32}
_Figure 1. Embedded from a relative Markdown path._
""");

            var options = new MarkdownPdfSaveOptions();
            MarkdownPdfConverter.SaveFileAsPdf(markdownPath, pdfPath, options);

            byte[] pdf = File.ReadAllBytes(pdfPath);
            string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
            IReadOnlyList<PdfCore.PdfExtractedImage> images = PdfCore.PdfImageExtractor.ExtractImages(pdf);

            Assert.True(pdf.Length > 0);
            Assert.Empty(options.Warnings);
            Assert.Null(options.BaseDirectory);
            Assert.Contains("Asset Report", text);
            Assert.Contains("Figure 1", text);
            Assert.Single(images);
            Assert.Equal(1, images[0].Width);
            Assert.Equal(1, images[0].Height);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void Markdown_TrySaveAsPdf_ReturnsCoreSaveResult() {
        string markdown = "# Result Adapter\n\nPDF output should report bytes and diagnostics.";
        using var stream = new MemoryStream();

        PdfCore.PdfSaveResult streamResult = markdown.TrySaveAsPdf(stream);

        Assert.True(streamResult.Succeeded);
        Assert.Null(streamResult.OutputPath);
        Assert.True(streamResult.BytesWritten > 0);
        Assert.Empty(streamResult.Diagnostics);
        Assert.Equal(streamResult.BytesWritten, stream.ToArray().LongLength);

        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Markdown.Pdf.Result", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string pdfPath = Path.Combine(directory, "result.pdf");
        try {
            PdfCore.PdfSaveResult pathResult = markdown.TrySaveAsPdf(pdfPath);

            Assert.True(pathResult.Succeeded);
            Assert.Equal(Path.GetFullPath(pdfPath), pathResult.OutputPath);
            Assert.Equal(File.ReadAllBytes(pdfPath).LongLength, pathResult.BytesWritten);

            PdfCore.PdfSaveResult directoryResult = markdown.TrySaveAsPdf(directory);

            Assert.False(directoryResult.Succeeded);
            Assert.NotEmpty(directoryResult.Diagnostics);
            Assert.Throws<InvalidOperationException>(() => directoryResult.RequireSuccess());
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }

        PdfCore.PdfSaveResult conversionFailure = ((string)null!).TrySaveAsPdf(new MemoryStream());

        Assert.False(conversionFailure.Succeeded);
        Assert.NotEmpty(conversionFailure.Diagnostics);
    }

    [Fact]
    public void Markdown_SaveAsPdf_EmbedsDataUriImages() {
        var options = new MarkdownPdfSaveOptions();
        string markdown = """
# Inline Asset

![Inline pixel](data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=){width=24 height=24}
""";

        byte[] pdf = markdown.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
        IReadOnlyList<PdfCore.PdfExtractedImage> images = PdfCore.PdfImageExtractor.ExtractImages(pdf);

        Assert.True(pdf.Length > 0);
        Assert.Empty(options.Warnings);
        Assert.Contains("Inline Asset", text);
        Assert.DoesNotContain("[Image:", text, StringComparison.Ordinal);
        Assert.Single(images);
        Assert.Equal(1, images[0].Width);
        Assert.Equal(1, images[0].Height);
    }

    [Fact]
    public void Markdown_SaveAsPdf_RendersTaskListsAsCheckboxes() {
        var options = new MarkdownPdfSaveOptions {
            VisualTheme = MarkdownPdfVisualTheme.GitHubLike()
        };
        string markdown = """
# Checklist

- [x] Ship first-party Markdown PDF
- [ ] Add multi-block panel primitive
""";

        byte[] pdf = markdown.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.True(pdf.Length > 0);
        Assert.Empty(options.Warnings);
        Assert.Contains("Ship first-party Markdown PDF", text);
        Assert.Contains("Add multi-block panel primitive", text);
        Assert.DoesNotContain("[x]", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("[ ]", text, StringComparison.Ordinal);
        Assert.False(info.HasForms);
        Assert.DoesNotContain("/FT /Btn", rawPdf, StringComparison.Ordinal);
        Assert.Contains(" S\n", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_AppliesChecklistThemeColors() {
        MarkdownPdfVisualTheme theme = MarkdownPdfVisualTheme.Plain();
        theme.ChecklistCheckedIconColor = PdfCore.PdfColor.FromRgb(255, 0, 0);
        theme.ChecklistUncheckedIconColor = PdfCore.PdfColor.FromRgb(0, 0, 255);
        theme.ChecklistCheckedTextColor = PdfCore.PdfColor.FromRgb(0, 128, 0);
        theme.ChecklistUncheckedTextColor = PdfCore.PdfColor.FromRgb(255, 0, 255);
        theme.ChecklistCheckedFillColor = PdfCore.PdfColor.FromRgb(255, 255, 204);
        theme.ChecklistUncheckedFillColor = PdfCore.PdfColor.FromRgb(204, 238, 255);
        var options = new MarkdownPdfSaveOptions {
            VisualTheme = theme
        };
        string markdown = """
# Checklist Theme

- [x] Completed item
- [ ] Open item
""";

        byte[] pdf = markdown.SaveAsPdf(options);
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.Empty(options.Warnings);
        Assert.Contains("1 0 0 RG", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0 0 1 RG", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0 0.502 0 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("1 0 1 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("1 1 0.8 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.8 0.933 1 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_AppliesVisualThemeLinkStyleAcrossInlineSurfaces() {
        MarkdownPdfVisualTheme theme = MarkdownPdfVisualTheme.TechnicalDocument();
        theme.LinkColor = PdfCore.PdfColor.FromRgb(128, 0, 128);
        theme.UnderlineLinks = false;
        var options = new MarkdownPdfSaveOptions {
            VisualTheme = theme
        };
        string markdown = """
# Link Theme

Paragraph [paragraph link](https://example.com/paragraph).

- [x] [task link](https://example.com/task)

| Surface | Link |
| --- | --- |
| Table | [table link](https://example.com/table) |
""";

        byte[] pdf = markdown.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf);

        Assert.Empty(options.Warnings);
        Assert.Contains("paragraph link", text);
        Assert.Contains("task link", text);
        Assert.Contains("table link", text);
        Assert.Equal("paragraph link", Assert.Single(logical.GetLinksByUri("https://example.com/paragraph")).Contents);
        Assert.Equal("task link", Assert.Single(logical.GetLinksByUri("https://example.com/task")).Contents);
        Assert.Equal("table link", Assert.Single(logical.GetLinksByUri("https://example.com/table")).Contents);
        Assert.True(CountOccurrences(rawPdf, "0.502 0 0.502 rg") >= 3, "Expected the custom theme link fill color on paragraph, checklist, and table links.");
        Assert.DoesNotContain("0.502 0 0.502 RG", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_UsesFrontMatterAndHeadingMetadata() {
        string markdown = """
---
title: PDF Roadmap
author: OfficeIMO
tags: [pdf, markdown, native]
description: Dependency-free export
---
# Visible Heading

Content.
""";

        byte[] pdf = markdown.SaveAsPdf();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);

        Assert.Equal("PDF Roadmap", info.Metadata.Title);
        Assert.Equal("OfficeIMO", info.Metadata.Author);
        Assert.Equal("Dependency-free export", info.Metadata.Subject);
        Assert.Equal("pdf, markdown, native", info.Metadata.Keywords);
    }

    [Fact]
    public void Markdown_SaveAsPdf_RendersFrontMatterAsDocumentHeader() {
        string markdown = """
---
title: PDF Roadmap
subtitle: Native Markdown export
author: OfficeIMO
date: 2026-06-01
tags: [pdf, markdown]
---
# PDF Roadmap

Content.
""";

        var options = new MarkdownPdfSaveOptions {
            VisualTheme = MarkdownPdfVisualTheme.TechnicalDocument()
        };
        byte[] pdf = markdown.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.Empty(options.Warnings);
        Assert.Contains("PDF Roadmap", text);
        Assert.Contains("Native Markdown export", text);
        Assert.Contains("OfficeIMO", text);
        Assert.Contains("2026-06-01", text);
        Assert.Contains("Tags: pdf, markdown", text);
        Assert.Equal(1, CountOccurrences(text, "PDF Roadmap"));
        Assert.DoesNotContain("Key", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Value", text, StringComparison.Ordinal);
        Assert.Contains("0.059 0.09 0.165 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_CanRenderFrontMatterAsTable() {
        string markdown = """
---
title: PDF Roadmap
author: OfficeIMO
---
# Visible Heading
""";

        var options = new MarkdownPdfSaveOptions {
            FrontMatterRenderMode = MarkdownPdfFrontMatterRenderMode.Table
        };
        byte[] pdf = markdown.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.Empty(options.Warnings);
        Assert.Contains("Key", text);
        Assert.Contains("Value", text);
        Assert.Contains("title", text);
        Assert.Contains("PDF Roadmap", text);
        Assert.Contains("Visible Heading", text);
    }

    [Fact]
    public void Markdown_SaveAsPdf_RendersSemanticMarkdownBlocks() {
        var toc = new TocBlock {
            Ordered = true
        };
        toc.Entries.Add(new TocBlock.Entry {
            Level = 1,
            Text = "PDF Playbook",
            Anchor = "pdf-playbook"
        });

        MarkdownDoc document = MarkdownDoc.Create()
            .Add(toc)
            .H1("PDF Playbook")
            .Callout("warning", "Deployment note", "Keep backup enabled.")
            .Details("More detail", body => body.P("Hidden content."), open: true)
            .Dl(list => list.Item("Term", "Definition value"))
            .Add(new SemanticFencedBlock("diagram", "mermaid", "graph TD\nA-->B", "Flow caption"))
            .Add(new FootnoteDefinitionBlock("audit", "Footnote audit trail."));

        var options = new MarkdownPdfSaveOptions();
        byte[] pdf = document.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.True(pdf.Length > 0);
        Assert.Empty(options.Warnings);
        Assert.Contains("PDF Playbook", text);
        Assert.Contains("Deployment note", text);
        Assert.Contains("Keep backup enabled", text);
        Assert.Contains("More detail", text);
        Assert.Contains("Hidden content", text);
        Assert.Contains("Term", text);
        Assert.Contains("Definition value", text);
        Assert.Contains("diagram", text);
        Assert.Contains("mermaid", text);
        Assert.Contains("graph", text);
        Assert.Contains("Flow caption", text);
        Assert.Contains("audit", text);
        Assert.Contains("Footnote audit trail", text);
    }

    [Fact]
    public void Markdown_SaveAsPdf_RendersFluentTocAsLinkedThemedPanel() {
        MarkdownDoc document = MarkdownDoc.Create()
            .Toc(options => {
                options.Title = "Contents";
                options.Layout = TocLayout.Panel;
                options.MinLevel = 1;
                options.MaxLevel = 2;
            }, placeAtTop: true)
            .H1("PDF Playbook")
            .P("Introductory copy.")
            .H2("Install")
            .P("Installation notes.")
            .H2("Validate")
            .P("Validation notes.");

        var options = new MarkdownPdfSaveOptions {
            VisualTheme = MarkdownPdfVisualTheme.TechnicalDocument()
        };

        byte[] pdf = document.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf);

        Assert.Empty(options.Warnings);
        Assert.Contains("Contents", text);
        Assert.Contains("PDF Playbook", text);
        Assert.Contains("Install", text);
        Assert.Contains("Validate", text);
        Assert.Contains(logical.GetLinksByDestinationName("pdf-playbook"), link => link.Contents == "Table of contents: PDF Playbook");
        Assert.Contains(logical.GetLinksByDestinationName("install"), link => link.Contents == "Table of contents: Install");
        Assert.Contains(logical.GetLinksByDestinationName("validate"), link => link.Contents == "Table of contents: Validate");
    }

    [Fact]
    public void Markdown_SaveAsPdf_RespectsParsedTocRequireTopLevelFalse() {
        string markdown = """
# PDF Playbook

[TOC min=2 max=2 layout=panel title="Contents" requiretoplevel=false]

## Install

Installation notes.

## Validate

Validation notes.
""";

        var options = new MarkdownPdfSaveOptions {
            VisualTheme = MarkdownPdfVisualTheme.TechnicalDocument()
        };

        byte[] pdf = markdown.SaveAsPdf(options);
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf);

        Assert.Empty(options.Warnings);
        Assert.Empty(logical.GetLinksByDestinationName("pdf-playbook"));
        Assert.Contains(logical.GetLinksByDestinationName("install"), link => link.Contents == "Table of contents: Install");
        Assert.Contains(logical.GetLinksByDestinationName("validate"), link => link.Contents == "Table of contents: Validate");
    }

    [Fact]
    public void Markdown_SaveAsPdf_AppliesBuiltInVisualTheme() {
        string markdown = """
# Styled Document

> [!TIP] Better PDFs
> Theme-aware callout.

| Area | State |
| --- | --- |
| Visuals | Technical |
""";

        var options = new MarkdownPdfSaveOptions {
            VisualTheme = MarkdownPdfVisualTheme.TechnicalDocument()
        };

        byte[] pdf = markdown.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.Empty(options.Warnings);
        Assert.Contains("Better PDFs", text);
        Assert.Contains("Visuals", text);
        Assert.Contains("0.059 0.09 0.165 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownPdfVisualThemes_UseDocumentRhythmForTables() {
        MarkdownPdfVisualTheme theme = MarkdownPdfVisualTheme.TechnicalDocument();

        PdfCore.PdfTableStyle tableStyle = theme.TableStyle!;
        PdfCore.PdfTableStyle frontMatterStyle = theme.FrontMatterTableStyle!;

        Assert.True(tableStyle.AutoFitColumns);
        Assert.True(tableStyle.SpacingAfter >= 8);
        Assert.True(tableStyle.LineHeight >= 1.18);
        Assert.True(frontMatterStyle.AutoFitColumns);
        Assert.True(frontMatterStyle.SpacingAfter >= 8);
        Assert.Equal(PdfCore.PdfCellVerticalAlign.Top, theme.ChecklistTableStyle!.VerticalAlignments![0]);
        Assert.Equal(PdfCore.PdfCellVerticalAlign.Top, theme.ChecklistTableStyle.VerticalAlignments![1]);
    }

    [Fact]
    public void MarkdownPdfVisualThemes_CloneLinkStyleOptions() {
        MarkdownPdfVisualTheme theme = MarkdownPdfVisualTheme.Report();
        theme.LinkColor = PdfCore.PdfColor.FromRgb(12, 34, 56);
        theme.UnderlineLinks = false;

        MarkdownPdfVisualTheme clone = theme.Clone();
        theme.LinkColor = PdfCore.PdfColor.FromRgb(90, 90, 90);
        theme.UnderlineLinks = true;

        Assert.Equal(PdfCore.PdfColor.FromRgb(12, 34, 56), clone.LinkColor);
        Assert.False(clone.UnderlineLinks);
    }

    [Fact]
    public void Markdown_SaveAsPdf_KeepsFrontMatterTableAwayFromBodyText() {
        string markdown = """
---
date: 2026-06-01
---
The body paragraph must not touch the front matter table.
""";

        var options = new MarkdownPdfSaveOptions {
            FrontMatterRenderMode = MarkdownPdfFrontMatterRenderMode.Table,
            VisualTheme = MarkdownPdfVisualTheme.TechnicalDocument()
        };

        byte[] pdf = markdown.SaveAsPdf(options);
        IReadOnlyList<PdfLineProbe> lines = ExtractPdfLines(pdf);

        Assert.Empty(options.Warnings);
        Assert.True(BaselineGap(lines, "2026-06-01", "The body paragraph") > 18);
    }

    [Fact]
    public void Markdown_SaveAsPdf_RendersInlineCodeWithoutRawCourierFallback() {
        string markdown = """
# Inline Code

Use `OfficeIMO.Pdf` inside normal prose.
""";

        var options = new MarkdownPdfSaveOptions {
            VisualTheme = MarkdownPdfVisualTheme.TechnicalDocument()
        };

        byte[] pdf = markdown.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.Empty(options.Warnings);
        Assert.Contains("OfficeIMO.Pdf", text);
        Assert.DoesNotContain("/BaseFont /Courier", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_KeepsReadableVerticalRhythm() {
        string markdown = """
---
title: Rhythm Gate
description: Spacing should feel like a document, not extracted text.
author: OfficeIMO
---
# Rhythm Gate

This paragraph checks the gap after the document header.

> [!TIP] Rhythm check
> Panels need enough breathing room around their content.

## Checklist

- [x] Completed task keeps a readable row height.
- [ ] Open task keeps a readable row height.

| Surface | Signal |
| --- | --- |
| Table | It follows the checklist without collision. |

```csharp
Console.WriteLine("Rhythm");
```
""";

        var options = new MarkdownPdfSaveOptions {
            VisualTheme = MarkdownPdfVisualTheme.TechnicalDocument()
        };

        byte[] pdf = markdown.SaveAsPdf(options);
        IReadOnlyList<PdfLineProbe> lines = ExtractPdfLines(pdf);

        Assert.Empty(options.Warnings);
        Assert.True(BaselineGap(lines, "Rhythm Gate", "This paragraph checks") > 36);
        Assert.InRange(BaselineGap(lines, "Rhythm check", "Panels need enough"), 12, 36);
        Assert.True(BaselineGap(lines, "Panels need enough", "Checklist") > 22);
        Assert.InRange(BaselineGap(lines, "Completed task keeps", "Open task keeps"), 12, 28);
        Assert.True(BaselineGap(lines, "Open task keeps", "Surface") > 14);
        Assert.True(BaselineGap(lines, "Table", "csharp") > 22);
    }

    [Fact]
    public void Markdown_SaveAsPdf_AppliesCodeTypographyFromVisualTheme() {
        string markdown = """
# Code Theme

```csharp
Console.WriteLine("OfficeIMO");
```
""";

        MarkdownPdfVisualTheme theme = MarkdownPdfVisualTheme.Plain();
        theme.CodeBlockLabelFontSize = 7;
        theme.CodeBlockFontSize = 11;
        theme.CodeBlockLabelColor = PdfCore.PdfColor.FromRgb(255, 0, 0);
        theme.CodeBlockTextColor = PdfCore.PdfColor.FromRgb(0, 128, 0);

        var options = new MarkdownPdfSaveOptions {
            VisualTheme = theme
        };

        byte[] pdf = markdown.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.Empty(options.Warnings);
        Assert.Contains("csharp", text);
        Assert.Contains("Console.WriteLine", text);
        Assert.Contains("1 0 0 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0 0.502 0 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_PreservesNestedCalloutBlocks() {
        string markdown = """
# Nested Callout

> [!WARNING] Deployment window
> The change plan includes structured evidence.
>
> | Area | State |
> | --- | --- |
> | Backup | Ready |
> | Rollback | Tested |
>
> - [x] Snapshot copied
> - [ ] Approval recorded
>
> ```powershell
> Invoke-Deployment -WhatIf
> ```
""";

        var options = new MarkdownPdfSaveOptions {
            VisualTheme = MarkdownPdfVisualTheme.TechnicalDocument()
        };

        byte[] pdf = markdown.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.Empty(options.Warnings);
        Assert.Contains("Deployment window", text);
        Assert.Contains("Area: Backup", text);
        Assert.Contains("State: Ready", text);
        Assert.Contains("Area: Rollback", text);
        Assert.Contains("State: Tested", text);
        Assert.Contains("Done:", text);
        Assert.Contains("Open:", text);
        Assert.DoesNotContain("[x]", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("[ ]", text, StringComparison.Ordinal);
        Assert.Contains("powershell", text);
        Assert.Contains("Invoke-Deployment", text);
        Assert.Contains("0.059 0.09 0.165 rg", rawPdf, StringComparison.Ordinal);
        Assert.DoesNotContain("| Backup | Ready |", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_PreservesNestedQuoteBlocks() {
        string markdown = """
# Nested Block Content

> Context with nested structure.
>
> - First decision
> - Second decision
>
> | Decision | Owner |
> | --- | --- |
> | Ship | OfficeIMO |
""";

        var options = new MarkdownPdfSaveOptions {
            VisualTheme = MarkdownPdfVisualTheme.Report()
        };

        byte[] pdf = markdown.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.Empty(options.Warnings);
        Assert.DoesNotContain("Quote", text, StringComparison.Ordinal);
        Assert.Contains("First decision", text);
        Assert.Contains("Second decision", text);
        Assert.Contains("Decision: Ship", text);
        Assert.Contains("Owner: OfficeIMO", text);
        Assert.Contains("Ship", text);
        Assert.Contains("OfficeIMO", text);
        Assert.DoesNotContain("| Decision | Owner |", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_UsesFrontMatterVisualTheme() {
        string markdown = """
---
pdfTheme: report
---
# Report Theme

| Metric | Value |
| --- | --- |
| Quality | High |
""";

        var options = new MarkdownPdfSaveOptions();
        byte[] pdf = markdown.SaveAsPdf(options);
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.Empty(options.Warnings);
        Assert.Contains("0.118 0.251 0.686 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_AppliesReportPageDecorations() {
        string markdown = """
---
title: Quarterly Readiness
pdfTheme: report
---

# Quarterly Readiness

The report profile should feel intentionally designed without the Markdown source carrying visual markup.
""";

        var options = new MarkdownPdfSaveOptions();
        byte[] pdf = markdown.SaveAsPdf(options);
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.Empty(options.Warnings);
        Assert.Contains("Quarterly Readiness", text);
        Assert.Contains("/ExtGState", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/Shading << /SH", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/ca 0.58", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/CA 0.45", rawPdf, StringComparison.Ordinal);
        Assert.Contains("34 34 544 724 re", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_AppliesTechnicalDocumentPageDecorations() {
        string markdown = """
# Technical Readiness

The technical profile should remain quiet while still giving the page a deliberate document frame.
""";

        byte[] pdf = markdown.SaveAsPdf(new MarkdownPdfSaveOptions {
            VisualTheme = MarkdownPdfVisualTheme.TechnicalDocument()
        });
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.Contains("/ExtGState", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/Shading << /SH", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/ca 0.82", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/CA 0.55", rawPdf, StringComparison.Ordinal);
        Assert.Contains("36 36 540 720 re", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_AllowsCustomPageDecorationTheme() {
        string markdown = """
# Custom Theme

Markdown should stay semantic while the visual theme controls the page treatment.
""";

        MarkdownPdfVisualTheme theme = MarkdownPdfVisualTheme.WordLike();
        var decoration = new MarkdownPdfPageDecoration {
            BackgroundColor = PdfCore.PdfColor.White,
            PageBorder = new PdfCore.PdfPageBorder {
                Color = PdfCore.PdfColor.FromRgb(15, 118, 110),
                Width = 0.8,
                Inset = 28,
                Opacity = 0.5
            }
        }.AddBackgroundShape(PdfCore.PdfPageBackgroundShape.Rectangle(
            42,
            700,
            128,
            32,
            fill: PdfCore.PdfColor.FromRgb(204, 251, 241),
            fillOpacity: 0.44));

        theme.PageDecoration = decoration;
        decoration.PageBorder = new PdfCore.PdfPageBorder { Inset = 80 };

        byte[] pdf = markdown.SaveAsPdf(new MarkdownPdfSaveOptions {
            VisualTheme = theme
        });
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.Contains("42 700 128 32 re", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/ca 0.44", rawPdf, StringComparison.Ordinal);
        Assert.Contains("28 28 556 736 re", rawPdf, StringComparison.Ordinal);
        Assert.DoesNotContain("80 80 452 632 re", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_CanDisableReportPageDecorations() {
        string markdown = """
# Plain Report

The report colors can remain while page decoration is disabled.

| Area | State |
| --- | --- |
| Visuals | Quiet |
""";

        MarkdownPdfVisualTheme theme = MarkdownPdfVisualTheme.Report();
        theme.PageDecoration = null;

        byte[] pdf = markdown.SaveAsPdf(new MarkdownPdfSaveOptions {
            VisualTheme = theme
        });
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.Contains("0.118 0.251 0.686 rg", rawPdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Shading << /SH", rawPdf, StringComparison.Ordinal);
        Assert.DoesNotContain("34 34 544 724 re", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_PageDecorationRespectsExplicitPdfOptions() {
        string markdown = """
# Overridden Report

Explicit low-level PDF options should win over theme page decoration.
""";

        var options = new MarkdownPdfSaveOptions {
            VisualTheme = MarkdownPdfVisualTheme.Report(),
            PdfOptions = new PdfCore.PdfOptions {
                BackgroundColor = PdfCore.PdfColor.White,
                PageBorder = new PdfCore.PdfPageBorder {
                    Color = PdfCore.PdfColor.Black,
                    Width = 1,
                    Inset = 50,
                    Opacity = 1
                },
                PageBackgroundShapes = new[] {
                    PdfCore.PdfPageBackgroundShape.Rectangle(
                        12,
                        12,
                        30,
                        30,
                        fill: PdfCore.PdfColor.FromRgb(220, 252, 231))
                }
            }
        };

        byte[] pdf = markdown.SaveAsPdf(options);
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.Contains("12 12 30 30 re", rawPdf, StringComparison.Ordinal);
        Assert.Contains("50 50 512 692 re", rawPdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/ca 0.58", rawPdf, StringComparison.Ordinal);
        Assert.DoesNotContain("34 34 544 724 re", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_WarnsForUnknownFrontMatterVisualTheme() {
        string markdown = """
---
pdfTheme: spaceship
---
# Unknown Theme

Content.
""";

        var options = new MarkdownPdfSaveOptions();
        byte[] pdf = markdown.SaveAsPdf(options);
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        MarkdownPdfExportWarning warning = Assert.Single(options.Warnings);
        Assert.Equal("UnsupportedVisualTheme", warning.Code);
        Assert.Equal("spaceship", warning.Source);
        Assert.Contains("Unknown Theme", text);
    }

    private static int CountOccurrences(string value, string search) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(search, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += search.Length;
        }

        return count;
    }

    private sealed record PdfLineProbe(string Text, double BaselineY);

    private static IReadOnlyList<PdfLineProbe> ExtractPdfLines(byte[] pdf) {
        using PdfPigDocument document = PdfPigDocument.Open(new MemoryStream(pdf));
        return document.GetPage(1)
            .Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderByDescending(group => group.Key)
            .Select(group => new PdfLineProbe(string.Concat(group.OrderBy(letter => letter.StartBaseLine.X).Select(letter => letter.Value)), group.Key))
            .ToList();
    }

    private static double BaselineGap(IReadOnlyList<PdfLineProbe> lines, string upperText, string lowerText) {
        double upperY = FindBaseline(lines, upperText);
        double lowerY = FindBaseline(lines, lowerText);
        return upperY - lowerY;
    }

    private static double FindBaseline(IReadOnlyList<PdfLineProbe> lines, string text) {
        string normalizedText = NormalizePdfProbeText(text);
        foreach (PdfLineProbe line in lines) {
            if (NormalizePdfProbeText(line.Text).Contains(normalizedText, StringComparison.Ordinal)) {
                return line.BaselineY;
            }
        }

        throw new InvalidOperationException("Could not find rendered PDF line containing '" + text + "'. Lines: " + string.Join(" | ", lines.Select(line => line.Text)));
    }

    private static string NormalizePdfProbeText(string text) => new string(text.Where(ch => !char.IsWhiteSpace(ch)).ToArray());
}
