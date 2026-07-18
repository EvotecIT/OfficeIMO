using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Tests.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class MarkdownPdfTests {
    private static byte[] CreateMinimalRgbPng() => PdfPngTestImages.CreateRgbPng(1, 1);

    private static string CreateMinimalRgbPngDataUri() =>
        "data:image/png;base64," + Convert.ToBase64String(CreateMinimalRgbPng());

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

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);

        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

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
    public void Markdown_SaveAsPdf_PanelLists_DoNotDuplicate_ListItem_LeadParagraphs() {
        string markdown = """
> - Alpha
>   - Beta
""";

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(new MarkdownPdfSaveOptions());
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Equal(1, CountOccurrences(text, "Alpha"));
        Assert.Equal(1, CountOccurrences(text, "Beta"));
    }

    [Fact]
    public void Markdown_SaveAsPdf_LongBlockquotePanelSplitsAcrossPages() {
        string markdown = string.Join("\n", Enumerable.Range(1, 24).Select(index => "> QuoteLine" + index.ToString()));
        var options = new MarkdownPdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                PageWidth = 180,
                PageHeight = 130,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFont = PdfCore.PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            }
        };

        byte[] pdfBytes = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);

        using var pdf = PdfPigDocument.Open(new MemoryStream(pdfBytes));
        string text = PdfCore.PdfReadDocument.Open(pdfBytes).ExtractText();

        Assert.True(pdf.NumberOfPages > 1);
        Assert.Contains("QuoteLine1", text, StringComparison.Ordinal);
        Assert.Contains("QuoteLine24", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_DefaultTextFallbacksCoverCommonSymbolsWhenAvailable() {
        const string symbol = "\u26A0";
        PdfCore.PdfEmbeddedFontFallbackSet? fallbackSet = new PdfCore.PdfOptions()
            .UseTextFallbacks()
            .EmbeddedFontFallbacks;
        if (fallbackSet == null ||
            !fallbackSet.PlanText(symbol).IsFullyCovered) {
            return;
        }

        var options = new MarkdownPdfSaveOptions();
        string markdown = "> Status " + symbol + " marker";

        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("Status", text, StringComparison.Ordinal);
        Assert.Contains("marker", text, StringComparison.Ordinal);
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "unsupported-text-glyph");
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "missing-embedded-font-fallback-glyph");
    }

    [Fact]
    public void Markdown_SaveAsPdf_MultilingualTextFallbacksCoverCjkWhenAvailable() {
        const string cjk = "\u300D\u4E00";
        PdfCore.PdfEmbeddedFontFallbackSet? fallbackSet = new PdfCore.PdfOptions()
            .UseTextFallbacks(
                PdfCore.PdfTextFallbackFeatures.DocumentFont |
                PdfCore.PdfTextFallbackFeatures.MultilingualFonts |
                PdfCore.PdfTextFallbackFeatures.SymbolAndEmojiFonts)
            .EmbeddedFontFallbacks;
        if (fallbackSet == null || !fallbackSet.PlanText(cjk).IsFullyCovered) return;

        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader
            .Parse("# Multilingual\n\nText " + cjk)
            .ToPdfDocumentResult(new MarkdownPdfSaveOptions {
                TextFallbacks = PdfCore.PdfTextFallbackFeatures.Default |
                    PdfCore.PdfTextFallbackFeatures.MultilingualFonts
            });
        byte[] pdf = result.ToBytes();

        Assert.Equal("%PDF", System.Text.Encoding.ASCII.GetString(pdf, 0, 4));
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "unsupported-text-glyph");
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "missing-embedded-font-fallback-glyph");
    }

    [Fact]
    public void MarkdownDoc_SaveAsPdf_Renders_SoftBreak_As_Space() {
        var markdown = MarkdownDoc.Create();
        markdown.Add(new ParagraphBlock(new InlineSequence()
            .Text("Alpha")
            .SoftBreak()
            .Text("Beta")));

        byte[] pdf = markdown.ToPdf(new MarkdownPdfSaveOptions());
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("Alpha Beta", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Alpha\nBeta", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_PreservesExplicitTableColumnAlignment() {
        string markdown = """
# Invoice

| Item | Center Qty | Right Qty |
| :--- | :---: | ---: |
| Service | 77 | 88 |
""";

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(new MarkdownPdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 260,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            }
        });

        PdfTextBounds centerHeader = FindPdfTextBounds(pdf, "Center", "Qty");
        PdfTextBounds rightHeader = FindPdfTextBounds(pdf, "Right", "Qty");
        PdfTextBounds centerQuantity = FindPdfTextBounds(pdf, "77");
        PdfTextBounds rightQuantity = FindPdfTextBounds(pdf, "88");

        Assert.InRange(Math.Abs(centerQuantity.Center - centerHeader.Center), 0D, 2D);
        Assert.InRange(Math.Abs(rightQuantity.Right - rightHeader.Right), 0D, 2D);
    }

    [Fact]
    public void Markdown_SaveAsPdf_UsesTableColumnWidthHints() {
        var defaultTable = new TableBlock();
        defaultTable.Headers.Add("Code");
        defaultTable.Headers.Add("Description");
        defaultTable.Rows.Add(new[] { "A-100", "Consulting" });

        var narrowFirstColumnTable = new TableBlock();
        narrowFirstColumnTable.Headers.Add("Code");
        narrowFirstColumnTable.Headers.Add("Description");
        narrowFirstColumnTable.Rows.Add(new[] { "A-100", "Consulting" });
        narrowFirstColumnTable.ColumnWidthPoints.Add(48D);
        narrowFirstColumnTable.ColumnWidthPoints.Add(null);

        var options = new MarkdownPdfSaveOptions {
            PdfOptions = new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 260,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            }
        };

        byte[] defaultPdf = MarkdownDoc.Create().Add(defaultTable).ToPdf(options);
        byte[] narrowPdf = MarkdownDoc.Create().Add(narrowFirstColumnTable).ToPdf(options);

        PdfTextBounds defaultDescription = FindPdfTextBounds(defaultPdf, "Description");
        PdfTextBounds narrowDescription = FindPdfTextBounds(narrowPdf, "Description");

        Assert.True(narrowDescription.Left < defaultDescription.Left - 70D, $"Expected the second column to move left when the first column width is fixed. Default left: {defaultDescription.Left}; narrow left: {narrowDescription.Left}.");
    }

    [Fact]
    public void Markdown_SaveAsPdf_RecordsWarningsForRemoteImages() {
        var options = new MarkdownPdfSaveOptions();
        string markdown = """
# Remote Asset

![OfficeIMO logo](https://example.com/logo.png)
""";

        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.True(pdf.Length > 0);
        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "RemoteImageDisabled");
        Assert.Equal("OfficeIMO.Markdown.Pdf", warning.Converter);
        Assert.Contains("OfficeIMO logo", text);
    }

    [Fact]
    public void Markdown_ToPdfDocumentResult_ReturnsPdfDocumentAndReportSnapshot() {
        var options = new MarkdownPdfSaveOptions();
        string markdown = """
# Remote Asset

![OfficeIMO logo](https://example.com/logo.png)
""";

        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdfDocumentResult(options);
        PdfCore.PdfDocument processed = result.Value.AppendMetadataRevision(title: "Processed Markdown PDF");

        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "RemoteImageDisabled");
        Assert.True(result.HasWarnings);
        Assert.Equal("OfficeIMO.Markdown.Pdf", warning.Converter);
        Assert.Equal("Processed Markdown PDF", processed.Inspect().Metadata.Title);
        Assert.Contains("OfficeIMO logo", result.Value.Read.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_SaveAsPdf_BlocksLocalImagesByDefault() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Markdown.Pdf.LocalImages", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string imagePath = Path.Combine(directory, "pixel.png");
            File.WriteAllBytes(imagePath, CreateMinimalRgbPng());

            var options = new MarkdownPdfSaveOptions {
                BaseDirectory = directory
            };
            string markdown = "![Local pixel](pixel.png){width=24 height=24}";

            PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdfDocumentResult(options);
            byte[] pdf = result.ToBytes();
            string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
            IReadOnlyList<PdfCore.PdfExtractedImage> images = PdfCore.PdfImageExtractor.ExtractImages(pdf);

            PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings);
            Assert.Equal("LocalImageDisabled", warning.Code);
            Assert.Contains("[Image unavailable:", text, StringComparison.Ordinal);
            Assert.Empty(images);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    [Fact]
    public void Markdown_SaveAsPdf_EmbedsRemoteImagesThroughExplicitResolver() {
        Uri? requestedUri = null;
        var options = new MarkdownPdfSaveOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost(),
            RemoteImageResolver = uri => {
                requestedUri = uri;
                return CreateMinimalRgbPng();
            }
        };
        string markdown = """
# Remote Asset

![OfficeIMO logo](https://example.com/logo.png){width=24 height=24}
""";

        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
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
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost(),
            MaximumRemoteImageBytes = 3,
            RemoteImageResolver = _ => new byte[] { 1, 2, 3, 4 }
        };
        string markdown = """
# Remote Asset

![OfficeIMO logo](https://example.com/logo.png)
""";

        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "ImageTooLarge");
        Assert.Equal("ImageTooLarge", warning.Code);
        Assert.Contains("OfficeIMO logo", text);
        Assert.Contains("[Image unavailable:", text, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownPdfConverter_SaveFileAsPdf_ResolvesRelativeLocalImages() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Markdown.Pdf", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string markdownPath = Path.Combine(directory, "README.md");
            string pdfPath = Path.Combine(directory, "README.pdf");
            string imagePath = Path.Combine(directory, "pixel.png");

            File.WriteAllBytes(imagePath, CreateMinimalRgbPng());
            File.WriteAllText(markdownPath, """
# Asset Report

![Local pixel](pixel.png){width=32 height=32}
_Figure 1. Embedded from a relative Markdown path._
""");

            var options = new MarkdownPdfSaveOptions {
                ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost(),
                BaseDirectory = directory
            };
            PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownDoc.Load(markdownPath).ToPdfDocumentResult(options);
            result.Save(pdfPath);

            byte[] pdf = File.ReadAllBytes(pdfPath);
            string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
            IReadOnlyList<PdfCore.PdfExtractedImage> images = PdfCore.PdfImageExtractor.ExtractImages(pdf);

            Assert.True(pdf.Length > 0);
            Assert.Empty(options.Warnings);
            Assert.Equal(directory, options.BaseDirectory);
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
    public void MarkdownPdfConverter_SaveFileAsPdf_BlocksLocalImagesOutsideBaseDirectory() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Markdown.Pdf.BaseDirectory", Guid.NewGuid().ToString("N"));
        string markdownDirectory = Path.Combine(directory, "docs");
        Directory.CreateDirectory(markdownDirectory);
        try {
            string markdownPath = Path.Combine(markdownDirectory, "README.md");
            string pdfPath = Path.Combine(markdownDirectory, "README.pdf");
            string outsideImagePath = Path.Combine(directory, "secret.png");

            File.WriteAllBytes(outsideImagePath, CreateMinimalRgbPng());
            File.WriteAllText(markdownPath, """
# Escaped Asset

![Secret pixel](../secret.png){width=32 height=32}
""");

            var options = new MarkdownPdfSaveOptions {
                ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost(),
                BaseDirectory = markdownDirectory
            };
            PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownDoc.Load(markdownPath).ToPdfDocumentResult(options);
            result.Save(pdfPath);

            byte[] pdf = File.ReadAllBytes(pdfPath);
            string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
            IReadOnlyList<PdfCore.PdfExtractedImage> images = PdfCore.PdfImageExtractor.ExtractImages(pdf);

            PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings);
            Assert.Equal("LocalImageOutsideBaseDirectory", warning.Code);
            Assert.Equal(markdownDirectory, options.BaseDirectory);
            Assert.Contains("[Image unavailable:", text, StringComparison.Ordinal);
            Assert.Empty(images);
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void MarkdownPdfConverter_BlocksLocalImageSymlinkThatEscapesBaseDirectory() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Markdown.Pdf.Symlink", Guid.NewGuid().ToString("N"));
        string markdownDirectory = Path.Combine(directory, "docs");
        string outsideDirectory = Path.Combine(directory, "outside");
        string linkPath = Path.Combine(markdownDirectory, "linked.png");
        Directory.CreateDirectory(markdownDirectory);
        Directory.CreateDirectory(outsideDirectory);
        try {
            string outsideImagePath = Path.Combine(outsideDirectory, "secret.png");
            File.WriteAllBytes(outsideImagePath, CreateMinimalRgbPng());
            File.CreateSymbolicLink(linkPath, outsideImagePath);

            var options = new MarkdownPdfSaveOptions {
                ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost(),
                BaseDirectory = markdownDirectory
            };
            PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader
                .Parse("![Secret pixel](linked.png){width=32 height=32}")
                .ToPdfDocumentResult(options);

            PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings);
            Assert.Equal("LocalImageOutsideBaseDirectory", warning.Code);
            Assert.Empty(PdfCore.PdfImageExtractor.ExtractImages(result.ToBytes()));
        } finally {
            if (File.Exists(linkPath)) File.Delete(linkPath);
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }
#endif

    [Fact]
    public void Markdown_SaveAsPdf_ScalesOversizedLocalImagesIntoContentFrame() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Markdown.Pdf.ScaleDown", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string imagePath = Path.Combine(directory, "pixel.png");
            File.WriteAllBytes(imagePath, CreateMinimalRgbPng());

            var options = new MarkdownPdfSaveOptions {
                ApplyDefaultTheme = false,
                BaseDirectory = directory,
                ResourcePolicy = PdfCore.PdfResourcePolicy.CreateTrustedHost(),
                PdfOptions = new PdfCore.PdfOptions {
                    PageWidth = 220,
                    PageHeight = 180,
                    MarginLeft = 20,
                    MarginRight = 20,
                    MarginTop = 20,
                    MarginBottom = 20
                }
            };
            string markdown = "![Wide local image](pixel.png){width=360 height=180}";

            byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
            string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

            Assert.Empty(options.Warnings);
            Assert.Contains("q\n180 0 0 90 20 70 cm\n/Im1 Do\nQ", rawPdf);
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

        PdfCore.PdfSaveResult streamResult = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).TrySaveAsPdf(stream);

        Assert.True(streamResult.Succeeded);
        Assert.Null(streamResult.OutputPath);
        Assert.True(streamResult.BytesWritten > 0);
        Assert.Empty(streamResult.Diagnostics);
        Assert.Equal(streamResult.BytesWritten, stream.ToArray().LongLength);

        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Markdown.Pdf.Result", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string pdfPath = Path.Combine(directory, "result.pdf");
        try {
            PdfCore.PdfSaveResult pathResult = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).TrySaveAsPdf(pdfPath);

            Assert.True(pathResult.Succeeded);
            Assert.Equal(Path.GetFullPath(pdfPath), pathResult.OutputPath);
            Assert.Equal(File.ReadAllBytes(pdfPath).LongLength, pathResult.BytesWritten);

            PdfCore.PdfSaveResult directoryResult = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).TrySaveAsPdf(directory);

            Assert.False(directoryResult.Succeeded);
            Assert.NotEmpty(directoryResult.Diagnostics);
            Assert.Throws<InvalidOperationException>(() => directoryResult.RequireSuccess());
        } finally {
            if (Directory.Exists(directory)) {
                Directory.Delete(directory, recursive: true);
            }
        }

        Assert.Throws<ArgumentNullException>(() => OfficeIMO.Markdown.MarkdownReader.Parse((string)null!));
    }

    [Fact]
    public void Markdown_SaveAsPdf_EmbedsDataUriImages() {
        var options = new MarkdownPdfSaveOptions();
        string markdown =
            "# Inline Asset\n\n" +
            "![Inline pixel](" + CreateMinimalRgbPngDataUri() + "){width=24 height=24}\n";

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
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
            Style = MarkdownPdfStyle.GitHubLike()
        };
        string markdown = """
# Checklist

- [x] Ship first-party Markdown PDF
- [ ] Add multi-block panel primitive
""";

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
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
        MarkdownPdfStyle theme = MarkdownPdfStyle.Plain();
        theme.ChecklistCheckedIconColor = PdfCore.PdfColor.FromRgb(255, 0, 0);
        theme.ChecklistUncheckedIconColor = PdfCore.PdfColor.FromRgb(0, 0, 255);
        theme.ChecklistCheckedTextColor = PdfCore.PdfColor.FromRgb(0, 128, 0);
        theme.ChecklistUncheckedTextColor = PdfCore.PdfColor.FromRgb(255, 0, 255);
        theme.ChecklistCheckedFillColor = PdfCore.PdfColor.FromRgb(255, 255, 204);
        theme.ChecklistUncheckedFillColor = PdfCore.PdfColor.FromRgb(204, 238, 255);
        var options = new MarkdownPdfSaveOptions {
            Style = theme
        };
        string markdown = """
# Checklist Theme

- [x] Completed item
- [ ] Open item
""";

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
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
    public void MarkdownPdfStyle_Maps_SharedMarkdownTheme() {
        MarkdownVisualTheme sharedTheme = MarkdownVisualTheme.Report()
            .WithColors(
                accent: "#ff0000",
                heading: "#00ff00",
                text: "#0000ff",
                border: "#112233",
                tableHeaderBackground: "#ffff00",
                tableHeaderText: "#000000")
            .WithTable(table => {
                table.BorderWidth = 1.25;
                table.CellPaddingX = 11;
                table.CellPaddingY = 7;
            });

        MarkdownPdfStyle pdfTheme = MarkdownPdfStyle.FromMarkdownTheme(sharedTheme);

        Assert.Equal("Report", pdfTheme.Name);
        Assert.Equal(0, pdfTheme.DocumentHeaderTitleColor!.Value.R, 3);
        Assert.Equal(1, pdfTheme.DocumentHeaderTitleColor!.Value.G, 3);
        Assert.Equal(0, pdfTheme.DocumentHeaderTitleColor!.Value.B, 3);
        Assert.Equal(1, pdfTheme.LinkColor!.Value.R, 3);
        Assert.Equal(0, pdfTheme.LinkColor!.Value.G, 3);
        Assert.Equal(0, pdfTheme.LinkColor!.Value.B, 3);
        Assert.Equal(1.25, pdfTheme.TableStyle!.BorderWidth);
        Assert.Equal(11, pdfTheme.TableStyle.CellPaddingX);
        Assert.Equal(7, pdfTheme.TableStyle.CellPaddingY);
        Assert.Equal(1, pdfTheme.TableStyle.HeaderFill!.Value.R, 3);
        Assert.Equal(1, pdfTheme.TableStyle.HeaderFill!.Value.G, 3);
        Assert.Equal(0, pdfTheme.TableStyle.HeaderFill!.Value.B, 3);
    }

    [Fact]
    public void Markdown_SaveAsPdf_Renders_With_SharedMarkdownTheme() {
        MarkdownVisualTheme theme = MarkdownVisualTheme.Report()
            .WithColorScheme(MarkdownColorSchemeKind.Emerald);
        string markdown = """
# Shared Theme

Markdown PDF should accept the same visual theme object as HTML and Word.

| Surface | Status |
| --- | --- |
| PDF | themed |
""";

        var options = new MarkdownPdfSaveOptions {
            Theme = theme
        };

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.True(pdf.Length > 0);
        Assert.Empty(options.Warnings);
        Assert.Contains("Shared Theme", text);
        Assert.Contains("themed", text);
    }

    [Fact]
    public void Markdown_SaveAsPdf_AppliesVisualThemeLinkStyleAcrossInlineSurfaces() {
        MarkdownPdfStyle theme = MarkdownPdfStyle.TechnicalDocument();
        theme.LinkColor = PdfCore.PdfColor.FromRgb(128, 0, 128);
        theme.UnderlineLinks = false;
        var options = new MarkdownPdfSaveOptions {
            Style = theme
        };
        string markdown = """
# Link Theme

Paragraph [paragraph link](https://example.com/paragraph).

- [x] [task link](https://example.com/task)

| Surface | Link |
| --- | --- |
| Table | [table link](https://example.com/table) |
""";

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
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

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf();
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
            Style = MarkdownPdfStyle.TechnicalDocument()
        };
        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
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
        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

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
        PdfCore.PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.True(pdf.Length > 0);
        Assert.Contains(result.Warnings, warning => warning.Code == "UnsupportedSemanticFence" && warning.Source == "diagram");
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
            Style = MarkdownPdfStyle.TechnicalDocument()
        };

        byte[] pdf = document.ToPdf(options);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
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
            Style = MarkdownPdfStyle.TechnicalDocument()
        };

        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
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
            Style = MarkdownPdfStyle.TechnicalDocument()
        };

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.Empty(options.Warnings);
        Assert.Contains("Better PDFs", text);
        Assert.Contains("Visuals", text);
        Assert.Contains("0.059 0.09 0.165 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownPdfStyles_UseDocumentRhythmForTables() {
        MarkdownPdfStyle theme = MarkdownPdfStyle.TechnicalDocument();

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
    public void MarkdownPdfStyles_CloneLinkStyleOptions() {
        MarkdownPdfStyle theme = MarkdownPdfStyle.Report();
        theme.LinkColor = PdfCore.PdfColor.FromRgb(12, 34, 56);
        theme.UnderlineLinks = false;

        MarkdownPdfStyle clone = theme.Clone();
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
            Style = MarkdownPdfStyle.TechnicalDocument()
        };

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
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
            Style = MarkdownPdfStyle.TechnicalDocument()
        };

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
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
            Style = MarkdownPdfStyle.TechnicalDocument()
        };

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
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

        MarkdownPdfStyle theme = MarkdownPdfStyle.Plain();
        theme.CodeBlockLabelFontSize = 7;
        theme.CodeBlockFontSize = 11;
        theme.CodeBlockLabelColor = PdfCore.PdfColor.FromRgb(255, 0, 0);
        theme.CodeBlockTextColor = PdfCore.PdfColor.FromRgb(0, 128, 0);

        var options = new MarkdownPdfSaveOptions {
            Style = theme
        };

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
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
            Style = MarkdownPdfStyle.TechnicalDocument()
        };

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
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
            Style = MarkdownPdfStyle.Report()
        };

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

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
        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);

        Assert.Empty(options.Warnings);
        Assert.Contains("0.859 0.918 0.996 rg", rawPdf, StringComparison.Ordinal);
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
        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
        string rawPdf = System.Text.Encoding.ASCII.GetString(pdf);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

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

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(new MarkdownPdfSaveOptions {
            Style = MarkdownPdfStyle.TechnicalDocument()
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

        MarkdownPdfStyle theme = MarkdownPdfStyle.WordLike();
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

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(new MarkdownPdfSaveOptions {
            Style = theme
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

        MarkdownPdfStyle theme = MarkdownPdfStyle.Report();
        theme.PageDecoration = null;

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(new MarkdownPdfSaveOptions {
            Style = theme
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
            Style = MarkdownPdfStyle.Report(),
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

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
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
        PdfCore.PdfDocumentConversionResult result = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings);
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

    private sealed record PdfTextBounds(double Left, double Right) {
        public double Center => (Left + Right) / 2D;
    }

    private static PdfTextBounds FindPdfTextBounds(byte[] pdf, params string[] texts) {
        using PdfPigDocument document = PdfPigDocument.Open(new MemoryStream(pdf));
        var lines = document.GetPage(1)
            .GetWords()
            .GroupBy(word => Math.Round(word.BoundingBox.Bottom, 1))
            .Select(group => group.OrderBy(word => word.BoundingBox.Left).ToList())
            .ToList();

        foreach (var line in lines) {
            for (int index = 0; index <= line.Count - texts.Length; index++) {
                bool matches = true;
                for (int offset = 0; offset < texts.Length; offset++) {
                    if (!string.Equals(line[index + offset].Text, texts[offset], StringComparison.Ordinal)) {
                        matches = false;
                        break;
                    }
                }

                if (matches) {
                    double left = line.Skip(index).Take(texts.Length).Min(word => word.BoundingBox.Left);
                    double right = line.Skip(index).Take(texts.Length).Max(word => word.BoundingBox.Right);
                    return new PdfTextBounds(left, right);
                }
            }
        }

        throw new InvalidOperationException("Could not find rendered PDF text '" + string.Join(" ", texts) + "'.");
    }

    private static string NormalizePdfProbeText(string text) => new string(text.Where(ch => !char.IsWhiteSpace(ch)).ToArray());
}
