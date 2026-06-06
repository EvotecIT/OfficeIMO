using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPdfTests {
    [Fact]
    public void Html_SaveAsPdf_SemanticProfile_ExportsThroughMarkdownPipeline() {
        var options = new HtmlPdfSaveOptions {
            Profile = HtmlPdfProfile.Semantic
        };

        byte[] pdf = """
<article>
  <h1>HTML Report</h1>
  <p><strong>OfficeIMO</strong> turns semantic HTML into PDF.</p>
  <ul><li>Markdown bridge</li><li>Shared PDF engine</li></ul>
</article>
""".SaveAsPdf(options);

        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.True(pdf.Length > 0);
        Assert.Contains("HTML Report", text);
        Assert.Contains("Markdown bridge", text);
        Assert.False(options.ConversionReport.HasWarnings);
    }

    [Fact]
    public void Html_SaveAsPdf_SemanticProfile_ForwardsMarkdownPdfWarningsToSharedReport() {
        var markdownOptions = new MarkdownPdfSaveOptions();
        var options = new HtmlPdfSaveOptions {
            Profile = HtmlPdfProfile.Semantic,
            MarkdownHtmlOptions = HtmlToMarkdownOptions.CreateOfficeIMOProfile(),
            MarkdownPdfOptions = markdownOptions
        };

        byte[] pdf = """
<h1>Remote Asset</h1>
<p><img src="https://example.com/logo.png" alt="OfficeIMO logo"></p>
""".SaveAsPdf(options);

        Assert.True(pdf.Length > 0);
        Assert.Single(markdownOptions.Warnings);
        PdfCore.PdfConversionWarning warning = Assert.Single(options.ConversionReport.Warnings);
        Assert.Equal("OfficeIMO.Markdown.Pdf", warning.Converter);
        Assert.Equal("UnsupportedImage", warning.Code);
    }

    [Fact]
    public void Html_SaveAsPdf_DocumentProfile_ExportsThroughWordPipeline() {
        var options = new HtmlPdfSaveOptions {
            Profile = HtmlPdfProfile.Document,
            WordHtmlOptions = HtmlToWordOptions.CreateOfficeIMOProfile(),
            WordPdfOptions = new PdfSaveOptions()
        };

        byte[] pdf = """
<html>
  <body>
    <h1>Document HTML</h1>
    <p>Rendered through the Word HTML bridge.</p>
    <table><tr><th>Area</th><th>Status</th></tr><tr><td>HTML</td><td>PDF</td></tr></table>
  </body>
</html>
""".SaveAsPdf(options);

        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.True(pdf.Length > 0);
        Assert.Contains("Document HTML", text);
        Assert.Contains("Word HTML bridge", text);
    }

    [Fact]
    public void Pdf_ToHtml_SemanticProfile_ExportsLogicalStructure() {
        byte[] pdf = CreateLogicalSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            LayoutOptions = new PdfCore.PdfTextLayoutOptions {
                ForceSingleColumn = true
            }
        };

        string html = PdfHtmlConverter.ToHtml(pdf, options);

        Assert.Contains("<title>Logical PDF sample</title>", html, StringComparison.Ordinal);
        Assert.Contains("<h1>Logical Heading</h1>", html, StringComparison.Ordinal);
        Assert.Contains("<p>Logical readback marker.</p>", html, StringComparison.Ordinal);
        Assert.Contains("<ul data-pdf-list-level=\"1\"><li>Detected logical bullet</li></ul>", html, StringComparison.Ordinal);
        Assert.Contains("<table>", html, StringComparison.Ordinal);
        Assert.Contains("<th>Code</th>", html, StringComparison.Ordinal);
        Assert.Contains("<td>A-100</td>", html, StringComparison.Ordinal);
        Assert.Equal(1, CountOccurrences(html, "A-100"));
        Assert.False(options.ConversionReport.HasWarnings);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_ExportsPageGeometryAndTextBlocks() {
        byte[] pdf = CreateLogicalSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            LayoutOptions = new PdfCore.PdfTextLayoutOptions {
                ForceSingleColumn = true
            }
        };

        string html = PdfHtmlConverter.ToHtml(PdfCore.PdfReadDocument.Load(pdf), options);

        Assert.Contains(".pdf-page{position:relative", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-page\" data-page-number=\"1\" style=\"width:420pt;height:360pt;\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-text pdf-heading\"", html, StringComparison.Ordinal);
        Assert.Contains("style=\"left:", html, StringComparison.Ordinal);
        Assert.Contains("Logical Heading", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_ExportsPositionedImagePlaceholders() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Canvas(canvas => canvas.Image(PdfPngTestImages.CreateRgbPng(1, 1), 40, 50, 60, 30))
            .ToBytes();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview
        };

        string html = PdfHtmlConverter.ToHtml(pdf, options);

        Assert.Contains("class=\"pdf-image-placeholder\"", html, StringComparison.Ordinal);
        Assert.Contains("style=\"position:absolute;left:40pt;top:50pt;width:60pt;height:30pt;\"", html, StringComparison.Ordinal);
        Assert.Contains("data-matrix=\"60 0 0 30 40 140\"", html, StringComparison.Ordinal);
        Assert.False(options.ConversionReport.HasWarnings);
    }

    [Fact]
    public void Pdf_ToHtml_PageRanges_ExportsSelectedPagesThroughSameBridgePackage() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .Paragraph(paragraph => paragraph.Text("First PDF page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second PDF page"))
            .ToBytes();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            PageRanges = new[] {
                PdfCore.PdfPageRange.From(2, 2)
            }
        };

        string html = PdfHtmlConverter.ToHtml(pdf, options);

        Assert.DoesNotContain("First PDF page", html, StringComparison.Ordinal);
        Assert.Contains("Second PDF page", html, StringComparison.Ordinal);
    }

    private static byte[] CreateLogicalSamplePdf() {
        return PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "Logical PDF sample", author: "OfficeIMO")
            .H1("Logical Heading")
            .Paragraph(paragraph => paragraph.Text("Logical readback marker."))
            .Bullets(new[] { "Detected logical bullet" })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += value.Length;
        }

        return count;
    }
}
