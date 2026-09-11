using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData(220, 240)]
    [InlineData(400, 160)]
    public void NestedPaddedHeadingPreservesTextLinksAndOneHeadingTag(double pageWidth, double pageHeight) {
        var options = new PdfOptions { PageWidth = pageWidth, PageHeight = pageHeight, MarginLeft = 25, MarginRight = 25, MarginTop = 25, MarginBottom = 25, DefaultFontSize = 10 };
        string heading = string.Join(" ", Enumerable.Repeat("Heading text", 24)) + " FINAL-MARKER";
        var outer = new PdfPanelStyle { PaddingY = 12, PaddingX = 8, KeepTogether = false, SpacingAfter = 0, Background = PdfColor.FromRgb(235, 240, 248) };
        var inner = new PdfPanelStyle { PaddingY = 8, PaddingX = 5, KeepTogether = false, SpacingAfter = 0, BorderColor = PdfColor.FromRgb(60, 90, 140) };
        byte[] bytes = PdfDocument.Create(options).TaggedPdfCatalogMarkers().Panel(panel => panel
            .Paragraph(p => p.Text("Outer lead"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .Panel(nested => nested
                .Paragraph(p => p.Text("Inner lead"), style: new PdfParagraphStyle { SpacingAfter = 0 })
                .H1(heading, linkUri: "https://example.test/heading", style: new PdfHeadingStyle { FontSize = 20, LineHeight = 1.15, SpacingBefore = 6, SpacingAfter = 0, KeepWithNext = false })
                .Paragraph(p => p.Text("Following paragraph"), style: new PdfParagraphStyle { SpacingAfter = 0 }), inner), outer).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 2);
        string text = string.Join("", pdf.GetPages().Select(page => page.Text));
        Assert.Contains("FINAL-MARKER", text, StringComparison.Ordinal);
        Assert.Contains("Followingparagraph", text.Replace(" ", ""), StringComparison.Ordinal);
        foreach (var page in pdf.GetPages()) Assert.All(page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)), letter => {
            Assert.InRange(letter.BoundingBox.Top, options.MarginBottom - 1, options.PageHeight - options.MarginTop - outer.PaddingY + 1);
            Assert.InRange(letter.BoundingBox.Bottom, options.MarginBottom - 1, options.PageHeight - options.MarginTop - outer.PaddingY + 1);
        });
        string raw = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Equal(1, CountOccurrences(raw, "/Type /StructElem /S /H1 "));
        var links = ExtractLinkRectangles(raw);
        Assert.NotEmpty(links);
        Assert.All(links, link => {
            Assert.InRange(link.Y1, options.MarginBottom - 1, options.PageHeight - options.MarginTop - outer.PaddingY - inner.PaddingY + 1);
            Assert.InRange(link.Y2, options.MarginBottom - 1, options.PageHeight - options.MarginTop - outer.PaddingY - inner.PaddingY + 1);
        });
        string? evidence = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_PDF_EVIDENCE");
        if (!string.IsNullOrWhiteSpace(evidence)) {
            Directory.CreateDirectory(evidence!);
            File.WriteAllBytes(Path.Combine(evidence!, "padded-heading-" + pageWidth + "-" + pageHeight + ".pdf"), bytes);
        }
    }

    [Fact]
    public void PaddedParagraphMakesProgressWhenOrphanPreferenceExceedsFrameCapacity() {
        var options = new PdfOptions { PageWidth = 400, PageHeight = 150, MarginLeft = 25, MarginRight = 25, MarginTop = 25, MarginBottom = 25, DefaultFontSize = 10 };
        byte[] bytes = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .Paragraph(p => { for (int i = 1; i <= 6; i++) { if (i > 1) p.LineBreak(); p.Text("Line" + i); } },
                style: new PdfParagraphStyle { MinimumOrphanLines = 6, MinimumWidowLines = 2, SpacingAfter = 0 }),
            new PdfPanelStyle { PaddingY = 30, PaddingX = 5, KeepTogether = false, SpacingAfter = 0 }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.InRange(pdf.NumberOfPages, 2, 3);
        string text = string.Join("", pdf.GetPages().Select(page => page.Text));
        for (int i = 1; i <= 6; i++) Assert.Contains("Line" + i, text, StringComparison.Ordinal);
    }

    [Fact]
    public void PaddedColumnParagraphRelaxesAnImpossibleOrphanPreferenceAtFrameStart() {
        var options = new PdfOptions { PageWidth = 400, PageHeight = 170, MarginLeft = 25, MarginRight = 25, MarginTop = 25, MarginBottom = 25, DefaultFontSize = 10 };
        byte[] bytes = PdfDocument.Create(options).Compose(document => document.Page(page => page.Content(content => content
            .Row(row => row.PercentColumn(100, column => column.Panel(panel => panel
                .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = 0 })
                .Paragraph(p => { for (int i = 1; i <= 6; i++) { if (i > 1) p.LineBreak(); p.Text("Line" + i); } },
                    style: new PdfParagraphStyle { MinimumOrphanLines = 6, MinimumWidowLines = 2, SpacingAfter = 0 }),
                new PdfPanelStyle { PaddingY = 25, PaddingX = 5, KeepTogether = false, SpacingAfter = 0 })))))).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.InRange(pdf.NumberOfPages, 2, 3);
        string text = string.Join("", pdf.GetPages().Select(page => page.Text));
        for (int i = 1; i <= 6; i++) Assert.Contains("Line" + i, text, StringComparison.Ordinal);
    }

    [Fact]
    public void OversizedPaddedTableLineCannotOverflowThePage() {
        var options = new PdfOptions { PageWidth = 400, PageHeight = 150, MarginLeft = 25, MarginRight = 25, MarginTop = 25, MarginBottom = 25, DefaultFontSize = 10 };
        PdfDocument document = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .Table(new[] { new[] { "A" } }, style: new PdfTableStyle { FontSize = 60, HeaderRowCount = 0, CellPaddingY = 0 }),
            new PdfPanelStyle { PaddingY = 30, PaddingX = 5, KeepTogether = false, SpacingAfter = 0 });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TableHeightRequirementsMustFitThePaddedFrame(bool fixedHeight) {
        var options = new PdfOptions { PageWidth = 400, PageHeight = 150, MarginLeft = 25, MarginRight = 25, MarginTop = 25, MarginBottom = 25, DefaultFontSize = 10 };
        var style = new PdfTableStyle { HeaderRowCount = 0, CellPaddingY = 0 };
        if (fixedHeight) style.FixedRowHeights = new System.Collections.Generic.List<double?> { 75 };
        else style.MinRowHeight = 75;
        PdfDocument document = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .Table(new[] { new[] { "A" } }, style: style),
            new PdfPanelStyle { PaddingY = 30, PaddingX = 5, KeepTogether = false, SpacingAfter = 0 });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Fact]
    public void NestedPaddedHeadingAccountsForCumulativePaddingAndRequiredSpacing() {
        var options = new PdfOptions { PageWidth = 400, PageHeight = 150, MarginLeft = 25, MarginRight = 25, MarginTop = 25, MarginBottom = 25, DefaultFontSize = 10 };
        PdfDocument document = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .Panel(nested => nested.Paragraph(p => p.Text("Inner"), style: new PdfParagraphStyle { SpacingAfter = 0 })
                .H1("A", style: new PdfHeadingStyle { FontSize = 50, LineHeight = 1, SpacingBefore = 25, ApplySpacingBeforeAtTop = true, SpacingAfter = 0, KeepWithNext = false }),
                new PdfPanelStyle { PaddingY = 20, PaddingX = 5, KeepTogether = false, SpacingAfter = 0 }),
            new PdfPanelStyle { PaddingY = 10, PaddingX = 5, KeepTogether = false, SpacingAfter = 0 });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Fact]
    public void ParagraphAfterContentRejectsALineThatCannotFitThePaddedFrame() {
        var options = new PdfOptions { PageWidth = 400, PageHeight = 150, MarginLeft = 25, MarginRight = 25, MarginTop = 25, MarginBottom = 25, DefaultFontSize = 10 };
        PdfDocument document = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .Paragraph(p => p.FontSize(55).Text("A"), style: new PdfParagraphStyle { SpacingAfter = 0 }),
            new PdfPanelStyle { PaddingY = 30, PaddingX = 5, KeepTogether = false, SpacingAfter = 0 });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Fact]
    public void OversizedPaddedListCannotSilentlyDiscardText() {
        var options = new PdfOptions { PageWidth = 400, PageHeight = 150, MarginLeft = 25, MarginRight = 25, MarginTop = 25, MarginBottom = 25, DefaultFontSize = 10 };
        PdfDocument document = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .RichBullets(new[] { PdfListItem.Rich(new[] { PdfTextRun.Normal("A") }) }, style: new PdfListStyle { FontSize = 75, LineHeight = 1 }),
            new PdfPanelStyle { PaddingY = 30, PaddingX = 5, KeepTogether = false, SpacingAfter = 0 });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Fact]
    public void OversizedPaddedRowHeadingIsRejectedAsUnrenderableInput() {
        var options = new PdfOptions { PageWidth = 400, PageHeight = 150, MarginLeft = 25, MarginRight = 25, MarginTop = 25, MarginBottom = 25, DefaultFontSize = 10 };
        PdfDocument document = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .Row(row => row.PercentColumn(100, column => column.H1("A", style: new PdfHeadingStyle { FontSize = 75, LineHeight = 1, SpacingBefore = 0, SpacingAfter = 0, KeepWithNext = false }))),
            new PdfPanelStyle { PaddingY = 30, PaddingX = 5, KeepTogether = false, SpacingAfter = 0 });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Fact]
    public void HeadingAfterContentRejectsALineThatCannotFitThePaddedFrame() {
        var options = new PdfOptions { PageWidth = 400, PageHeight = 150, MarginLeft = 25, MarginRight = 25, MarginTop = 25, MarginBottom = 25, DefaultFontSize = 10 };
        PdfDocument document = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .H1("A", style: new PdfHeadingStyle { FontSize = 75, LineHeight = 1, SpacingBefore = 0, SpacingAfter = 0, KeepWithNext = false }),
            new PdfPanelStyle { PaddingY = 30, PaddingX = 5, KeepTogether = false, SpacingAfter = 0 });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }
}
