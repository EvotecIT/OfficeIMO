using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData(3, 2, 200, 0)]
    [InlineData(2, 0, 200, 0)]
    [InlineData(3, 2, 205, 0)]
    [InlineData(2, 0, 205, 0)]
    [InlineData(3, 2, 205, 1)]
    [InlineData(2, 0, 205, 1)]
    [InlineData(3, 2, 208, 2)]
    [InlineData(2, 0, 208, 2)]
    [InlineData(3, 2, 216, 3)]
    [InlineData(2, 0, 216, 3)]
    public void ShortWidowParagraphMovesBeforeOpeningItsPanel(int lineCount, int orphanLines, double pageHeight, int wrapper) {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(400), PageHeight = global::OfficeIMO.OfficeLength.Points(pageHeight), MarginLeft = global::OfficeIMO.OfficeLength.Points(20), MarginRight = global::OfficeIMO.OfficeLength.Points(20),
            MarginTop = global::OfficeIMO.OfficeLength.Points(20), MarginBottom = global::OfficeIMO.OfficeLength.Points(20), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        var paragraphStyle = new PdfParagraphStyle { FontSize = global::OfficeIMO.OfficeLength.Points(10), LineHeight = 1, SpacingBefore = global::OfficeIMO.OfficeLength.Points(0),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), MinimumOrphanLines = orphanLines, MinimumWidowLines = 2 };
        var panelStyle = new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(5), PaddingX = global::OfficeIMO.OfficeLength.Points(5), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), RepeatFragmentDecoration = false };
        Action<PdfContentBuilder> paragraph = panel => panel.Paragraph(p => {
            for (int i = 1; i <= lineCount; i++) { if (i > 1) p.LineBreak(); p.Text("Target" + i); }
        }, style: paragraphStyle);
        Action<PdfContentBuilder> content = wrapper switch {
            1 => panel => panel.Bookmark("start").Semantic(PdfSemanticRole.Section,
                semantic => semantic.Layer("text", layer => layer.Flow(paragraph))).Bookmark("end"),
            2 or 3 => panel => panel.Panel(paragraph, new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(3), PaddingX = global::OfficeIMO.OfficeLength.Points(0),
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(wrapper == 3 ? 7 : 0), RepeatFragmentDecoration = false }),
            _ => paragraph
        };
        byte[] referenceBytes = PdfDocument.Create(options).Panel(content, panelStyle).ToBytes();
        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => { for (int i = 0; i < 16 - lineCount; i++) { if (i > 0) p.LineBreak(); p.Text("Lead" + i); } },
                style: new PdfParagraphStyle { FontSize = global::OfficeIMO.OfficeLength.Points(10), LineHeight = 1, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
            .Panel(content, panelStyle).ToBytes();
        SavePaddedPaginationEvidence($"short-widow-{lineCount}-{pageHeight}-{wrapper}-reference", referenceBytes);
        SavePaddedPaginationEvidence($"short-widow-{lineCount}-{pageHeight}-{wrapper}", bytes);
        using var reference = PdfPigDocument.Open(referenceBytes);
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text.StartsWith("Target", StringComparison.Ordinal));
        double expected = Assert.Single(reference.GetPage(1).GetWords(), word => word.Text == "Target1").Letters[0].StartBaseLine.Y;
        double actual = Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == "Target1").Letters[0].StartBaseLine.Y;
        Assert.Equal(expected, actual, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PaddedFixedObjectsRejectHeightBeyondTheirResumedFrame(bool column) {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(400), PageHeight = global::OfficeIMO.OfficeLength.Points(240), MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        var style = new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(12), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = false, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) };
        PdfDocument document = PdfDocument.Create(options);
        if (column) document.Compose(d => d.Page(page => page.Content(content => content.Row(row => row.PercentColumn(100, col => col.Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
            .Shape(OfficeIMO.Drawing.OfficeShape.Rectangle(25, 185), style: new PdfDrawingStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }), style))))));
        else document.Panel(panel => panel.Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
            .Shape(OfficeIMO.Drawing.OfficeShape.Rectangle(25, 185), style: new PdfDrawingStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }), style);
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Theory]
    [InlineData(false, false, false, false)]
    [InlineData(true, false, false, false)]
    [InlineData(false, true, false, false)]
    [InlineData(true, true, false, false)]
    [InlineData(false, true, false, true)]
    [InlineData(true, true, false, true)]
    [InlineData(false, true, true, false)]
    [InlineData(true, true, true, false)]
    [InlineData(false, true, true, true)]
    [InlineData(true, true, true, true)]
    public void PaddedImageScalingFitsTheResumedFrame(bool column, bool firstChild, bool nested, bool keepTogether) {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(400), PageHeight = global::OfficeIMO.OfficeLength.Points(240), MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        var style = new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(12), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = keepTogether, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) };
        var imageStyle = new PdfImageStyle { ScaleDownToFit = true, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) };
        PdfDocument document = PdfDocument.Create(options);
        Action<PdfContentBuilder> content = panel => {
            if (!firstChild) panel.Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) });
            panel.Image(CreateMinimalRgbPng(), global::OfficeIMO.OfficeLength.Points(25), global::OfficeIMO.OfficeLength.Points(185), style: imageStyle, linkUri: "https://example.test/image");
        };
        Action<PdfContentBuilder> outerContent = nested
            ? panel => panel.Panel(content, new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(8), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = keepTogether, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
            : content;
        if (column) document.Compose(d => d.Page(page => page.Content(body => body.Row(row => row.PercentColumn(100, col => col.Panel(outerContent, style))))));
        else document.Panel(outerContent, style);
        byte[] bytes = document.ToBytes();
        var links = ExtractLinkRectangles(System.Text.Encoding.ASCII.GetString(bytes));
        var link = Assert.Single(links);
        double padding = nested ? 20 : 12;
        Assert.InRange(link.Y1, 25 + padding - 0.01, 215 - padding);
        Assert.InRange(link.Y2, 25 + padding, 215 - padding + 0.01);
        SavePaddedPaginationEvidence("padded-image-" + (column ? "column" : "flow") + (firstChild ? "-first" : "") + (nested ? "-nested" : "") + (keepTogether ? "-keep" : ""), bytes);
    }

    [Fact]
    public void PaddedRepeatedTableHeaderReservesTheWholeFixedRow() {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(400), PageHeight = global::OfficeIMO.OfficeLength.Points(240), MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        byte[] bytes = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
            .Table(new[] { new[] { "Header" }, new[] { "First" }, new[] { "Final" } }, style: new PdfTableStyle {
                HeaderRowCount = 1, RepeatHeaderRowCount = 1, CellPaddingY = global::OfficeIMO.OfficeLength.Points(0),
                FixedRowHeights = new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(150) },
                VerticalAlignments = new System.Collections.Generic.List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom },
                RowAllowBreakAcrossPages = new System.Collections.Generic.List<bool?> { false, false, false }
            }), new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(20), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = false, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Contains("Final", string.Join("", pdf.GetPages().Select(page => page.Text)), StringComparison.Ordinal);
        Assert.All(pdf.GetPages().SelectMany(page => page.Letters), letter => Assert.True(letter.BoundingBox.Bottom >= 25, "Table text exceeded the bottom margin."));
        SavePaddedPaginationEvidence("padded-table-fixed-row", bytes);
    }

    private static void SavePaddedPaginationEvidence(string name, byte[] bytes) {
        string? evidence = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_PDF_EVIDENCE");
        if (string.IsNullOrWhiteSpace(evidence)) return;
        Directory.CreateDirectory(evidence!);
        File.WriteAllBytes(Path.Combine(evidence!, name + ".pdf"), bytes);
    }

    [Theory]
    [InlineData(220, 240)]
    [InlineData(400, 160)]
    public void NestedPaddedHeadingPreservesTextLinksAndOneHeadingTag(double pageWidth, double pageHeight) {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(pageWidth), PageHeight = global::OfficeIMO.OfficeLength.Points(pageHeight), MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        string heading = string.Join(" ", Enumerable.Repeat("Heading text", 24)) + " FINAL-MARKER";
        var outer = new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(12), PaddingX = global::OfficeIMO.OfficeLength.Points(8), KeepTogether = false, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), Background = global::OfficeIMO.OfficeColor.FromRgb(235, 240, 248) };
        var inner = new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(8), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = false, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), BorderColor = global::OfficeIMO.OfficeColor.FromRgb(60, 90, 140) };
        byte[] bytes = PdfDocument.Create(options).TaggedPdfCatalogMarkers().Panel(panel => panel
            .Paragraph(p => p.Text("Outer lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
            .Panel(nested => nested
                .Paragraph(p => p.Text("Inner lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                .H1(heading, linkUri: "https://example.test/heading", style: new PdfHeadingStyle { FontSize = global::OfficeIMO.OfficeLength.Points(20), LineHeight = 1.15, SpacingBefore = global::OfficeIMO.OfficeLength.Points(6), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), KeepWithNext = false })
                .Paragraph(p => p.Text("Following paragraph"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }), inner), outer).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 2);
        string text = string.Join("", pdf.GetPages().Select(page => page.Text));
        Assert.Contains("FINAL-MARKER", text, StringComparison.Ordinal);
        Assert.Contains("Followingparagraph", text.Replace(" ", ""), StringComparison.Ordinal);
        foreach (var page in pdf.GetPages()) Assert.All(page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)), letter => {
            Assert.InRange(letter.BoundingBox.Top, options.MarginBottom.ToPoints() - 1, options.PageHeight.ToPoints() - options.MarginTop.ToPoints() - outer.PaddingY.ToPoints() + 1);
            Assert.InRange(letter.BoundingBox.Bottom, options.MarginBottom.ToPoints() - 1, options.PageHeight.ToPoints() - options.MarginTop.ToPoints() - outer.PaddingY.ToPoints() + 1);
        });
        string raw = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Equal(1, CountOccurrences(raw, "/Type /StructElem /S /H1 "));
        var links = ExtractLinkRectangles(raw);
        Assert.NotEmpty(links);
        Assert.All(links, link => {
            Assert.InRange(link.Y1, options.MarginBottom.ToPoints() - 1, options.PageHeight.ToPoints() - options.MarginTop.ToPoints() - outer.PaddingY.ToPoints() - inner.PaddingY.ToPoints() + 1);
            Assert.InRange(link.Y2, options.MarginBottom.ToPoints() - 1, options.PageHeight.ToPoints() - options.MarginTop.ToPoints() - outer.PaddingY.ToPoints() - inner.PaddingY.ToPoints() + 1);
        });
        string? evidence = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_PDF_EVIDENCE");
        if (!string.IsNullOrWhiteSpace(evidence)) {
            Directory.CreateDirectory(evidence!);
            File.WriteAllBytes(Path.Combine(evidence!, "padded-heading-" + pageWidth + "-" + pageHeight + ".pdf"), bytes);
        }
    }

    [Fact]
    public void PaddedParagraphMakesProgressWhenOrphanPreferenceExceedsFrameCapacity() {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(400), PageHeight = global::OfficeIMO.OfficeLength.Points(150), MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        byte[] bytes = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
            .Paragraph(p => { for (int i = 1; i <= 6; i++) { if (i > 1) p.LineBreak(); p.Text("Line" + i); } },
                style: new PdfParagraphStyle { MinimumOrphanLines = 6, MinimumWidowLines = 2, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }),
            new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(30), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = false, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.InRange(pdf.NumberOfPages, 2, 3);
        string text = string.Join("", pdf.GetPages().Select(page => page.Text));
        for (int i = 1; i <= 6; i++) Assert.Contains("Line" + i, text, StringComparison.Ordinal);
    }

    [Fact]
    public void PaddedColumnParagraphRelaxesAnImpossibleOrphanPreferenceAtFrameStart() {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(400), PageHeight = global::OfficeIMO.OfficeLength.Points(170), MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        byte[] bytes = PdfDocument.Create(options).Compose(document => document.Page(page => page.Content(content => content
            .Row(row => row.PercentColumn(100, column => column.Panel(panel => panel
                .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                .Paragraph(p => { for (int i = 1; i <= 6; i++) { if (i > 1) p.LineBreak(); p.Text("Line" + i); } },
                    style: new PdfParagraphStyle { MinimumOrphanLines = 6, MinimumWidowLines = 2, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }),
                new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(25), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = false, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })))))).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.InRange(pdf.NumberOfPages, 2, 3);
        string text = string.Join("", pdf.GetPages().Select(page => page.Text));
        for (int i = 1; i <= 6; i++) Assert.Contains("Line" + i, text, StringComparison.Ordinal);
    }

    [Fact]
    public void PaddedRowTableRepeatedHeaderDoesNotCountAsBodyProgress() {
        var first = new[] { PdfTextRun.Normal("First") };
        var kept = new[] { PdfTextRun.Normal("Kept1\nKept2\nKept3\nKept4") };
        var body = new PdfTableCell(first.Concat(kept), new[] {
            new PdfTableCellParagraph(first, fontSize: OfficeLength.Points(12), lineSpacing: PdfLineSpacing.Exact(OfficeLength.Points(15)), widowControl: false),
            new PdfTableCellParagraph(kept, fontSize: OfficeLength.Points(12), lineSpacing: PdfLineSpacing.Exact(OfficeLength.Points(15)), widowControl: false, keepTogether: true)
        });
        var tableStyle = new PdfTableStyle {
            HeaderRowCount = 1, RepeatHeaderRowCount = 1,
            MinimumBodyRowsOnFirstPage = 0, MinimumBodyRowsOnLastPage = 0,
            CellPaddingX = OfficeLength.Points(0), CellPaddingY = OfficeLength.Points(0), BorderWidth = OfficeLength.Points(0),
            FontSize = OfficeLength.Points(12), LineHeight = 1.25, SpacingBefore = OfficeLength.Points(0), SpacingAfter = OfficeLength.Points(0)
        };
        PdfDocument document = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = OfficeLength.Points(220), PageHeight = OfficeLength.Points(120), MarginLeft = OfficeLength.Points(20), MarginRight = OfficeLength.Points(20),
            MarginTop = OfficeLength.Points(20), MarginBottom = OfficeLength.Points(20), DefaultFontSize = OfficeLength.Points(12), MaxGeneratedPages = 3
        }).Compose(root => root.Page(page => page.Content(content => content.Row(row =>
            row.PercentColumn(100, column => column.Panel(panel => panel.Table(new[] {
                new[] { new PdfTableCell("Header") }, new[] { body }
            }, style: tableStyle), new PdfPanelStyle { PaddingX = OfficeLength.Points(0), PaddingY = OfficeLength.Points(6), KeepTogether = false }))))));

        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => document.ToBytes());
        Assert.Contains("progress", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OversizedPaddedTableLineCannotOverflowThePage() {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(400), PageHeight = global::OfficeIMO.OfficeLength.Points(150), MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        PdfDocument document = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
            .Table(new[] { new[] { "A" } }, style: new PdfTableStyle { FontSize = global::OfficeIMO.OfficeLength.Points(60), HeaderRowCount = 0, CellPaddingY = global::OfficeIMO.OfficeLength.Points(0) }),
            new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(30), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = false, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TableHeightRequirementsMustFitThePaddedFrame(bool fixedHeight) {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(400), PageHeight = global::OfficeIMO.OfficeLength.Points(150), MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        var style = new PdfTableStyle { HeaderRowCount = 0, CellPaddingY = global::OfficeIMO.OfficeLength.Points(0) };
        if (fixedHeight) style.FixedRowHeights = new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(75) };
        else style.MinRowHeight = global::OfficeIMO.OfficeLength.Points(75);
        PdfDocument document = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
            .Table(new[] { new[] { "A" } }, style: style),
            new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(30), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = false, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Fact]
    public void NestedPaddedHeadingAccountsForCumulativePaddingAndRequiredSpacing() {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(400), PageHeight = global::OfficeIMO.OfficeLength.Points(150), MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        PdfDocument document = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
            .Panel(nested => nested.Paragraph(p => p.Text("Inner"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                .H1("A", style: new PdfHeadingStyle { FontSize = global::OfficeIMO.OfficeLength.Points(50), LineHeight = 1, SpacingBefore = global::OfficeIMO.OfficeLength.Points(25), ApplySpacingBeforeAtTop = true, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), KeepWithNext = false }),
                new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(20), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = false, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }),
            new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(10), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = false, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Fact]
    public void ParagraphAfterContentRejectsALineThatCannotFitThePaddedFrame() {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(400), PageHeight = global::OfficeIMO.OfficeLength.Points(150), MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        PdfDocument document = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
            .Paragraph(p => p.FontSize(global::OfficeIMO.OfficeLength.Points(55)).Text("A"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }),
            new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(30), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = false, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Fact]
    public void OversizedPaddedListCannotSilentlyDiscardText() {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(400), PageHeight = global::OfficeIMO.OfficeLength.Points(150), MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        PdfDocument document = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
            .RichBullets(new[] { PdfListItem.Rich(new[] { PdfTextRun.Normal("A") }) }, style: new PdfListStyle { FontSize = global::OfficeIMO.OfficeLength.Points(75), LineHeight = 1 }),
            new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(30), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = false, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Fact]
    public void OversizedPaddedRowHeadingIsRejectedAsUnrenderableInput() {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(400), PageHeight = global::OfficeIMO.OfficeLength.Points(150), MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        PdfDocument document = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
            .Row(row => row.PercentColumn(100, column => column.H1("A", style: new PdfHeadingStyle { FontSize = global::OfficeIMO.OfficeLength.Points(75), LineHeight = 1, SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), KeepWithNext = false }))),
            new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(30), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = false, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Fact]
    public void HeadingAfterContentRejectsALineThatCannotFitThePaddedFrame() {
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(400), PageHeight = global::OfficeIMO.OfficeLength.Points(150), MarginLeft = global::OfficeIMO.OfficeLength.Points(25), MarginRight = global::OfficeIMO.OfficeLength.Points(25), MarginTop = global::OfficeIMO.OfficeLength.Points(25), MarginBottom = global::OfficeIMO.OfficeLength.Points(25), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) };
        PdfDocument document = PdfDocument.Create(options).Panel(panel => panel
            .Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle { SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
            .H1("A", style: new PdfHeadingStyle { FontSize = global::OfficeIMO.OfficeLength.Points(75), LineHeight = 1, SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), KeepWithNext = false }),
            new PdfPanelStyle { PaddingY = global::OfficeIMO.OfficeLength.Points(30), PaddingX = global::OfficeIMO.OfficeLength.Points(5), KeepTogether = false, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }
}
