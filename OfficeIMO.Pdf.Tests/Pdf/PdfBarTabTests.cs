using System;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using PdfPigPage = UglyToad.PdfPig.Content.Page;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfBarTabTests {
    [Theory]
    [InlineData(PdfAlign.Left)]
    [InlineData(PdfAlign.Center)]
    [InlineData(PdfAlign.Right)]
    [InlineData(PdfAlign.Justify)]
    public void BarRulesKeepTextStopsAndGlyphFlowIndependentOfParagraphAlignment(PdfAlign align) {
        byte[] Render(bool bar) {
            PdfParagraphStyle style = ParagraphStyle();
            style.LeftIndent = global::OfficeIMO.OfficeLength.Points(36);
            style.FirstLineIndent = global::OfficeIMO.OfficeLength.Points(-18);
            style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(6);
            style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(8);
            style.AddTabStop(global::OfficeIMO.OfficeLength.Points(90));
            if (bar) style.AddTabStop(global::OfficeIMO.OfficeLength.Points(36), PdfTabAlignment.Bar);
            return PdfDocument.Create(builder => builder.Content(content => content
                .Paragraph(paragraph => paragraph.Text("ANCHOR"), style: ParagraphStyle())
                .Paragraph(paragraph => paragraph.Text("A\tFIELD\nNEXT"), align, style: style)
                .Paragraph(paragraph => paragraph.Text("TAIL"), style: ParagraphStyle())), Options()).ToBytes();
        }

        using var control = PdfPigDocument.Open(Render(false));
        using var actual = PdfPigDocument.Open(Render(true));
        AssertSameGlyphFlow(control, actual);
        var bars = VerticalRules(actual.GetPage(1));
        Assert.Equal(2, bars.Length);
        Assert.All(bars, bar => Assert.Equal(102D, bar.X, 3));
        Assert.Equal(192D, bars[0].Top, 3);
        Assert.Equal(24D, bars[0].Height, 3);
        Assert.Equal(168D, bars[1].Top, 3);
        Assert.Equal(26D, bars[1].Height, 3);
    }

    [Fact]
    public void BarAppearanceAndEmptyParagraphAreCapturedWithoutInventingText() {
        PdfParagraphStyle style = ParagraphStyle();
        style.LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(20));
        style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(6); // Ordinary top-of-page flow suppresses this spacing.
        style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(7);
        PdfTabStop stop = style.AddTabStop(global::OfficeIMO.OfficeLength.Inches(7D / 12D), PdfTabAlignment.Bar);
        stop.BarWidth = global::OfficeIMO.OfficeLength.Emus(1778);
        stop.BarColor = global::OfficeIMO.OfficeColor.FromRgb(0, 0, 255);
        var document = PdfDocument.Create(builder => builder.Content(content => content
            .Paragraph(_ => { }, style: style)
            .Paragraph(paragraph => paragraph.Text("TAIL"), style: ParagraphStyle())), Options());
        stop.Position = global::OfficeIMO.OfficeLength.Points(99);
        stop.BarWidth = global::OfficeIMO.OfficeLength.Points(3);
        stop.BarColor = global::OfficeIMO.OfficeColor.Black;
        style.ClearTabStops();
        byte[] bytes = document.ToBytes();

        using var actual = PdfPigDocument.Open(bytes);
        PdfPigPage page = actual.GetPage(1);
        Assert.Equal("TAIL", string.Concat(page.Letters.Select(letter => letter.Value)));
        var bar = Assert.Single(VerticalRules(page));
        Assert.Equal(72D, bar.X, 3);
        Assert.Equal(210D, bar.Top, 3);
        Assert.Equal(27D, bar.Height, 3);
        string literal = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.14 w\n", literal, StringComparison.Ordinal);
        Assert.Contains("0 0 1 RG\n", literal, StringComparison.Ordinal);
        Assert.Contains("2 J\n", literal, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("canvas")]
    public void LinkedCellRulesKeepParagraphOriginAndDoNotAbsorbFollowingBeforeSpacing(string route) {
        byte[] Render(bool bar) {
            var runs = new[] { new PdfTextRun("A\nB", linkUri: "https://example.com/inline") };
            var next = new[] { PdfTextRun.Normal("C") };
            var stops = bar ? new[] { new PdfTabStop(global::OfficeIMO.OfficeLength.Points(36), PdfTabAlignment.Bar) } : Array.Empty<PdfTabStop>();
            var paragraphs = new[] {
                new PdfTableCellParagraph(runs, spacingAfter: global::OfficeIMO.OfficeLength.Points(8), leftIndent: global::OfficeIMO.OfficeLength.Points(12), firstLineIndent: global::OfficeIMO.OfficeLength.Points(-6),
                    fontSize: global::OfficeIMO.OfficeLength.Points(12), lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(18)), tabStops: stops),
                new PdfTableCellParagraph(next, spacingBefore: global::OfficeIMO.OfficeLength.Points(6), leftIndent: global::OfficeIMO.OfficeLength.Points(12), fontSize: global::OfficeIMO.OfficeLength.Points(12),
                    lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(18)), tabStops: stops)
            };
            var cell = new PdfTableCell(runs.Concat(next), paragraphs, linkUri: "https://example.com/cell");
            var table = new PdfTableStyle {
                HeaderRowCount = 0, CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0), BorderWidth = global::OfficeIMO.OfficeLength.Points(0),
                FontSize = global::OfficeIMO.OfficeLength.Points(12), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            };
            PdfTableCell[][] rows = { new[] { cell } };
            return PdfDocument.Create(builder => builder.Content(content => {
                if (route == "flow") content.Table(rows, style: table);
                else if (route == "column") content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: table)));
                else content.Canvas(canvas => canvas.Table(rows, global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(120), table));
            }), Options()).ToBytes();
        }

        using var control = PdfPigDocument.Open(Render(false));
        using var actual = PdfPigDocument.Open(Render(true));
        AssertSameGlyphFlow(control, actual);
        var bars = VerticalRules(actual.GetPage(1));
        Assert.Equal(3, bars.Length);
        Assert.All(bars, bar => Assert.Equal(78D, bar.X, 3));
        Assert.Equal(18D, bars[0].Height, 3);
        Assert.Equal(26D, bars[1].Height, 3);
        Assert.Equal(24D, bars[2].Height, 3);
        Assert.Equal(bars[1].Top - bars[1].Height, bars[2].Top, 3);
        Assert.NotEmpty(actual.GetPage(1).GetHyperlinks());
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("canvas")]
    public void FixedCellRulesTrimAfterSpacingAtTheActualCellBoundary(string route) {
        var runs = new[] { PdfTextRun.Normal("A") };
        var paragraph = new PdfTableCellParagraph(runs, spacingAfter: global::OfficeIMO.OfficeLength.Points(10), fontSize: global::OfficeIMO.OfficeLength.Points(10),
            lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(12)), tabStops: new[] { new PdfTabStop(global::OfficeIMO.OfficeLength.Points(36), PdfTabAlignment.Bar) });
        PdfTableCell[][] rows = { new[] { new PdfTableCell(runs, new[] { paragraph }) }, new[] { new PdfTableCell("B") } };
        var table = new PdfTableStyle {
            HeaderRowCount = 0, CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0), BorderWidth = global::OfficeIMO.OfficeLength.Points(0),
            FontSize = global::OfficeIMO.OfficeLength.Points(10), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0),
            FixedRowHeights = new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(16), global::OfficeIMO.OfficeLength.Points(16) }
        };
        byte[] bytes = PdfDocument.Create(builder => builder.Content(content => {
            if (route == "flow") content.Table(rows, style: table);
            else content.Canvas(canvas => canvas.Table(rows, global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(32), table));
        }), Options()).ToBytes();
        using var actual = PdfPigDocument.Open(bytes);
        PdfPigPage page = Assert.Single(actual.GetPages());
        Assert.Equal("AB", string.Concat(page.Letters.Select(letter => letter.Value)));
        var bar = Assert.Single(VerticalRules(page));
        Assert.Equal(66D, bar.X, 3);
        Assert.Equal(16D, bar.Height, 3);
        Assert.Equal(210D, bar.Top, 3);
    }

    [Fact]
    public void PageFragmentsPaintOnlyTheirOwnedLinesAndTerminalAfterSpacing() {
        PdfCreateOptions options = Options();
        options.PageHeight = global::OfficeIMO.OfficeLength.Points(120);
        options.MarginTop = options.MarginBottom = global::OfficeIMO.OfficeLength.Points(20);
        byte[] Render(bool bar) {
            PdfParagraphStyle style = ParagraphStyle();
            style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(6);
            style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(8);
            if (bar) style.AddTabStop(global::OfficeIMO.OfficeLength.Points(36), PdfTabAlignment.Bar);
            return PdfDocument.Create(builder => builder.Content(content => content
                .Paragraph(paragraph => paragraph.Text("ANCHOR"), style: ParagraphStyle())
                .Paragraph(paragraph => paragraph.Text("A\nB\nC\nD\nE\nF"), style: style)), options).ToBytes();
        }

        using var control = PdfPigDocument.Open(Render(false));
        using var actual = PdfPigDocument.Open(Render(true));
        AssertSameGlyphFlow(control, actual);
        Assert.Equal(2, actual.NumberOfPages);
        var first = VerticalRules(actual.GetPage(1));
        var last = VerticalRules(actual.GetPage(2));
        Assert.Equal(3, first.Length);
        Assert.Equal(3, last.Length);
        Assert.Equal(24D, first[0].Height, 3);
        Assert.All(first.Skip(1), bar => Assert.Equal(18D, bar.Height, 3));
        Assert.All(last.Take(2), bar => Assert.Equal(18D, bar.Height, 3));
        Assert.Equal(26D, last[2].Height, 3);
        Assert.All(first.Concat(last), bar => Assert.Equal(66D, bar.X, 3));
    }

    [Fact]
    public void NaturalBarIntervalsFollowEachLinesFontBoxWithoutChangingTextFlow() {
        byte[] Render(bool bar) {
            PdfParagraphStyle style = ParagraphStyle();
            style.LineSpacing = null;
            style.LineHeight = 1.2D;
            style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(7);
            if (bar) style.AddTabStop(global::OfficeIMO.OfficeLength.Points(36), PdfTabAlignment.Bar);
            return PdfDocument.Create(builder => builder.Content(content => content
                .Paragraph(paragraph => paragraph.FontSize(global::OfficeIMO.OfficeLength.Points(24)).Text("A").FontSize(global::OfficeIMO.OfficeLength.Points(12)).Text("\nB"), style: style)
                .Paragraph(paragraph => paragraph.Text("TAIL"), style: ParagraphStyle())), Options()).ToBytes();
        }

        using var control = PdfPigDocument.Open(Render(false));
        using var actual = PdfPigDocument.Open(Render(true));
        AssertSameGlyphFlow(control, actual);
        var bars = VerticalRules(actual.GetPage(1));
        Assert.Equal(2, bars.Length);
        Assert.Equal(28.8D, bars[0].Height, 3);
        Assert.Equal(21.4D, bars[1].Height, 3);
        Assert.Equal(bars[0].Height, bars[0].Top - bars[1].Top, 3);
    }

    [Fact]
    public void WidowCarriedFloatSideRulesReturnToTheDestinationParagraphAnchor() {
        byte[] Render(bool bar) {
            PdfCreateOptions options = Options();
            options.PageWidth = global::OfficeIMO.OfficeLength.Points(400);
            options.PageHeight = global::OfficeIMO.OfficeLength.Points(240);
            options.MarginLeft = options.MarginRight = options.MarginTop = options.MarginBottom = global::OfficeIMO.OfficeLength.Points(40);
            PdfParagraphStyle style = ParagraphStyle();
            style.LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(14.4D));
            style.WidowControl = true;
            style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(8);
            if (bar) style.AddTabStop(global::OfficeIMO.OfficeLength.Points(72), PdfTabAlignment.Bar);
            return PdfDocument.Create(builder => builder.Content(content => content
                .Table(new[] { new[] { "T" } }, style: new PdfTableStyle {
                    HeaderRowCount = 0, FontSize = global::OfficeIMO.OfficeLength.Points(2), CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0),
                    ColumnWidths = new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(100) },
                    MinRowHeight = global::OfficeIMO.OfficeLength.Points(145), BorderColor = null, Position = new PdfTablePosition()
                })
                .Paragraph(paragraph => paragraph.Text(string.Join("\n", Enumerable.Range(0, 12)
                    .Select(index => "Line" + index.ToString("D2")))), style: style)), options).ToBytes();
        }

        using var control = PdfPigDocument.Open(Render(false));
        using var actual = PdfPigDocument.Open(Render(true));
        AssertSameGlyphFlow(control, actual);
        Assert.Equal(2, actual.NumberOfPages);
        var bars = actual.GetPages().SelectMany(VerticalRules).ToArray();
        Assert.Equal(12, bars.Length);
        Assert.All(bars, bar => Assert.Equal(112D, bar.X, 3));
        Assert.All(bars.Take(11), bar => Assert.Equal(14.4D, bar.Height, 3));
        Assert.Equal(22.4D, bars[11].Height, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HeadingRulesUseTheSharedPainterAcrossFlowAndColumnPaths(bool column) {
        byte[] Render(bool bar) {
            var style = new PdfHeadingStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(12), Bold = false, KeepWithNext = false,
                LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(18)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(6), SpacingAfter = global::OfficeIMO.OfficeLength.Points(7),
                ApplySpacingBeforeAtTop = true
            };
            if (bar) style.AddTabStop(global::OfficeIMO.OfficeLength.Points(36), PdfTabAlignment.Bar);
            var document = PdfDocument.Create(builder => builder.Content(content => {
                if (column) content.Row(row => row.PercentColumn(100, frame => frame.Heading(2, "A\nB", style: style)));
                else content.Heading(2, "A\nB", style: style);
            }), Options());
            style.ClearTabStops();
            return document.ToBytes();
        }

        using var control = PdfPigDocument.Open(Render(false));
        using var actual = PdfPigDocument.Open(Render(true));
        AssertSameGlyphFlow(control, actual);
        var bars = VerticalRules(actual.GetPage(1));
        Assert.Equal(2, bars.Length);
        Assert.All(bars, bar => Assert.Equal(66D, bar.X, 3));
        Assert.Equal(210D, bars[0].Top, 3);
        Assert.Equal(24D, bars[0].Height, 3);
        Assert.Equal(25D, bars[1].Height, 3);
    }

    [Fact]
    public void NoteContinuationRetainsBarAnchorsAndOnlyTheLastAfterExtent() {
        PdfCreateOptions options = Options();
        options.PageHeight = global::OfficeIMO.OfficeLength.Points(160);
        options.MarginTop = options.MarginBottom = global::OfficeIMO.OfficeLength.Points(20);
        byte[] Render(bool bar) {
            PdfParagraphStyle style = ParagraphStyle();
            style.LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(14));
            style.LeftIndent = global::OfficeIMO.OfficeLength.Points(12);
            style.FirstLineIndent = global::OfficeIMO.OfficeLength.Points(-6);
            style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(6);
            style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(5);
            if (bar) style.AddTabStop(global::OfficeIMO.OfficeLength.Points(30), PdfTabAlignment.Bar);
            return PdfDocument.Create(builder => builder.Content(content => {
                PdfFootnoteDefinition note = content.DefineFootnote(body => body.Paragraph(paragraph => paragraph.Text(
                    string.Join("\n", Enumerable.Range(0, 13).Select(index => "Note" + index.ToString("D2")))), style: style));
                content.Paragraph(paragraph => paragraph.Runs(new[] { PdfTextRun.Normal("REF").WithFootnoteReference(note) }), style: ParagraphStyle());
            }), options).ToBytes();
        }

        using var control = PdfPigDocument.Open(Render(false));
        using var actual = PdfPigDocument.Open(Render(true));
        AssertSameGlyphFlow(control, actual);
        Assert.True(actual.NumberOfPages > 1);
        var bars = actual.GetPages().SelectMany(VerticalRules).ToArray();
        Assert.Equal(13, bars.Length);
        Assert.Equal(20D, bars[0].Height, 3);
        Assert.All(bars.Skip(1).Take(11), bar => Assert.Equal(14D, bar.Height, 3));
        Assert.Equal(19D, bars[12].Height, 3);
        Assert.All(bars, bar => Assert.Equal(72D, bar.X, 3));
    }

    [Fact]
    public void BarAlignmentIsParagraphOnlyAndAppearanceIsCloned() {
        Assert.Throws<ArgumentException>(() => PdfTextRun.Tab(alignment: PdfTabAlignment.Bar));
        Assert.Throws<ArgumentException>(() => new PdfTextRun("\t", tabAlignment: PdfTabAlignment.Bar));
        var stop = new PdfTabStop(global::OfficeIMO.OfficeLength.Points(-12), PdfTabAlignment.Bar) { BarWidth = global::OfficeIMO.OfficeLength.Points(2), BarColor = global::OfficeIMO.OfficeColor.FromRgb(12, 34, 56) };
        PdfTabStop clone = stop.Clone();
        stop.BarWidth = global::OfficeIMO.OfficeLength.Points(3);
        Assert.Equal(-12D, clone.Position.ToPoints());
        Assert.Equal(PdfTabAlignment.Bar, clone.Alignment);
        Assert.Equal(2D, clone.BarWidth.ToPoints());
        Assert.Equal(global::OfficeIMO.OfficeColor.FromRgb(12, 34, 56), clone.BarColor);
    }

    [Theory]
    [InlineData(0D)]
    [InlineData(-1D)]
    public void BarWidthRejectsNonpositiveStrokeDimensions(double width) {
        var stop = new PdfTabStop(global::OfficeIMO.OfficeLength.Points(36), PdfTabAlignment.Bar);
        Assert.Throws<ArgumentOutOfRangeException>(() => stop.BarWidth = global::OfficeIMO.OfficeLength.Points(width));
    }

    [Theory]
    [InlineData("flow-viewport")]
    [InlineData("canvas-viewport")]
    [InlineData("flow-partial")]
    public void EmptyCellGlyphOmissionRetainsItsIndependentParagraphRule(string route) {
        var mark = new PdfTextRun(string.Empty, fontSize: global::OfficeIMO.OfficeLength.Points(12)) { IsLineBoxStrut = true };
        var paragraph = new PdfTableCellParagraph(new[] { mark }, fontSize: global::OfficeIMO.OfficeLength.Points(12),
            lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(18)), tabStops: new[] { new PdfTabStop(global::OfficeIMO.OfficeLength.Points(36), PdfTabAlignment.Bar) });
        var cell = new PdfTableCell(new[] { mark }, new[] { paragraph });
        if (route != "flow-partial") cell = cell.WithViewport(new PdfTableCellViewport(global::OfficeIMO.OfficeLength.Points(100), global::OfficeIMO.OfficeLength.Points(36), global::OfficeIMO.OfficeLength.Points(100), global::OfficeIMO.OfficeLength.Points(36)));
        var table = BarOnlyCellStyle(100, 36);
        table.PreservePartialCellLines = route == "flow-partial";
        byte[] bytes = PdfDocument.Create(builder => builder.Content(content => {
            PdfTableCell[][] rows = { new[] { cell } };
            if (route == "canvas-viewport") content.Canvas(canvas => canvas.Table(rows, global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(100), global::OfficeIMO.OfficeLength.Points(36), table));
            else content.Table(rows, style: table);
        }), Options()).ToBytes();

        using var actual = PdfPigDocument.Open(bytes);
        PdfPigPage page = Assert.Single(actual.GetPages());
        Assert.Empty(page.Letters);
        var bar = Assert.Single(VerticalRules(page));
        Assert.Equal(66D, bar.X, 3);
        Assert.Equal(210D, bar.Top, 3);
        Assert.Equal(18D, bar.Height, 3);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("canvas")]
    public void HorizontalViewportCanOmitEveryGlyphWhileKeepingAnInFrameBar(string route) {
        var runs = new[] { PdfTextRun.Normal("A") };
        var paragraph = new PdfTableCellParagraph(runs, fontSize: global::OfficeIMO.OfficeLength.Points(12),
            lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(18)), tabStops: new[] { new PdfTabStop(global::OfficeIMO.OfficeLength.Points(36), PdfTabAlignment.Bar) });
        var cell = new PdfTableCell(runs, new[] { paragraph })
            .WithViewport(new PdfTableCellViewport(global::OfficeIMO.OfficeLength.Points(100), global::OfficeIMO.OfficeLength.Points(36), global::OfficeIMO.OfficeLength.Points(60), global::OfficeIMO.OfficeLength.Points(36), offsetX: global::OfficeIMO.OfficeLength.Points(24)));
        var table = BarOnlyCellStyle(60, 36);
        byte[] bytes = PdfDocument.Create(builder => builder.Content(content => {
            PdfTableCell[][] rows = { new[] { cell } };
            if (route == "canvas") content.Canvas(canvas => canvas.Table(rows, global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(30), global::OfficeIMO.OfficeLength.Points(60), global::OfficeIMO.OfficeLength.Points(36), table));
            else content.Table(rows, style: table);
        }), Options()).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });

        using var actual = PdfPigDocument.Open(bytes);
        PdfPigPage page = Assert.Single(actual.GetPages());
        Assert.Empty(page.Letters);
        var bar = Assert.Single(VerticalRules(page));
        // Full-cell origin30-offset24 plus source stop36; the glyph lies left
        // of the visible interval, while the stationary rule remains inside.
        Assert.Equal(42D, bar.X, 3);
        Assert.Equal(210D, bar.Top, 3);
        Assert.Equal(18D, bar.Height, 3);
    }

    [Theory]
    [InlineData(90, 30D, 146D)]
    [InlineData(-90, 48D, 174D)]
    public void RotatedEmptyCellKeepsTheRuleAfterItsZeroWidthGlyphLineIsOmitted(int rotation, double left, double y) {
        var mark = new PdfTextRun(string.Empty, fontSize: global::OfficeIMO.OfficeLength.Points(12)) { IsLineBoxStrut = true };
        var paragraph = new PdfTableCellParagraph(new[] { mark }, fontSize: global::OfficeIMO.OfficeLength.Points(12),
            lineSpacing: PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(18)), tabStops: new[] { new PdfTabStop(global::OfficeIMO.OfficeLength.Points(36), PdfTabAlignment.Bar) });
        var cell = new PdfTableCell(new[] { mark }, new[] { paragraph }, textRotation: rotation);
        byte[] bytes = PdfDocument.Create(builder => builder.Content(content => content.Table(
            new[] { new[] { cell } }, style: BarOnlyCellStyle(36, 100))), Options()).ToBytes();

        using var actual = PdfPigDocument.Open(bytes);
        PdfPigPage page = Assert.Single(actual.GetPages());
        Assert.Empty(page.Letters);
        var bar = Assert.Single(page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle()),
            bounds => bounds.HasValue && Math.Abs(bounds.Value.Height) < .001D && bounds.Value.Width > 1D);
        Assert.Equal(left, bar!.Value.Left, 3);
        Assert.Equal(y, bar.Value.Bottom, 3);
        Assert.Equal(18D, bar.Value.Width, 3);
    }

    private static PdfTableStyle BarOnlyCellStyle(double width, double height) => new() {
        HeaderRowCount = 0, FontSize = global::OfficeIMO.OfficeLength.Points(12), CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0),
        BorderWidth = global::OfficeIMO.OfficeLength.Points(0), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), MaxWidth = global::OfficeIMO.OfficeLength.Points(width), PreserveWidth = true,
        ColumnWidths = new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(width) },
        FixedRowHeights = new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(height) }
    };

    private static PdfCreateOptions Options() => new() {
        PageWidth = global::OfficeIMO.OfficeLength.Points(360), PageHeight = global::OfficeIMO.OfficeLength.Points(240), MarginLeft = global::OfficeIMO.OfficeLength.Points(30), MarginRight = global::OfficeIMO.OfficeLength.Points(30),
        MarginTop = global::OfficeIMO.OfficeLength.Points(30), MarginBottom = global::OfficeIMO.OfficeLength.Points(30), DefaultFont = PdfStandardFont.Helvetica,
        DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12), CompressContentStreams = false
    };

    private static PdfParagraphStyle ParagraphStyle() => new() {
        FontSize = global::OfficeIMO.OfficeLength.Points(12), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(18)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0),
        WidowControl = false, KeepTogether = false, KeepWithNext = false
    };

    private static (double X, double Top, double Height)[] VerticalRules(PdfPigPage page) => page.Paths
        .Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
        .Where(bounds => bounds.HasValue && Math.Abs(bounds.Value.Width) < .001D && bounds.Value.Height > 1D)
        .Select(bounds => (bounds!.Value.Left, bounds.Value.Top, bounds.Value.Height))
        .OrderByDescending(bounds => bounds.Top).ToArray();

    private static void AssertSameGlyphFlow(PdfPigDocument expected, PdfPigDocument actual) {
        Assert.Equal(expected.NumberOfPages, actual.NumberOfPages);
        for (int pageIndex = 1; pageIndex <= expected.NumberOfPages; pageIndex++) {
            var expectedLetters = expected.GetPage(pageIndex).Letters;
            var actualLetters = actual.GetPage(pageIndex).Letters;
            Assert.Equal(expectedLetters.Count, actualLetters.Count);
            for (int index = 0; index < expectedLetters.Count; index++) {
                var left = expectedLetters[index];
                var right = actualLetters[index];
                Assert.Equal(left.Value, right.Value);
                Assert.Equal(left.FontName, right.FontName);
                Assert.Equal(left.FontSize, right.FontSize, 3);
                Assert.Equal(left.StartBaseLine.X, right.StartBaseLine.X, 3);
                Assert.Equal(left.StartBaseLine.Y, right.StartBaseLine.Y, 3);
                Assert.Equal(left.EndBaseLine.X, right.EndBaseLine.X, 3);
                Assert.Equal(left.EndBaseLine.Y, right.EndBaseLine.Y, 3);
            }
        }
    }
}
