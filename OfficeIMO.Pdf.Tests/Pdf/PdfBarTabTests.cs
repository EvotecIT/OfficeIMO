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
            style.LeftIndent = 36;
            style.FirstLineIndent = -18;
            style.SpacingBefore = 6;
            style.SpacingAfter = 8;
            style.AddTabStop(90);
            if (bar) style.AddTabStop(36, PdfTabAlignment.Bar);
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
        style.LineSpacing = PdfLineSpacing.Exactly(20);
        style.SpacingBefore = 6; // Ordinary top-of-page flow suppresses this spacing.
        style.SpacingAfter = 7;
        PdfTabStop stop = style.AddTabStop(42, PdfTabAlignment.Bar);
        stop.BarWidth = .14D;
        stop.BarColor = PdfColor.FromRgb(0, 0, 255);
        var document = PdfDocument.Create(builder => builder.Content(content => content
            .Paragraph(_ => { }, style: style)
            .Paragraph(paragraph => paragraph.Text("TAIL"), style: ParagraphStyle())), Options());
        stop.Position = 99;
        stop.BarWidth = 3;
        stop.BarColor = PdfColor.Black;
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
            var stops = bar ? new[] { new PdfTabStop(36, PdfTabAlignment.Bar) } : Array.Empty<PdfTabStop>();
            var paragraphs = new[] {
                new PdfTableCellParagraph(runs, spacingAfter: 8, leftIndent: 12, firstLineIndent: -6,
                    fontSize: 12, lineSpacing: PdfLineSpacing.Exactly(18), tabStops: stops),
                new PdfTableCellParagraph(next, spacingBefore: 6, leftIndent: 12, fontSize: 12,
                    lineSpacing: PdfLineSpacing.Exactly(18), tabStops: stops)
            };
            var cell = new PdfTableCell(runs.Concat(next), paragraphs, linkUri: "https://example.com/cell");
            var table = new PdfTableStyle {
                HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0, BorderWidth = 0,
                FontSize = 12, SpacingBefore = 0, SpacingAfter = 0
            };
            PdfTableCell[][] rows = { new[] { cell } };
            return PdfDocument.Create(builder => builder.Content(content => {
                if (route == "flow") content.Table(rows, style: table);
                else if (route == "column") content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: table)));
                else content.Canvas(canvas => canvas.Table(rows, 30, 30, 300, 120, table));
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
        var paragraph = new PdfTableCellParagraph(runs, spacingAfter: 10, fontSize: 10,
            lineSpacing: PdfLineSpacing.Exactly(12), tabStops: new[] { new PdfTabStop(36, PdfTabAlignment.Bar) });
        PdfTableCell[][] rows = { new[] { new PdfTableCell(runs, new[] { paragraph }) }, new[] { new PdfTableCell("B") } };
        var table = new PdfTableStyle {
            HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0, BorderWidth = 0,
            FontSize = 10, SpacingBefore = 0, SpacingAfter = 0,
            FixedRowHeights = new System.Collections.Generic.List<double?> { 16, 16 }
        };
        byte[] bytes = PdfDocument.Create(builder => builder.Content(content => {
            if (route == "flow") content.Table(rows, style: table);
            else content.Canvas(canvas => canvas.Table(rows, 30, 30, 300, 32, table));
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
        PdfOptions options = Options();
        options.PageHeight = 120;
        options.MarginTop = options.MarginBottom = 20;
        byte[] Render(bool bar) {
            PdfParagraphStyle style = ParagraphStyle();
            style.SpacingBefore = 6;
            style.SpacingAfter = 8;
            if (bar) style.AddTabStop(36, PdfTabAlignment.Bar);
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
            style.SpacingAfter = 7;
            if (bar) style.AddTabStop(36, PdfTabAlignment.Bar);
            return PdfDocument.Create(builder => builder.Content(content => content
                .Paragraph(paragraph => paragraph.FontSize(24).Text("A").FontSize(12).Text("\nB"), style: style)
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
            PdfOptions options = Options();
            options.PageWidth = 400;
            options.PageHeight = 240;
            options.MarginLeft = options.MarginRight = options.MarginTop = options.MarginBottom = 40;
            PdfParagraphStyle style = ParagraphStyle();
            style.LineSpacing = PdfLineSpacing.Exactly(14.4D);
            style.WidowControl = true;
            style.SpacingAfter = 8;
            if (bar) style.AddTabStop(72, PdfTabAlignment.Bar);
            return PdfDocument.Create(builder => builder.Content(content => content
                .Table(new[] { new[] { "T" } }, style: new PdfTableStyle {
                    HeaderRowCount = 0, FontSize = 2, CellPaddingX = 0, CellPaddingY = 0,
                    ColumnWidthPoints = new System.Collections.Generic.List<double?> { 100 },
                    MinRowHeight = 145, BorderColor = null, Position = new PdfTablePosition()
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
                FontSize = 12, Bold = false, KeepWithNext = false,
                LineSpacing = PdfLineSpacing.Exactly(18), SpacingBefore = 6, SpacingAfter = 7,
                ApplySpacingBeforeAtTop = true
            };
            if (bar) style.AddTabStop(36, PdfTabAlignment.Bar);
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
        PdfOptions options = Options();
        options.PageHeight = 160;
        options.MarginTop = options.MarginBottom = 20;
        byte[] Render(bool bar) {
            PdfParagraphStyle style = ParagraphStyle();
            style.LineSpacing = PdfLineSpacing.Exactly(14);
            style.LeftIndent = 12;
            style.FirstLineIndent = -6;
            style.SpacingBefore = 6;
            style.SpacingAfter = 5;
            if (bar) style.AddTabStop(30, PdfTabAlignment.Bar);
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
        var stop = new PdfTabStop(-12, PdfTabAlignment.Bar) { BarWidth = 2, BarColor = PdfColor.FromRgb(12, 34, 56) };
        PdfTabStop clone = stop.Clone();
        stop.BarWidth = 3;
        Assert.Equal(-12D, clone.Position);
        Assert.Equal(PdfTabAlignment.Bar, clone.Alignment);
        Assert.Equal(2D, clone.BarWidth);
        Assert.Equal(PdfColor.FromRgb(12, 34, 56), clone.BarColor);
    }

    [Theory]
    [InlineData(0D)]
    [InlineData(-1D)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void BarWidthRejectsNonpositiveOrNonfiniteStrokeDimensions(double width) {
        var stop = new PdfTabStop(36, PdfTabAlignment.Bar);
        Assert.Throws<ArgumentOutOfRangeException>(() => stop.BarWidth = width);
    }

    [Theory]
    [InlineData("flow-viewport")]
    [InlineData("canvas-viewport")]
    [InlineData("flow-partial")]
    public void EmptyCellGlyphOmissionRetainsItsIndependentParagraphRule(string route) {
        var mark = new PdfTextRun(string.Empty, fontSize: 12) { IsLineBoxStrut = true };
        var paragraph = new PdfTableCellParagraph(new[] { mark }, fontSize: 12,
            lineSpacing: PdfLineSpacing.Exactly(18), tabStops: new[] { new PdfTabStop(36, PdfTabAlignment.Bar) });
        var cell = new PdfTableCell(new[] { mark }, new[] { paragraph });
        if (route != "flow-partial") cell = cell.WithViewport(new PdfTableCellViewport(100, 36, 100, 36));
        var table = BarOnlyCellStyle(100, 36);
        table.PreservePartialCellLines = route == "flow-partial";
        byte[] bytes = PdfDocument.Create(builder => builder.Content(content => {
            PdfTableCell[][] rows = { new[] { cell } };
            if (route == "canvas-viewport") content.Canvas(canvas => canvas.Table(rows, 30, 30, 100, 36, table));
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
        var paragraph = new PdfTableCellParagraph(runs, fontSize: 12,
            lineSpacing: PdfLineSpacing.Exactly(18), tabStops: new[] { new PdfTabStop(36, PdfTabAlignment.Bar) });
        var cell = new PdfTableCell(runs, new[] { paragraph })
            .WithViewport(new PdfTableCellViewport(100, 36, 60, 36, offsetX: 24));
        var table = BarOnlyCellStyle(60, 36);
        byte[] bytes = PdfDocument.Create(builder => builder.Content(content => {
            PdfTableCell[][] rows = { new[] { cell } };
            if (route == "canvas") content.Canvas(canvas => canvas.Table(rows, 30, 30, 60, 36, table));
            else content.Table(rows, style: table);
        }), Options()).ToBytes();

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
        var mark = new PdfTextRun(string.Empty, fontSize: 12) { IsLineBoxStrut = true };
        var paragraph = new PdfTableCellParagraph(new[] { mark }, fontSize: 12,
            lineSpacing: PdfLineSpacing.Exactly(18), tabStops: new[] { new PdfTabStop(36, PdfTabAlignment.Bar) });
        var cell = new PdfTableCell(new[] { mark }, new[] { paragraph }, textRotation: rotation);
        byte[] bytes = PdfDocument.Create(builder => builder.Content(content => content.Table(
            new[] { new[] { cell } }, style: BarOnlyCellStyle(36, 100))), Options()).ToBytes();

        using var actual = PdfPigDocument.Open(bytes);
        PdfPigPage page = Assert.Single(actual.GetPages());
        Assert.Empty(page.Letters);
        var bar = Assert.Single(page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds.HasValue && Math.Abs(bounds.Value.Height) < .001D && bounds.Value.Width > 1D));
        Assert.Equal(left, bar!.Value.Left, 3);
        Assert.Equal(y, bar.Value.Bottom, 3);
        Assert.Equal(18D, bar.Value.Width, 3);
    }

    private static PdfTableStyle BarOnlyCellStyle(double width, double height) => new() {
        HeaderRowCount = 0, FontSize = 12, CellPaddingX = 0, CellPaddingY = 0,
        BorderWidth = 0, SpacingBefore = 0, SpacingAfter = 0, MaxWidth = width, PreserveWidth = true,
        ColumnWidthPoints = new System.Collections.Generic.List<double?> { width },
        FixedRowHeights = new System.Collections.Generic.List<double?> { height }
    };

    private static PdfOptions Options() => new() {
        PageWidth = 360, PageHeight = 240, MarginLeft = 30, MarginRight = 30,
        MarginTop = 30, MarginBottom = 30, DefaultFont = PdfStandardFont.Helvetica,
        DefaultFontSize = 12, CompressContentStreams = false
    };

    private static PdfParagraphStyle ParagraphStyle() => new() {
        FontSize = 12, LineSpacing = PdfLineSpacing.Exactly(18), SpacingBefore = 0, SpacingAfter = 0,
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
