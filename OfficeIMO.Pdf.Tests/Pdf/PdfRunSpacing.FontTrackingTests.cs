using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRunSpacingFontTrackingTests {
    [Theory]
    [InlineData(200D, 1D, true, 500, -100)]
    [InlineData(200D, 1D, false, 500, -100)]
    [InlineData(50D, -0.5D, true, -500, 100)]
    [InlineData(50D, -0.5D, false, -500, 100)]
    [InlineData(100D, 0D, true, 500, -100)]
    [InlineData(100D, 0D, false, 500, -100)]
    public void TableShrinkingFitsThePaintedWidthWithSizeDependentFontTracking(double scale, double spacing, bool explicitSize, short smallTracking, short largeTracking) {
        var options = TrackingOptions(smallTracking, largeTracking);
        PdfTextRun Run(double? size) => new PdfTextRun("AABBAABBAABBAABB", fontSize: size)
            .WithHorizontalTextScaling(scale).WithCharacterSpacing(spacing);
        using var table = PdfPigDocument.Open(PdfDocument.Create(options).Table(new[] {
            new[] { PdfTableCell.RichTextCell(new[] { Run(explicitSize ? 30D : null) }) }
        }, style: new PdfTableStyle {
            FontSize = 30D, MinimumShrinkFontSize = 1D, ShrinkTextToFit = true,
            ColumnWidthPoints = new() { 100D }, CellPaddingX = 0D, CellPaddingY = 0D, HeaderRowCount = 0
        }).ToBytes());
        var letters = table.GetPage(1).Letters;
        Assert.Equal("AABBAABBAABBAABB", string.Concat(letters.Select(letter => letter.Value)));
        Assert.All(letters, letter => Assert.Equal(letters[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
        double chosenSize = letters[0].FontSize;
        using var probe = PdfPigDocument.Open(PdfDocument.Create(options).Paragraph(p =>
            p.Runs(new[] { Run(chosenSize), PdfTextRun.Normal("X", fontSize: chosenSize) })).ToBytes());
        var painted = probe.GetPage(1).Letters;
        double width = painted.Single(letter => letter.Value == "X").StartBaseLine.X - painted[0].StartBaseLine.X;
        Assert.InRange(width, 99.9D, 100.02D);
    }

    [Theory]
    [InlineData(false, 50D, 1D)]
    [InlineData(false, 200D, -0.5D)]
    [InlineData(true, 50D, 1D)]
    [InlineData(true, 200D, -0.5D)]
    public void WordUnderlinesFollowTrackedGlyphsAndWordStarts(bool positioned, double scale, double spacing) {
        var options = TrackingOptions(-100, 0);
        var run = new PdfTextRun("AA BB", fontSize: 12D, underlineStyle: OfficeIMO.Drawing.OfficeTextDecorationStyle.Words)
            .WithHorizontalTextScaling(scale).WithCharacterSpacing(spacing);
        var runs = new[] { run, PdfTextRun.Normal("X", fontSize: 12D) };
        var document = PdfDocument.Create(options);
        if (positioned) document.Canvas(canvas => canvas.PositionedText(runs, PdfCanvasTextStructureRole.Paragraph,
            40D, 40D, 400D, 40D, fontSize: 12D, fontMetricScale: 0.75D));
        else document.Paragraph(p => p.Runs(runs));
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = pdf.GetPage(1);
        var underlines = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds.HasValue).Select(bounds => bounds!.Value).OrderBy(bounds => bounds.Left).ToArray();
        Assert.Equal(2, underlines.Length);
        double firstStart = page.Letters.First(letter => letter.Value == "A").StartBaseLine.X;
        double firstEnd = page.Letters.Single(letter => letter.Value == " ").StartBaseLine.X;
        double secondStart = page.Letters.First(letter => letter.Value == "B").StartBaseLine.X;
        double secondEnd = page.Letters.Single(letter => letter.Value == "X").StartBaseLine.X;
        Assert.InRange(Math.Abs(underlines[0].Left - firstStart), 0D, 0.02D);
        Assert.InRange(Math.Abs(underlines[0].Right - firstEnd), 0D, 0.02D);
        Assert.InRange(Math.Abs(underlines[1].Left - secondStart), 0D, 0.02D);
        Assert.InRange(Math.Abs(underlines[1].Right - secondEnd), 0D, 0.02D);
    }

    private static PdfOptions TrackingOptions(short smallTracking, short largeTracking) {
        string? path = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] font = ManagedTextShapingTestAssets.AddTrackingTable(File.ReadAllBytes(path!),
            ManagedTextShapingTestAssets.CreateTrackingTable(new[] { 10D, 20D }, new[] { 0D },
                new[] { new[] { smallTracking, largeTracking } }));
        var options = new PdfOptions { DefaultFontSize = 12D };
        options.EmbedStandardFont(PdfStandardFont.Helvetica, font, "TrackedSpacing");
        return options;
    }

    [Theory]
    [InlineData(false, 50D, 1D)]
    [InlineData(false, 200D, -0.5D)]
    [InlineData(true, 50D, 1D)]
    [InlineData(true, 200D, -0.5D)]
    public void IntrinsicFontTrackingCombinesWithRunSpacingAndConvertedFontMetrics(bool positioned, double scale, double spacing) {
        string? path = PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] font = ManagedTextShapingTestAssets.AddTrackingTable(File.ReadAllBytes(path!),
            ManagedTextShapingTestAssets.CreateTrackingTable(new[] { 10D, 20D }, new[] { 0D },
                new[] { new short[] { -100, 0 } }));
        byte[] Render(bool styled) {
            var options = new PdfOptions { DefaultFontSize = 12D };
            options.EmbedStandardFont(PdfStandardFont.Helvetica, font, "TrackedSpacing");
            var run = new PdfTextRun("AABB", fontSize: 12D);
            if (styled) run = run.WithHorizontalTextScaling(scale).WithCharacterSpacing(spacing);
            var runs = new[] { run, PdfTextRun.Normal("X", fontSize: 12D) };
            PdfDocument document = PdfDocument.Create(options);
            if (positioned) document.Canvas(canvas => canvas.PositionedText(runs, PdfCanvasTextStructureRole.Paragraph,
                40D, 40D, 400D, 40D, fontSize: 12D, fontMetricScale: 0.75D));
            else document.Paragraph(paragraph => paragraph.Runs(runs));
            return document.ToBytes();
        }
        using var natural = PdfPigDocument.Open(Render(false));
        using var styled = PdfPigDocument.Open(Render(true));
        var before = natural.GetPage(1).Letters;
        var after = styled.GetPage(1).Letters;
        Assert.Equal("AABBX", string.Concat(after.Select(letter => letter.Value)));
        double naturalAdvance = before.Single(letter => letter.Value == "X").StartBaseLine.X - before[0].StartBaseLine.X;
        double styledAdvance = after.Single(letter => letter.Value == "X").StartBaseLine.X - after[0].StartBaseLine.X;
        Assert.InRange(Math.Abs(styledAdvance - (naturalAdvance * scale / 100D + 4D * spacing)), 0D, 0.02D);
        for (int index = 0; index < 4; index++) {
            double expected = (before[index].StartBaseLine.X - before[0].StartBaseLine.X) * scale / 100D + index * spacing;
            Assert.InRange(Math.Abs(after[index].StartBaseLine.X - after[0].StartBaseLine.X - expected), 0D, 0.02D);
        }
    }
}
