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
