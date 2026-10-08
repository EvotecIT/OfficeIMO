using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextLabelFittingTests {
    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void ParagraphFittingScalesLabelAndBodyTogetherAndKeepsTheirPositions(double scale) {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Body", 10, OfficeColor.Black) },
            OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun("*", 30, OfficeColor.Red), 4, textPosition: 24),
            margins: new OfficeTextPadding(24, 0, 0, 0));
        var layout = OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 100 * scale, 15 * scale,
            Measure, scale: scale, shrinkToFit: true, minimumFontSize: 1 * scale);
        var line = Assert.Single(layout.Lines);
        var label = line.Segments.First(s => s.Text == "*");
        var body = line.Segments.First(s => s.Text == "Body");
        Assert.False(layout.Clipped);
        Assert.InRange(label.FontSize, 12.4 * scale, 12.6 * scale);
        Assert.InRange(label.FontSize / body.FontSize, 2.999, 3.001);
        Assert.Equal(4 * scale, line.OffsetX);
        Assert.Equal(24 * scale, line.OffsetX + line.Segments.TakeWhile(s => s.Text != "Body").Sum(s => s.Width));
    }

    [Fact]
    public void LabelOnlyParagraphParticipatesInFittingAndMinimumSize() {
        var paragraph = new OfficeRichTextParagraph(Array.Empty<OfficeRichTextRun>(),
            OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun("W", 30, OfficeColor.Black), 0));
        var fitted = OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 100, 15, Measure,
            shrinkToFit: true, minimumFontSize: 6);
        Assert.False(fitted.Clipped);
        Assert.InRange(Assert.Single(fitted.Lines).Segments.Single(s => s.Text == "W").FontSize, 12.4, 12.6);
        var bounded = OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph }, 100, 2, Measure,
            shrinkToFit: true, minimumFontSize: 6);
        Assert.True(bounded.Clipped);
    }

    [Fact]
    public void CancellationFromLabelMeasurementStopsTheRemainingLabelAndBody() {
        using var cancellation = new CancellationTokenSource();
        int calls = 0;
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Body", 10, OfficeColor.Black) },
            OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun("label with many separate words", 10, OfficeColor.Black), 0));
        Assert.Throws<OperationCanceledException>(() => OfficeDrawingTextLayout.CreateParagraphs(new[] { paragraph },
            100, 100, (text, size, family, style) => { calls++; cancellation.Cancel(); return 5; },
            cancellationToken: cancellation.Token));
        Assert.InRange(calls, 1, 2);
    }

    private static double Measure(string? text, double size, string? family, OfficeFontStyle style) => (text?.Length ?? 0) * size / 2;
}
