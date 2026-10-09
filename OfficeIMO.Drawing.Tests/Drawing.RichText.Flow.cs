using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextFlowTests {
    private static double Measure(string? value, double size, string? family, OfficeFontStyle style) => (value?.Length ?? 0) * size / 2;
    private static string Text(IReadOnlyList<OfficeRichTextParagraph> paragraphs) => string.Join("\n",
        paragraphs.Select(paragraph => string.Concat(paragraph.Runs.Select(run => run.Text))));

    [Fact]
    public void ContinuationUsesCompleteWrappedLinesAndKeepsSourceWhitespaceAndStyles() {
        var paragraph = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("alpha beta ", 10, OfficeColor.Black),
            new OfficeRichTextRun("gamma delta", 10, OfficeColor.Red, bold: true)
        });
        int measured = 0;
        var flow = new OfficeRichTextFlow(new[] { paragraph }, characters => measured += characters);
        var first = flow.Take(50, 12, Measure, default);
        Assert.Equal("alpha beta ", Text(first)); Assert.Equal(11, flow.CharacterPosition); Assert.True(flow.HasRemaining);
        var second = flow.Take(55, 12, Measure, default);
        Assert.Equal("gamma delta", Text(second)); Assert.Equal(22, flow.CharacterPosition); Assert.False(flow.HasRemaining);
        Assert.True(Assert.Single(second[0].Runs).Bold); Assert.Equal(OfficeColor.Red, second[0].Runs[0].Color);
        Assert.Equal(33, measured);
    }

    [Fact]
    public void ContinuationDoesNotRepeatParagraphIndentLabelOrSpaceBefore() {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("one two three four", 10, OfficeColor.Black) },
            OfficeTextParagraphLabel.InBox(new OfficeRichTextRun("•", 10, OfficeColor.Black), 0, 5, minimumDistance: 3),
            margins: new OfficeTextPadding(0, 2, 0, 4), indent: new OfficeTextParagraphIndent(5, 0));
        var flow = new OfficeRichTextFlow(new[] { paragraph }, _ => { });
        var first = flow.Take(50, 14, Measure, default);
        Assert.NotNull(Assert.Single(first).Label); Assert.True(flow.HasRemaining);
        var second = flow.Take(100, 40, Measure, default);
        OfficeRichTextParagraph tail = Assert.Single(second);
        Assert.Null(tail.Label); Assert.Equal(0, tail.Indent.FirstLineOffset); Assert.Equal(0, tail.Margins.Top);
        Assert.Equal(4, tail.Margins.Bottom); Assert.False(flow.HasRemaining);
    }

    [Fact]
    public void SourceEndpointsSurviveTabsHardBreaksAndGraphemeSplitting() {
        const string source = "a\tb\r\nCafé\u0301🐈xxxxxxxx";
        var flow = new OfficeRichTextFlow(new[] { new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(source, 10, OfficeColor.Black) }) }, _ => { });
        string recovered = string.Empty;
        for (int region = 0; region < 20 && flow.HasRemaining; region++) {
            var slice = flow.Take(30, 12, Measure, default);
            string text = Text(slice);
            Assert.NotEmpty(text);
            Assert.False(char.IsLowSurrogate(text[0])); Assert.False(char.IsHighSurrogate(text[text.Length - 1]));
            recovered += text;
        }
        Assert.Equal(source, recovered); Assert.Equal(source.Length, flow.CharacterPosition); Assert.False(flow.HasRemaining);
    }

    [Fact]
    public void EmptyParagraphsConsumeLinesAndUnusableRegionsDoNotConsumeText() {
        var paragraphs = new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(string.Empty, 10, OfficeColor.Black) }),
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("last", 10, OfficeColor.Black) })
        };
        var flow = new OfficeRichTextFlow(paragraphs, _ => { });
        Assert.Empty(flow.Take(40, 1, Measure, default)); Assert.Equal(0, flow.CharacterPosition);
        Assert.Single(flow.Take(40, 12, Measure, default)); Assert.Equal(1, flow.CharacterPosition);
        Assert.Equal("last", Text(flow.Take(40, 12, Measure, default))); Assert.Equal(5, flow.CharacterPosition);
        Assert.False(flow.HasRemaining);
    }

    [Fact]
    public void FlowHonorsCallerCancellationAndMeasurementBudgetBeforeLayout() {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("remaining", 10, OfficeColor.Black) });
        var flow = new OfficeRichTextFlow(new[] { paragraph }, _ => throw new InvalidOperationException("work limit"));
        Assert.Throws<InvalidOperationException>(() => flow.Take(40, 20, Measure, default)); Assert.Equal(0, flow.CharacterPosition);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => flow.Take(40, 20, Measure, cancelled.Token));
    }

    [Fact]
    public void JustifiedParagraphContinuesAtTheRegionEdgeIncludingDrawingCopies() {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("one two three four", 10, OfficeColor.Black) }, OfficeTextAlignment.Justify);
        var flow = new OfficeRichTextFlow(new[] { paragraph }, _ => { });
        var drawing = new OfficeDrawing(50, 12).AddRichTextParagraphs(flow.Take(50, 12, Measure, default), 0, 0, 50, 12);
        OfficeDrawing tinted = drawing.Clone(); tinted.ApplyColorTint(OfficeColor.Red);
        foreach (OfficeDrawing copy in new[] { drawing, drawing.Clone(), tinted }) {
            var text = Assert.IsType<OfficeDrawingRichText>(Assert.Single(copy.Elements));
            OfficeRichTextBlockLayout layout = OfficeDrawingTextLayout.Create(text, 50, 12, Measure);
            Assert.Equal(50, Assert.Single(layout.Lines).Width, 6);
        }
    }

    [Fact]
    public void FlowKeepsOversizedGlyphForANextRegionThatCanContainIt() {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("W", 10, OfficeColor.Black) });
        var flow = new OfficeRichTextFlow(new[] { paragraph }, _ => { });
        Assert.Empty(flow.Take(2, 12, Measure, default)); Assert.Equal(0, flow.CharacterPosition); Assert.True(flow.HasRemaining);
        Assert.Equal("W", Text(flow.Take(10, 12, Measure, default))); Assert.False(flow.HasRemaining);
    }

    [Theory]
    [InlineData("\n")]
    [InlineData("\r")]
    [InlineData("\r\n")]
    public void TerminalHardBreakKeepsItsPendingEmptyLineAcrossRegions(string breakText) {
        var paragraphs = new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A" + breakText, 10, OfficeColor.Black) }),
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("B", 10, OfficeColor.Black) })
        };
        var flow = new OfficeRichTextFlow(paragraphs, _ => { });
        Assert.Equal("A" + breakText, Text(flow.Take(100, 12, Measure, default)));
        Assert.Equal(1 + breakText.Length, flow.CharacterPosition); Assert.True(flow.HasRemaining);
        Assert.Equal(string.Empty, Text(flow.Take(100, 12, Measure, default)));
        Assert.Equal(2 + breakText.Length, flow.CharacterPosition); Assert.True(flow.HasRemaining);
        Assert.Equal("B", Text(flow.Take(100, 12, Measure, default))); Assert.False(flow.HasRemaining);
    }

    [Theory]
    [InlineData(10, 20, 12)]
    [InlineData(20, 10, 24)]
    public void PendingEmptyLineRetainsTheOriginalMixedRunHeight(double firstSize, double breakSize, double firstHeight) {
        var paragraphs = new[] {
            new OfficeRichTextParagraph(new[] {
                new OfficeRichTextRun("A", firstSize, OfficeColor.Black),
                new OfficeRichTextRun("\n", breakSize, OfficeColor.Black)
            }),
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("B", 10, OfficeColor.Black) })
        };
        var flow = new OfficeRichTextFlow(paragraphs, _ => { });
        Assert.Equal("A\n", Text(flow.Take(100, firstHeight, Measure, default)));
        Assert.Equal(2, flow.CharacterPosition);
        Assert.Empty(flow.Take(100, 12, Measure, default));
        Assert.Equal(2, flow.CharacterPosition); Assert.True(flow.HasRemaining);
        IReadOnlyList<OfficeRichTextParagraph> blank = flow.Take(100, 24, Measure, default);
        Assert.Equal(string.Empty, Text(blank)); Assert.Single(blank);
        OfficeRichTextBlockLayout layout = OfficeDrawingTextLayout.CreateParagraphs(blank, 100, 24, Measure);
        Assert.Equal(24, Assert.Single(layout.Lines).LineHeight);
        Assert.Equal(3, flow.CharacterPosition);
        Assert.Equal("B", Text(flow.Take(100, 12, Measure, default))); Assert.False(flow.HasRemaining);
    }
}
