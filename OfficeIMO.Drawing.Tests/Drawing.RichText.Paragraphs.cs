using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextParagraphTests {
    private static double Measure(string? value, double size, string? family, OfficeFontStyle style) => (value?.Length ?? 0) * size / 2;
    private static OfficeRichTextParagraph Paragraph(string value, double size = 10, OfficeTextAlignment alignment = OfficeTextAlignment.Left,
        double? height = null, OfficeTextPadding? margins = null, double? factor = null) => new OfficeRichTextParagraph(
            new[] { new OfficeRichTextRun(value, size, OfficeColor.Black) }, alignment, height, margins, lineHeightFactor: factor);

    [Fact]
    public void AlignmentInsetsAndSpacingStayIndependentAcrossParagraphs() {
        var drawing = new OfficeDrawing(100, 100).AddRichTextParagraphs(new[] {
            Paragraph("first", alignment: OfficeTextAlignment.Center, height: 14, margins: new OfficeTextPadding(10, 2, 20, 3)),
            Paragraph("last", alignment: OfficeTextAlignment.Right, height: 20, margins: new OfficeTextPadding(5, 4, 15, 6))
        }, 0, 0, 100, 100);
        var text = Assert.IsType<OfficeDrawingRichText>(Assert.Single(drawing.Elements));
        var layout = OfficeDrawingTextLayout.Create(text, 100, 100, Measure);
        var visible = layout.Lines.Where(l => l.Segments.Count > 0).ToArray();
        Assert.Equal(2, visible.Length); Assert.Equal(32.5, visible[0].OffsetX); Assert.Equal(65, visible[1].OffsetX);
        Assert.Equal(14, visible[0].LineHeight); Assert.Equal(20, visible[1].LineHeight);
        Assert.Equal(49, layout.Height); Assert.False(layout.Clipped);
        Assert.Equal("first\nlast", text.PlainText);
    }

    [Fact]
    public void AbsoluteLeadingAndRelativeLeadingHandleMixedWrappedFontSizes() {
        var runs = new[] { new OfficeRichTextRun("BIG\n", 20, OfficeColor.Black), new OfficeRichTextRun("small", 10, OfficeColor.Black) };
        OfficeRichTextBlockLayout Layout(double? height, double? factor) {
            var drawing = new OfficeDrawing(100, 100).AddRichTextParagraphs(new[] { new OfficeRichTextParagraph(runs, lineHeight: height, lineHeightFactor: factor) }, 0, 0, 100, 100);
            return OfficeDrawingTextLayout.Create((OfficeDrawingRichText)drawing.Elements[0], 100, 100, Measure);
        }
        Assert.Equal(new[] { 15D, 15D }, Layout(15, null).Lines.Select(l => l.LineHeight));
        Assert.Equal(new[] { 30D, 15D }, Layout(null, 1.5).Lines.Select(l => l.LineHeight));
    }

    [Fact]
    public void HardLineBreakContinuesIndentationWhileANewParagraphRestartsIt() {
        var p = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("first\nnext", 10, OfficeColor.Black) }, indent: OfficeTextParagraphIndent.FirstLine(15));
        var drawing = new OfficeDrawing(100, 80).AddRichTextParagraphs(new[] { p, p }, 0, 0, 100, 80);
        var layout = OfficeDrawingTextLayout.Create((OfficeDrawingRichText)drawing.Elements[0], 100, 80, Measure);
        Assert.Equal(new[] { 15D, 0D, 15D, 0D }, layout.Lines.Select(l => l.OffsetX));
    }

    [Fact]
    public void JustifiedAdvancesFillWrappedLinesButLeaveTheLastLineNatural() {
        var drawing = new OfficeDrawing(55, 100).AddRichTextParagraphs(new[] { Paragraph("one two three four", alignment: OfficeTextAlignment.Justify) }, 0, 0, 55, 100);
        var layout = OfficeDrawingTextLayout.Create((OfficeDrawingRichText)drawing.Elements[0], 55, 100, Measure);
        Assert.Equal(2, layout.Lines.Count); Assert.Equal(55, layout.Lines[0].Width); Assert.Equal(50, layout.Lines[1].Width);
        var svg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing));
        Assert.Contains(svg.Descendants(), e => e.Name.LocalName == "text" && e.Value == "one");
    }

    [Fact]
    public void WidthOverflowDoesNotDiscardLaterParagraphsAndHeightOverflowIsBounded() {
        var drawing = new OfficeDrawing(40, 50).AddRichTextParagraphs(new[] { Paragraph("wide wide wide"), Paragraph("last") }, 0, 0, 40, 50, wrapText: false);
        var text = (OfficeDrawingRichText)drawing.Elements[0];
        var wide = OfficeDrawingTextLayout.Create(text, 40, 50, Measure);
        Assert.True(wide.Clipped); Assert.Equal(2, wide.Lines.Count);
        Assert.Equal("last", string.Concat(wide.Lines[1].Segments.Select(s => s.Text)));
        var shortFrame = OfficeDrawingTextLayout.Create(text, 40, 12, Measure);
        Assert.True(shortFrame.Clipped); Assert.Single(shortFrame.Lines); Assert.Equal(12, shortFrame.Height);
    }

    [Theory]
    [InlineData("longcaption", OfficeTextAlignment.Center, -8.5)]
    [InlineData("longcaption", OfficeTextAlignment.Right, -21)]
    [InlineData("fit", OfficeTextAlignment.Center, 11.5)]
    [InlineData("fit", OfficeTextAlignment.Right, 19)]
    public void UnwrappedParagraphsRetainTheirAlignmentAnchorWhenWiderThanTheFrame(string value, OfficeTextAlignment alignment, double expectedLeft) {
        var drawing = new OfficeDrawing(40, 50).AddRichTextParagraphs(new[] {
            Paragraph(value, alignment: alignment, margins: new OfficeTextPadding(4, 0, 6, 0))
        }, 0, 0, 40, 50, wrapText: false);
        var text = Assert.IsType<OfficeDrawingRichText>(Assert.Single(drawing.Elements));
        var layout = OfficeDrawingTextLayout.Create(text, 40, 50, Measure);
        var line = Assert.Single(layout.Lines);
        Assert.Equal(expectedLeft, line.OffsetX);
        Assert.Equal(value, string.Concat(line.Segments.Select(segment => segment.Text)));
        Assert.Equal(value.Length > 6, layout.Clipped);
    }

    [Theory]
    [InlineData(100, 0)]
    [InlineData(60, 50)]
    public void ParagraphWithoutHorizontalRoomDoesNotHideFollowingContent(double left, double right) {
        var drawing = new OfficeDrawing(100, 100).AddRichTextParagraphs(new[] {
            Paragraph("Outside", margins: new OfficeTextPadding(left, 0, right, 0)), Paragraph("Visible")
        }, 0, 0, 100, 100);
        var layout = OfficeDrawingTextLayout.Create((OfficeDrawingRichText)drawing.Elements[0], 100, 100, Measure);
        Assert.True(layout.Clipped);
        Assert.Equal("Visible", string.Concat(layout.Lines.SelectMany(l => l.Segments).Select(s => s.Text)));
        Assert.Contains("Visible", OfficeDrawingSvgExporter.ToSvg(drawing));
    }

    [Fact]
    public void SnapshotsCloneNestingAndTintRetainParagraphFormatAndLinkTargets() {
        var runs = new List<OfficeRichTextRun> { new OfficeRichTextRun("linked", 10, OfficeColor.Red) { LinkUri = "https://example.invalid/help" } };
        var paragraphs = new List<OfficeRichTextParagraph> { new OfficeRichTextParagraph(runs, OfficeTextAlignment.Right,
            margins: new OfficeTextPadding(2, 3, 4, 5), indent: OfficeTextParagraphIndent.Hanging(6), lineHeightFactor: 1.5) };
        var drawing = new OfficeDrawing(80, 50).AddRichTextParagraphs(paragraphs, 0, 0, 80, 50);
        runs.Clear(); paragraphs.Clear();
        var copy = new OfficeDrawing(100, 70).AddDrawing(drawing.Clone(), 10, 10); copy.ApplyColorTint(OfficeColor.Blue);
        var text = Assert.IsType<OfficeDrawingRichText>(Assert.Single(copy.Elements));
        var p = Assert.Single(text.Paragraphs); var run = Assert.Single(p.Runs);
        Assert.Equal(OfficeTextAlignment.Right, p.Alignment); Assert.Equal(1.5, p.LineHeightFactor);
        Assert.Equal(3, p.Margins.Top); Assert.Equal(6, p.Indent.ContinuationLineOffset);
        Assert.Equal(OfficeColor.Blue, run.Color); Assert.Equal("https://example.invalid/help", run.LinkUri);
        Assert.Equal("linked", text.PlainText); Assert.Equal(10, text.X); Assert.Equal(10, text.Y);
    }

    [Fact]
    public void FontsAddedAfterParagraphCreationAffectRenderTimeAlignment() {
        var drawing = new OfficeDrawing(100, 50).AddRichTextParagraphs(new[] {
            new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("WWWWiiii", 12, OfficeColor.Black, fontFamily: "Proof Sans") }, OfficeTextAlignment.Right)
        }, 0, 0, 100, 50);
        double Position() => double.Parse(XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing)).Descendants()
            .First(e => e.Name.LocalName == "text").Attribute("x")!.Value, CultureInfo.InvariantCulture);
        double fallback = Position();
        drawing.Fonts.Add("Proof Sans", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf")));
        double embedded = Position(); Assert.NotEqual(fallback, embedded); Assert.InRange(embedded, 0, 100);
        Assert.Equal("WWWWiiii", ((OfficeDrawingRichText)drawing.Elements[0]).PlainText);
    }

    [Fact]
    public void InvalidOrOversizedParagraphInputLeavesTheDrawingUnchanged() {
        var drawing = new OfficeDrawing(100, 100);
        Assert.Throws<ArgumentException>(() => new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(new string('x', 100001), 10, OfficeColor.Black) }));
        Assert.Throws<ArgumentException>(() => new OfficeRichTextParagraph(Array.Empty<OfficeRichTextRun>(), lineHeight: 12, lineHeightFactor: 1.2));
        var half = Paragraph(new string('x', 50000));
        Assert.Throws<ArgumentException>(() => drawing.AddRichTextParagraphs(new[] { half, half }, 0, 0, 100, 100));
        Assert.Empty(drawing.Elements);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SvgRichTextLinksKeepTargetsAndRejectUnsafeSchemes(bool paragraphProfile) {
        var runs = new[] {
            new OfficeRichTextRun("first ", 10, OfficeColor.Black) { LinkUri = "https://example.invalid/?a=1&b=2" },
            new OfficeRichTextRun("second", 10, OfficeColor.Black) { LinkUri = "#anchor" },
            new OfficeRichTextRun(" unsafe", 10, OfficeColor.Black) { LinkUri = "javascript:alert(1)" }
        };
        var drawing = new OfficeDrawing(100, 50);
        if (paragraphProfile) drawing.AddRichTextParagraphs(new[] { new OfficeRichTextParagraph(runs) }, 0, 0, 100, 50);
        else drawing.AddRichText(runs, 0, 0, 100, 50);
        var xml = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing));
        var links = xml.Descendants().Where(e => e.Name.LocalName == "a").ToArray();
        Assert.Contains(links, e => (string?)e.Attribute("href") == "https://example.invalid/?a=1&b=2" && e.Value.Contains("first"));
        Assert.Contains(links, e => (string?)e.Attribute("href") == "#anchor" && e.Value.Contains("second"));
        Assert.DoesNotContain(links, e => ((string?)e.Attribute("href"))?.StartsWith("javascript:", StringComparison.Ordinal) == true);
        Assert.Contains("unsafe", xml.Root!.Value);
    }
}
