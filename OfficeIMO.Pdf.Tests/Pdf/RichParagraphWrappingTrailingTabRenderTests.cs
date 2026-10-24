using System;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class RichParagraphWrappingTests {
    [Theory]
    [InlineData(PdfTextWhitespaceMode.Collapse, false, false)]
    [InlineData(PdfTextWhitespaceMode.Collapse, false, true)]
    [InlineData(PdfTextWhitespaceMode.Collapse, true, true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, false, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, true, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, true, true)]
    public void ParagraphTabs_TrailingUnderlineLeaderRendersWithoutFollowingText(
        PdfTextWhitespaceMode whitespace, bool hardBreak, bool trailingSpace) {
        const double margin = 36D, firstStop = 72D, leaderStop = 216D;
        var style = new PdfParagraphStyle { TextWhitespaceMode = whitespace, SpacingAfter = 0D };
        style.AddTabStop(firstStop);
        style.AddTabStop(leaderStop, leader: PdfTabLeaderStyle.Underscores);
        byte[] bytes = PdfDocument.Create(TrailingTabOptions())
            .Paragraph(p => {
                p.Tab().Tab();
                if (trailingSpace) p.Text(" ");
                if (hardBreak) p.LineBreak().Text("AFTER");
            }, style: style)
            .Paragraph(p => p.Text("NEXT"), style: new PdfParagraphStyle { SpacingAfter = 0D })
            .ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var leaders = page.Letters.Where(letter => letter.Value == "_")
            .OrderBy(letter => letter.StartBaseLine.X).ToArray();
        Assert.True(leaders.Length >= 3, "A trailing tab must paint its authored underline leader without a following text run.");
        Assert.Equal(margin + firstStop, leaders[0].StartBaseLine.X, 2);
        double glyphAdvance = leaders[0].EndBaseLine.X - leaders[0].StartBaseLine.X;
        Assert.True(glyphAdvance > 0D);
        Assert.InRange(leaders[leaders.Length - 1].EndBaseLine.X,
            margin + leaderStop - glyphAdvance - .02D, margin + leaderStop + .02D);
        Assert.All(leaders, letter => Assert.Equal(leaders[0].StartBaseLine.Y, letter.StartBaseLine.Y, 2));
        var words = page.GetWords().ToArray();
        var next = Assert.Single(words, word => word.Text == "NEXT");
        Assert.Equal(margin, next.Letters[0].StartBaseLine.X, 2);
        Assert.True(next.Letters[0].StartBaseLine.Y < leaders[0].StartBaseLine.Y);
        if (hardBreak) {
            var after = Assert.Single(words, word => word.Text == "AFTER");
            Assert.Equal(margin, after.Letters[0].StartBaseLine.X, 2);
            Assert.True(after.Letters[0].StartBaseLine.Y < leaders[0].StartBaseLine.Y);
            Assert.True(after.Letters[0].StartBaseLine.Y > next.Letters[0].StartBaseLine.Y);
        }
    }

    [Fact]
    public void ParagraphTabs_PlainTrailingTabsDoNotPaintLeadersOrIndentFollowingContent() {
        byte[] bytes = PdfDocument.Create(TrailingTabOptions())
            .Paragraph(p => p.Text("BEFORE").Tab().Text(" "), style: new PdfParagraphStyle { SpacingAfter = 0D })
            .Paragraph(p => p.Tab().Text(" ").LineBreak().Text("AFTER").Tab().Text(" "),
                style: new PdfParagraphStyle { SpacingAfter = 0D })
            .Paragraph(p => p.Text("NEXT"), style: new PdfParagraphStyle { SpacingAfter = 0D })
            .ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var visible = page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
        Assert.Equal("BEFOREAFTERNEXT", string.Concat(visible.Select(letter => letter.Value)));
        var words = page.GetWords().ToArray();
        var before = Assert.Single(words, word => word.Text == "BEFORE");
        var after = Assert.Single(words, word => word.Text == "AFTER");
        var next = Assert.Single(words, word => word.Text == "NEXT");
        foreach (var word in new[] { before, after, next })
            Assert.Equal(36D, word.Letters[0].StartBaseLine.X, 2);
        Assert.True(before.Letters[0].StartBaseLine.Y > after.Letters[0].StartBaseLine.Y);
        Assert.True(after.Letters[0].StartBaseLine.Y > next.Letters[0].StartBaseLine.Y);
    }

    [Theory]
    [InlineData(PdfTextWhitespaceMode.Collapse, false)]
    [InlineData(PdfTextWhitespaceMode.Collapse, true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, true)]
    public void ParagraphTabs_SpaceBeforeNextTabRetainsBothLeaders(
        PdfTextWhitespaceMode whitespace, bool followingText) {
        const double margin = 36D, firstStop = 72D, secondStop = 216D;
        var style = new PdfParagraphStyle { TextWhitespaceMode = whitespace, SpacingAfter = 0D };
        style.AddTabStop(firstStop, leader: PdfTabLeaderStyle.Dots);
        style.AddTabStop(secondStop, leader: PdfTabLeaderStyle.Underscores);
        PdfOptions options = TrailingTabOptions();
        options.DefaultFont = PdfStandardFont.Courier;
        byte[] bytes = PdfDocument.Create(options).Paragraph(p => {
            p.Text("A").Tab().Text(" ").Tab();
            if (followingText) p.Text("B");
        }, style: style).ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var a = Assert.Single(page.Letters, letter => letter.Value == "A");
        var dots = page.Letters.Where(letter => letter.Value == ".").ToArray();
        var underscores = page.Letters.Where(letter => letter.Value == "_")
            .OrderBy(letter => letter.StartBaseLine.X).ToArray();
        Assert.NotEmpty(dots);
        Assert.NotEmpty(underscores);
        Assert.Equal(margin, a.StartBaseLine.X, 2);
        double monospaceAdvance = a.EndBaseLine.X - a.StartBaseLine.X;
        Assert.True(monospaceAdvance > 0D);
        double expectedStart = margin + firstStop +
            (whitespace == PdfTextWhitespaceMode.Collapse ? 0D : monospaceAdvance);
        Assert.Equal(expectedStart, underscores[0].StartBaseLine.X, 2);
        Assert.All(dots.Concat(underscores), letter => Assert.Equal(a.StartBaseLine.Y, letter.StartBaseLine.Y, 2));
        Assert.Equal(followingText ? "AB" : "A", string.Concat(page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value) && letter.Value != "." && letter.Value != "_")
            .Select(letter => letter.Value)));
        if (followingText) {
            var b = Assert.Single(page.Letters, letter => letter.Value == "B");
            Assert.Equal(margin + secondStop, b.StartBaseLine.X, 2);
            Assert.Equal(a.StartBaseLine.Y, b.StartBaseLine.Y, 2);
        }
    }

    [Theory]
    [InlineData(PdfTextWhitespaceMode.Collapse, false)]
    [InlineData(PdfTextWhitespaceMode.Collapse, true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, true)]
    public void ParagraphTabs_NoLeaderAdvancesTextAndInlineContentWithoutPaintingSpace(
        PdfTextWhitespaceMode whitespace, bool inline) {
        var style = new PdfParagraphStyle { TextWhitespaceMode = whitespace, SpacingAfter = 0D };
        style.AddTabStop(180D);
        PdfTextRun content = inline
            ? PdfTextRun.Inline(new PdfInlineBox(10D, 10D, background: PdfColor.Black))
            : PdfTextRun.Normal("VISIBLE", fontSize: 20D);
        byte[] bytes = PdfDocument.Create(TrailingTabInkOptions())
            .Paragraph(p => p.Runs(new[] { PdfTextRun.Normal("\t", fontSize: 24D), content }), style: style)
            .Paragraph(p => p.Text("NEXT"), style: new PdfParagraphStyle { SpacingAfter = 0D })
            .ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        // SPACE has an outline in this font. Retain every extracted glyph so
        // a manufactured separator cannot disappear behind a whitespace filter.
        Assert.Equal(inline ? "NEXT" : "VISIBLENEXT", string.Concat(page.Letters.Select(letter => letter.Value)));
        if (inline) {
            var bounds = Assert.Single(page.Paths.Where(path => path.IsFilled)
                .Select(path => path.GetBoundingRectangle()), rectangle => rectangle.HasValue);
            Assert.Equal(216D, bounds!.Value.Left, 3);
            Assert.Equal(10D, bounds.Value.Width, 3);
            Assert.Equal(10D, bounds.Value.Height, 3);
        } else {
            var visible = Assert.Single(page.Letters, letter => letter.Value == "V");
            Assert.Equal(216D, visible.StartBaseLine.X, 3);
            Assert.Equal(20D, visible.FontSize, 3);
        }
        var next = Assert.Single(page.Letters, letter => letter.Value == "N");
        Assert.Equal(36D, next.StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData(PdfTextWhitespaceMode.Collapse, false)]
    [InlineData(PdfTextWhitespaceMode.Collapse, true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, true)]
    public void ParagraphTabs_TerminalLeaderUsesItsPaintedFontBox(
        PdfTextWhitespaceMode whitespace, bool trailingSpace) {
        var style = new PdfParagraphStyle { TextWhitespaceMode = whitespace, SpacingAfter = 0D };
        style.AddTabStop(180D, leader: PdfTabLeaderStyle.Underscores);
        byte[] leaderBytes = PdfDocument.Create(TrailingTabInkOptions())
            .Paragraph(p => {
                // The effective paragraph stop overrides the run's dotted leader.
                p.Runs(new[] { new PdfTextRun("\t", fontSize: 24D, tabLeader: PdfTabLeaderStyle.Dots) });
                if (trailingSpace) p.Text(" ");
            }, style: style)
            .Paragraph(p => p.Text("NEXT"), style: new PdfParagraphStyle { SpacingAfter = 0D })
            .ToBytes();
        byte[] referenceBytes = PdfDocument.Create(TrailingTabInkOptions())
            .Paragraph(p => p.Runs(new[] { PdfTextRun.Normal("___", fontSize: 24D) }),
                style: new PdfParagraphStyle { SpacingAfter = 0D })
            .Paragraph(p => p.Text("NEXT"), style: new PdfParagraphStyle { SpacingAfter = 0D })
            .ToBytes();

        using var leaderPdf = UglyToad.PdfPig.PdfDocument.Open(leaderBytes);
        using var referencePdf = UglyToad.PdfPig.PdfDocument.Open(referenceBytes);
        var leaderPage = Assert.Single(leaderPdf.GetPages());
        var referencePage = Assert.Single(referencePdf.GetPages());
        var leaders = leaderPage.Letters.Where(letter => letter.Value == "_").ToArray();
        Assert.True(leaders.Length >= 3);
        Assert.Equal(new string('_', leaders.Length) + "NEXT", string.Concat(leaderPage.Letters.Select(letter => letter.Value)));
        var referenceLeader = referencePage.Letters.First(letter => letter.Value == "_");
        Assert.Equal(36D, leaders[0].StartBaseLine.X, 3);
        Assert.All(leaders, letter => {
            Assert.Equal(24D, letter.FontSize, 3);
            Assert.Equal(referenceLeader.StartBaseLine.Y, letter.StartBaseLine.Y, 3);
        });
        var next = Assert.Single(leaderPage.Letters, letter => letter.Value == "N");
        var referenceNext = Assert.Single(referencePage.Letters, letter => letter.Value == "N");
        Assert.Equal(referenceNext.StartBaseLine.Y, next.StartBaseLine.Y, 3);
        Assert.Equal(36D, next.StartBaseLine.X, 3);
    }

    [Fact]
    public void ParagraphTabs_SubThresholdLeaderKeepsThePlainTabFontBox() {
        byte[] Render(PdfTabLeaderStyle leader) {
            var style = new PdfParagraphStyle { TextWhitespaceMode = PdfTextWhitespaceMode.Preserve, SpacingAfter = 0D };
            // At24pt this12pt gap is smaller than the existing three-glyph
            // leader threshold. The selected leader therefore paints no glyphs.
            style.AddTabStop(12D, leader: leader);
            return PdfDocument.Create(TrailingTabInkOptions())
                .Paragraph(p => p.Runs(new[] { new PdfTextRun("\t", fontSize: 24D, tabLeader: leader) }), style: style)
                .Paragraph(p => p.Text("NEXT"), style: new PdfParagraphStyle { SpacingAfter = 0D })
                .ToBytes();
        }
        using var leaderPdf = UglyToad.PdfPig.PdfDocument.Open(Render(PdfTabLeaderStyle.Underscores));
        using var plainPdf = UglyToad.PdfPig.PdfDocument.Open(Render(PdfTabLeaderStyle.None));
        var leaderPage = Assert.Single(leaderPdf.GetPages());
        var plainPage = Assert.Single(plainPdf.GetPages());
        Assert.Equal("NEXT", string.Concat(leaderPage.Letters.Select(letter => letter.Value)));
        Assert.Equal("NEXT", string.Concat(plainPage.Letters.Select(letter => letter.Value)));
        for (int index = 0; index < plainPage.Letters.Count; index++) {
            Assert.Equal(plainPage.Letters[index].StartBaseLine.X, leaderPage.Letters[index].StartBaseLine.X, 3);
            Assert.Equal(plainPage.Letters[index].StartBaseLine.Y, leaderPage.Letters[index].StartBaseLine.Y, 3);
        }
    }
    [Fact]
    public void ParagraphTabs_BufferedLeaderRefitsItsSourceFontBeforeSelectingTheNextStop() {
        var requests = new System.Collections.Generic.List<(double Height, double MinimumWidth)>();
        Func<int, double, double, double, (double Width, double Origin, double Gap)> frame =
            (line, completed, height, minimumWidth) => {
                requests.Add((height, minimumWidth));
                return (height > 20D ? 60D : 288D, 0D, 0D);
            };
        var runs = new[] {
            new PdfTextRun("\t", fontSize: 24D, tabLeader: PdfTabLeaderStyle.Underscores),
            PdfTextRun.Normal(" ", fontSize: 12D),
            PdfTextRun.Tab()
        };
        object result = InvokePrivateFontMethod<object>("WrapRichRunsCoreWithFirstLineOrigin",
            runs, 288D, 12D, PdfStandardFont.Helvetica, 14.4D, null!, null!, 36D,
            TrailingTabInkOptions(), new[] {
                new PdfTabStop(72D, leader: PdfTabLeaderStyle.Underscores),
                new PdfTabStop(180D)
            }, frame, null!, null!, PdfTextWhitespaceMode.Preserve);

        // The existing private-owner harness admits a deterministic exclusion
        // frame: only the actual24pt source box narrows its available width.
        Assert.Contains(requests, request => request.Height >= 28.8D - .001D);
        var line = Assert.Single(ExtractLines(result));
        var leader = Assert.Single(line, segment => ExtractLeadingTabLeader(segment) == PdfTabLeaderStyle.Underscores);
        Assert.Equal(60D, ExtractLeadingAdvance(leader), 3);
        Assert.Equal(24D, (double)leader.GetType().GetProperty("FontSize")!.GetValue(leader)!, 3);
        var heights = (System.Collections.Generic.List<double>)result.GetType().GetField("Item2")!.GetValue(result)!;
        Assert.Equal(28.8D, Assert.Single(heights), 3);
    }
    private static PdfOptions TrailingTabInkOptions() {
        PdfOptions options = TrailingTabOptions();
        options.RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Arial",
            OfficeIMO.TestAssets.ManagedTextShapingTestAssets.CreateFontWithLineBoxMetrics(800, -200, 200, 200,
                Enumerable.Range(32, 95).ToArray())));
        return options;
    }
    private static PdfOptions TrailingTabOptions() => new PdfOptions {
        PageWidth = 360D, PageHeight = 220D,
        MarginLeft = 36D, MarginRight = 36D, MarginTop = 36D, MarginBottom = 36D,
        DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 12D
    };
}
