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

    private static PdfOptions TrailingTabOptions() => new PdfOptions {
        PageWidth = 360D, PageHeight = 220D,
        MarginLeft = 36D, MarginRight = 36D, MarginTop = 36D, MarginBottom = 36D,
        DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 12D
    };
}
