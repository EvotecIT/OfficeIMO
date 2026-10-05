using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfColumnKeepReservationTests {
    [Theory]
    [InlineData("paragraph", false, false)]
    [InlineData("paragraph", true, false)]
    [InlineData("paragraph", false, true)]
    [InlineData("list", false, false)]
    [InlineData("list", true, false)]
    [InlineData("heading", false, false)]
    [InlineData("heading", true, false)]
    [InlineData("panel", false, false)]
    public void Columns_BalancedKeepNextLeadersRetainTheirTerminalGroupWithoutEmptyPageRetries(
        string kind, bool interveningBookmark, bool balanceKeptLines) {
        byte[] bytes = PdfDocument.Create(new PdfOptions { PageWidth = 500, PageHeight = 400,
            MarginLeft = 40, MarginRight = 40, MarginTop = 40, MarginBottom = 40, DefaultFontSize = 12,
            MaxGeneratedPages = 4 }).Columns(content => {
                if (kind == "heading") content.H1("Heading001", style: new PdfHeadingStyle {
                    FontSize = 12, LineSpacing = PdfLineSpacing.Exactly(20), SpacingBefore = 0, SpacingAfter = 0,
                    KeepWithNext = true });
                if (kind == "list") content.RichNumbered(Enumerable.Range(1, 3).Select(index => new PdfListItem($"Lead{index:D3}")),
                    style: new PdfListStyle { LineSpacing = PdfLineSpacing.Exactly(20), ItemSpacing = 0,
                        SpacingBefore = 0, SpacingAfter = 0, KeepWithNext = true });
                else if (kind == "panel") content.Panel(panel => panel.Paragraph(p => p.Text("Lead001\nLead002\nLead003"), style: Style()),
                    new PdfPanelStyle { PaddingX = 0, PaddingY = 0, SpacingBefore = 0, SpacingAfter = 0, KeepWithNext = true });
                else {
                    var leader = Style(); leader.KeepWithNext = true;
                    content.Paragraph(p => p.Text("Lead001\nLead002\nLead003"), style: leader);
                }
                if (interveningBookmark) content.Bookmark("kept-next");
                var body = Style(); body.KeepTogether = true;
                content.Paragraph(p => p.Text("Body001\nBody002\nBody003\nBody004"), style: body);
            }, new PdfMultiColumnOptions { Gap = 20, BalanceLastPage = true, BalanceKeptParagraphLines = balanceKeptLines }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        var words = pdf.GetPage(1).GetWords().ToArray();
        foreach (string prefix in new[] { "Lead", "Body" }) {
            int count = prefix == "Lead" ? 3 : 4;
            for (int index = 1; index <= count; index++)
                Assert.Single(words, word => word.Text.EndsWith($"{prefix}{index:D3}", StringComparison.Ordinal));
        }
        double firstBodyX = words.Single(word => word.Text == "Body001").BoundingBox.Left;
        double lastBodyX = words.Single(word => word.Text == "Body004").BoundingBox.Left;
        Assert.InRange(firstBodyX, kind == "panel" ? 39.9 : 259.9, kind == "panel" ? 40.1 : 260.1);
        Assert.Equal(firstBodyX, lastBodyX, 2);
        double lastLeaderX = words.Single(word => word.Text.EndsWith("Lead003", StringComparison.Ordinal)).BoundingBox.Left;
        Assert.True(kind == "panel" ? lastLeaderX < 200 : lastLeaderX > 250);
    }

    private static PdfParagraphStyle Style() => new() { LineSpacing = PdfLineSpacing.Exactly(20),
        SpacingBefore = 0, SpacingAfter = 0, WidowControl = false };
}
