using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfNonFlowTableMeasurementTests {
    [Theory]
    [InlineData("direct")]
    [InlineData("row")]
    [InlineData("panel")]
    public void FooterOverlayTableDoesNotShiftFollowingFlow(string layout) {
        byte[] Render(bool overlay) {
            var document = PdfDocument.Create(Options());
            document.Footer(footer => footer.Content(content => {
                void AddStory(PdfContentBuilder story) {
                    if (overlay) story.Table(new[] { new[] { "OVERLAY" } }, PdfAlign.Right, OverlayStyle(36D));
                    story.Paragraph(paragraph => paragraph.Text("FOOTER"), style: ParagraphStyle());
                }
                if (layout == "row") content.Row(row => row.PercentColumn(100D, AddStory));
                else if (layout == "panel") content.Panel(AddStory, new PdfPanelStyle {
                    PaddingX = 0, PaddingY = 0, SpacingBefore = 0, SpacingAfter = 0, KeepTogether = true
                });
                else AddStory(content);
            }, 48D));
            return document.Paragraph(paragraph => paragraph.Text("BODY")).ToBytes();
        }

        using var control = PdfPigDocument.Open(Render(false));
        using var actual = PdfPigDocument.Open(Render(true));
        var expectedPage = Assert.Single(control.GetPages());
        var page = Assert.Single(actual.GetPages());
        var expected = Assert.Single(expectedPage.GetWords(), word => word.Text == "FOOTER");
        var footer = Assert.Single(page.GetWords(), word => word.Text == "FOOTER");
        Assert.Contains(page.GetWords(), word => word.Text == "OVERLAY");
        Assert.Equal(expected.Letters[0].StartBaseLine.Y, footer.Letters[0].StartBaseLine.Y, 3);
        Assert.Equal(expected.BoundingBox.Left, footer.BoundingBox.Left, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ZeroFlowTableOnlyRunningContentStillPaints(bool header) {
        var document = PdfDocument.Create(Options());
        if (header) document.Header(story => story.Content(content =>
            content.Table(new[] { new[] { "OVERLAY" } }, style: OverlayStyle(18D)), 48D));
        else document.Footer(story => story.Content(content =>
            content.Table(new[] { new[] { "OVERLAY" } }, style: OverlayStyle(18D)), 48D));
        document.Paragraph(paragraph => paragraph.Text("BODY"));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains(page.GetWords(), word => word.Text == "OVERLAY");
        Assert.Contains(page.GetWords(), word => word.Text == "BODY");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void KeepWithNextOverlayTableReservesPaintWithoutAddingFlowAdvance(bool row) {
        var style = OverlayStyle(36D);
        style.KeepWithNext = true;
        var document = PdfDocument.Create(Options());
        document.Compose(compose => compose.Page(page => page.Content(content => {
            content.Spacer(182D);
            void AddSequence(PdfContentBuilder flow) {
                flow.Paragraph(paragraph => paragraph.Text("LEAD"), style: new PdfParagraphStyle {
                    FontSize = 10D, LineSpacing = PdfLineSpacing.Exactly(12D),
                    SpacingBefore = 0, SpacingAfter = 0, KeepWithNext = true
                });
                flow.Table(new[] { new[] { "OVERLAY" } }, PdfAlign.Right, style);
                flow.Paragraph(paragraph => paragraph.Text("FOLLOWING"), style: ParagraphStyle());
            }
            if (row) content.Row(columns => columns.PercentColumn(100D, AddSequence));
            else AddSequence(content);
        })));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        var lead = Assert.Single(page.GetWords(), word => word.Text == "LEAD");
        var following = Assert.Single(page.GetWords(), word => word.Text == "FOLLOWING");
        Assert.Contains(page.GetWords(), word => word.Text == "OVERLAY");
        Assert.Equal(12D, lead.Letters[0].StartBaseLine.Y - following.Letters[0].StartBaseLine.Y, 3);
    }

    private static PdfOptions Options() => new() {
        PageWidth = 300D, PageHeight = 300D,
        MarginLeft = 32D, MarginRight = 32D, MarginTop = 32D, MarginBottom = 32D,
        DefaultFontSize = 10D
    };

    private static PdfParagraphStyle ParagraphStyle() => new() {
        FontSize = 10D, LineSpacing = PdfLineSpacing.Exactly(12D), SpacingBefore = 0, SpacingAfter = 0
    };

    private static PdfTableStyle OverlayStyle(double height) => new() {
        HeaderRowCount = 0, ConsumesVerticalFlow = false, FontSize = 10D,
        ColumnWidthPoints = new List<double?> { 80D }, FixedRowHeights = new List<double?> { height },
        CellPaddingX = 0, CellPaddingY = 0, BorderWidth = 0,
        SpacingBefore = 0, SpacingAfter = 0
    };
}
