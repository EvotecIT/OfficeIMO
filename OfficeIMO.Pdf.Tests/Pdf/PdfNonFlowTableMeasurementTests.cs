using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfNonFlowTableMeasurementTests {
    [Theory]
    [InlineData("direct")]
    [InlineData("row")]
    [InlineData("panel")]
    public void TallFooterOverlayReservesPhysicalPaintWhileFollowingTextUsesRestoredFlow(string layout) {
        var document = PdfDocument.Create(Options());
        document.Footer(footer => footer.Content(content => {
            void AddStory(PdfContentBuilder story) {
                story.Table(new[] { new[] { "OVERLAY" } }, PdfAlign.Right, OverlayStyle(60D));
                story.Paragraph(paragraph => paragraph.Text("FOOTER"), style: ParagraphStyle());
            }
            if (layout == "row") content.Row(row => row.PercentColumn(100D, AddStory));
            else if (layout == "panel") content.Panel(AddStory, new PdfPanelStyle {
                PaddingX = 0, PaddingY = 0, SpacingBefore = 0, SpacingAfter = 0, KeepTogether = true
            });
            else AddStory(content);
        }, 36D));
        document.Paragraph(paragraph => paragraph.Text("BODY"));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        var overlay = Assert.Single(page.GetWords(), word => word.Text == "OVERLAY");
        var footer = Assert.Single(page.GetWords(), word => word.Text == "FOOTER");
        Assert.Equal(overlay.Letters[0].StartBaseLine.Y, footer.Letters[0].StartBaseLine.Y, 3);
        Assert.True(footer.BoundingBox.Bottom >= 0D);
        Assert.Contains(page.GetWords(), word => word.Text == "BODY");
    }

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
    [InlineData(false, 182)]
    [InlineData(true, 182)]
    [InlineData(false, 194)]
    [InlineData(true, 194)]
    public void KeepWithNextOverlayTableReservesPaintWithoutAddingFlowAdvance(bool row, int leadingSpace) {
        var style = OverlayStyle(36D);
        style.KeepWithNext = true;
        var document = PdfDocument.Create(Options());
        document.Compose(compose => compose.Page(page => page.Content(content => {
            content.Spacer(global::OfficeIMO.OfficeLength.Points(leadingSpace));
            void AddSequence(PdfContentBuilder flow) {
                flow.Paragraph(paragraph => paragraph.Text("LEAD"), style: new PdfParagraphStyle {
                    FontSize = global::OfficeIMO.OfficeLength.Points(10D), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(12D)),
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), KeepWithNext = true
                });
                flow.Table(new[] { new[] { "OVERLAY" } }, PdfAlign.Right, style);
                flow.Paragraph(paragraph => paragraph.Text("FOLLOWING"), style: ParagraphStyle());
            }
            if (row) content.Row(columns => columns.PercentColumn(100D, AddSequence));
            else AddSequence(content);
        })));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var pages = pdf.GetPages().ToArray();
        if (leadingSpace == 182) Assert.Single(pages);
        var page = Assert.Single(pages, candidate => candidate.GetWords().Any(word => word.Text == "LEAD"));
        var lead = Assert.Single(page.GetWords(), word => word.Text == "LEAD");
        var following = Assert.Single(page.GetWords(), word => word.Text == "FOLLOWING");
        Assert.Contains(page.GetWords(), word => word.Text == "OVERLAY");
        Assert.Equal(12D, lead.Letters[0].StartBaseLine.Y - following.Letters[0].StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData("row")]
    [InlineData("panel")]
    public void KeptColumnGroupsRetainOverlayPaintWithoutAddingItsFlow(string grouping) {
        var document = PdfDocument.Create(Options());
        document.Compose(compose => compose.Page(page => page.Content(content => {
            content.Spacer(global::OfficeIMO.OfficeLength.Points(194D));
            void AddSequence(PdfContentBuilder flow) {
                flow.Paragraph(paragraph => paragraph.Text("LEAD"), style: ParagraphStyle());
                flow.Table(new[] { new[] { "OVERLAY" } }, PdfAlign.Right, OverlayStyle(36D));
                flow.Paragraph(paragraph => paragraph.Text("FOLLOWING"), style: ParagraphStyle());
            }
            content.Row(columns => {
                if (grouping == "row") {
                    columns.Style(new PdfRowStyle { KeepTogether = true }).PercentColumn(100D, AddSequence);
                } else {
                    columns.PercentColumn(100D, flow => flow.Panel(AddSequence, new PdfPanelStyle {
                        PaddingX = 0, PaddingY = 0, SpacingBefore = 0, SpacingAfter = 0, KeepTogether = true
                    }));
                }
            });
        })));
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = Assert.Single(pdf.GetPages(), candidate => candidate.GetWords().Any(word => word.Text == "LEAD"));
        var lead = Assert.Single(page.GetWords(), word => word.Text == "LEAD");
        var following = Assert.Single(page.GetWords(), word => word.Text == "FOLLOWING");
        Assert.Contains(page.GetWords(), word => word.Text == "OVERLAY");
        Assert.Equal(12D, lead.Letters[0].StartBaseLine.Y - following.Letters[0].StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData(false, 182)]
    [InlineData(false, 194)]
    [InlineData(true, 182)]
    [InlineData(true, 194)]
    public void KeptGroupsReserveOverlayExtentWhileFollowingTextUsesRestoredFlow(bool rowKeepsNext, int leadingSpace) {
        var document = PdfDocument.Create(Options());
        document.Compose(compose => compose.Page(page => page.Content(content => {
            content.Spacer(global::OfficeIMO.OfficeLength.Points(leadingSpace));
            content.Row(columns => {
                if (rowKeepsNext) columns.Style(new PdfRowStyle { KeepWithNext = true });
                columns.PercentColumn(100D, flow => {
                    flow.Paragraph(paragraph => paragraph.Text("LEAD"), style: new PdfParagraphStyle {
                        FontSize = global::OfficeIMO.OfficeLength.Points(10D), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(12D)),
                        SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), KeepWithNext = !rowKeepsNext
                    });
                    if (rowKeepsNext) flow.Table(new[] { new[] { "OVERLAY" } }, PdfAlign.Right, OverlayStyle(36D));
                    else flow.Panel(panel => panel.Table(new[] { new[] { "OVERLAY" } }, PdfAlign.Right, OverlayStyle(36D)),
                        new PdfPanelStyle {
                            PaddingX = 0, PaddingY = 0, SpacingBefore = 0, SpacingAfter = 0, KeepWithNext = true
                        });
                    if (!rowKeepsNext) flow.Paragraph(paragraph => paragraph.Text("FOLLOWING"), style: ParagraphStyle());
                });
            });
            if (rowKeepsNext) content.Paragraph(paragraph => paragraph.Text("FOLLOWING"), style: ParagraphStyle());
        })));
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var pages = pdf.GetPages().ToArray();
        if (leadingSpace == 182) Assert.Single(pages);
        var page = Assert.Single(pages, candidate => candidate.GetWords().Any(word => word.Text == "LEAD"));
        var lead = Assert.Single(page.GetWords(), word => word.Text == "LEAD");
        var following = Assert.Single(page.GetWords(), word => word.Text == "FOLLOWING");
        Assert.Contains(page.GetWords(), word => word.Text == "OVERLAY");
        Assert.Equal(12D, lead.Letters[0].StartBaseLine.Y - following.Letters[0].StartBaseLine.Y, 3);
    }

    private static PdfCreateOptions Options() => new() {
        PageWidth = global::OfficeIMO.OfficeLength.Points(300D), PageHeight = global::OfficeIMO.OfficeLength.Points(300D),
        MarginLeft = global::OfficeIMO.OfficeLength.Points(32D), MarginRight = global::OfficeIMO.OfficeLength.Points(32D), MarginTop = global::OfficeIMO.OfficeLength.Points(32D), MarginBottom = global::OfficeIMO.OfficeLength.Points(32D),
        DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10D)
    };

    private static PdfParagraphStyle ParagraphStyle() => new() {
        FontSize = global::OfficeIMO.OfficeLength.Points(10D), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(12D)), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
    };

    private static PdfTableStyle OverlayStyle(double height) => new() {
        HeaderRowCount = 0, ConsumesVerticalFlow = false, FontSize = global::OfficeIMO.OfficeLength.Points(10D),
        ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(80D) }, FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(height) },
        CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0), BorderWidth = global::OfficeIMO.OfficeLength.Points(0),
        SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
    };
}
