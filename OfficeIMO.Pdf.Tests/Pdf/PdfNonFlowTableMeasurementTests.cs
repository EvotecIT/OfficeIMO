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
                story.AddTable(new[] { new[] { "OVERLAY" } }, OfficeAlignment.Right, OverlayStyle(60D));
                story.AddParagraph(paragraph => paragraph.AddText("FOOTER"), style: ParagraphStyle());
            }
            if (layout == "row") content.Row(row => row.PercentColumn(100D, AddStory));
            else if (layout == "panel") content.Panel(AddStory, new PdfPanelStyle {
                PaddingX = global::OfficeIMO.OfficeLength.Points(0), PaddingY = global::OfficeIMO.OfficeLength.Points(0), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), KeepTogether = true
            });
            else AddStory(content);
        }, global::OfficeIMO.OfficeLength.Points(36D)));
        document.Paragraph(paragraph => paragraph.AddText("BODY"));

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
                    if (overlay) story.AddTable(new[] { new[] { "OVERLAY" } }, OfficeAlignment.Right, OverlayStyle(36D));
                    story.AddParagraph(paragraph => paragraph.AddText("FOOTER"), style: ParagraphStyle());
                }
                if (layout == "row") content.Row(row => row.PercentColumn(100D, AddStory));
                else if (layout == "panel") content.Panel(AddStory, new PdfPanelStyle {
                    PaddingX = global::OfficeIMO.OfficeLength.Points(0), PaddingY = global::OfficeIMO.OfficeLength.Points(0), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), KeepTogether = true
                });
                else AddStory(content);
            }, global::OfficeIMO.OfficeLength.Points(48D)));
            return document.Paragraph(paragraph => paragraph.AddText("BODY")).ToBytes();
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
            content.AddTable(new[] { new[] { "OVERLAY" } }, style: OverlayStyle(18D)), global::OfficeIMO.OfficeLength.Points(48D)));
        else document.Footer(story => story.Content(content =>
            content.AddTable(new[] { new[] { "OVERLAY" } }, style: OverlayStyle(18D)), global::OfficeIMO.OfficeLength.Points(48D)));
        document.Paragraph(paragraph => paragraph.AddText("BODY"));

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
                flow.AddParagraph(paragraph => paragraph.AddText("LEAD"), style: new PdfParagraphStyle {
                    FontSize = global::OfficeIMO.OfficeLength.Points(10D), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(12D)),
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), KeepWithNext = true
                });
                flow.AddTable(new[] { new[] { "OVERLAY" } }, OfficeAlignment.Right, style);
                flow.AddParagraph(paragraph => paragraph.AddText("FOLLOWING"), style: ParagraphStyle());
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
                flow.AddParagraph(paragraph => paragraph.AddText("LEAD"), style: ParagraphStyle());
                flow.AddTable(new[] { new[] { "OVERLAY" } }, OfficeAlignment.Right, OverlayStyle(36D));
                flow.AddParagraph(paragraph => paragraph.AddText("FOLLOWING"), style: ParagraphStyle());
            }
            content.Row(columns => {
                if (grouping == "row") {
                    columns.Style(new PdfRowStyle { KeepTogether = true }).PercentColumn(100D, AddSequence);
                } else {
                    columns.PercentColumn(100D, flow => flow.Panel(AddSequence, new PdfPanelStyle {
                        PaddingX = global::OfficeIMO.OfficeLength.Points(0), PaddingY = global::OfficeIMO.OfficeLength.Points(0), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), KeepTogether = true
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
                    flow.AddParagraph(paragraph => paragraph.AddText("LEAD"), style: new PdfParagraphStyle {
                        FontSize = global::OfficeIMO.OfficeLength.Points(10D), LineSpacing = PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(12D)),
                        SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), KeepWithNext = !rowKeepsNext
                    });
                    if (rowKeepsNext) flow.AddTable(new[] { new[] { "OVERLAY" } }, OfficeAlignment.Right, OverlayStyle(36D));
                    else flow.Panel(panel => panel.AddTable(new[] { new[] { "OVERLAY" } }, OfficeAlignment.Right, OverlayStyle(36D)),
                        new PdfPanelStyle {
                            PaddingX = global::OfficeIMO.OfficeLength.Points(0), PaddingY = global::OfficeIMO.OfficeLength.Points(0), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0), KeepWithNext = true
                        });
                    if (!rowKeepsNext) flow.AddParagraph(paragraph => paragraph.AddText("FOLLOWING"), style: ParagraphStyle());
                });
            });
            if (rowKeepsNext) content.AddParagraph(paragraph => paragraph.AddText("FOLLOWING"), style: ParagraphStyle());
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
        PageSetup = {
            Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(300D), global::OfficeIMO.OfficeLength.Points(300D)),
            Margins = { Left = global::OfficeIMO.OfficeLength.Points(32D), Right = global::OfficeIMO.OfficeLength.Points(32D), Top = global::OfficeIMO.OfficeLength.Points(32D), Bottom = global::OfficeIMO.OfficeLength.Points(32D) }
        },
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
