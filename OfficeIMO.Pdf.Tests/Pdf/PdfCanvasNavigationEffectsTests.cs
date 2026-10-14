using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCanvasNavigationEffectsTests {
    [Theory]
    [InlineData(1D)]
    [InlineData(0.5D)]
    public void NamedDestinationsAndOutlinesFollowNestedCanvasEffects(double opacity) {
        byte[] bytes = PdfDocument.Create().Page(page => page.Size(200D, 200D).Margin(0D)
            .Canvas(canvas => canvas.Effect(OfficeTransform.Translate(5D, 7D), opacity,
                outer => outer.Effect(OfficeTransform.Scale(2D, 2D), 1D,
                    inner => inner.NamedDestination("target", 10D, 20D)
                        .Outline("Target", 1, 20D)
                        .LinkToNamedDestination("target", 10D, 20D, 10D, 10D)))))
            .ToBytes();
        PdfNamedDestination destination = Assert.Single(PdfDocumentReadResult.Load(bytes).NamedDestinations);
        var link = Assert.Single(PdfReadDocument.Open(bytes).Pages[0].GetLinkAnnotations());
        PdfOutlineItem outline = Assert.Single(PdfInspector.Inspect(bytes).Outlines);

        Assert.Equal(1, destination.PageNumber);
        Assert.Equal(25D, destination.DestinationLeft!.Value, 6);
        Assert.Equal(153D, destination.DestinationTop!.Value, 6);
        Assert.Equal(destination.DestinationLeft.Value, link.X1, 6);
        Assert.Equal(destination.DestinationTop.Value, link.Y2, 6);
        Assert.Equal(1, outline.PageNumber);
        Assert.Equal(5D, outline.DestinationLeft!.Value, 6);
        Assert.Equal(destination.DestinationTop.Value, outline.DestinationTop!.Value, 6);
    }

    [Theory]
    [InlineData(1D, 0D, 1D)]
    [InlineData(1D, 1D, 0D)]
    [InlineData(0.5D, 0D, 1D)]
    [InlineData(0.5D, 1D, 0D)]
    public void SingularEffectsOmitInvisibleLinksAndPreserveNavigationTargets(double opacity, double xScale, double yScale) {
        byte[] bytes = PdfDocument.Create().Page(page => page.Size(200D, 200D).Margin(0D)
            .Canvas(canvas => canvas.LinkToNamedDestination("target", 5D, 5D, 10D, 10D)
                .TextAnnotation("Visible", 5D, 30D, 10D, 10D)
                .Effect(OfficeTransform.Scale(xScale, yScale), opacity, content => content
                    .NamedDestination("target", 10D, 20D)
                    .LinkToNamedDestination("target", 10D, 20D, 10D, 10D)
                    .Text(new[] { PdfTextRun.Link("Hidden", "https://example.test/invisible") }, 10D, 20D, 10D, 10D)
                    .TextAnnotation("Hidden", 10D, 20D, 10D, 10D)
                    .FreeTextAnnotation("Hidden", 10D, 20D, 10D, 10D)
                    .HighlightAnnotation("Hidden", 10D, 20D, 10D, 10D))))
            .ToBytes();

        Assert.Single(PdfDocumentReadResult.Load(bytes).NamedDestinations);
        var link = Assert.Single(PdfReadDocument.Open(bytes).Pages[0].GetLinkAnnotations());
        Assert.Equal("target", link.DestinationName);
        Assert.Equal(2, PdfReadDocument.Open(bytes).Pages[0].GetAnnotations().Count);
    }
}
