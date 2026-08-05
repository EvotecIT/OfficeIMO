using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public class HtmlPowerPointBubbleChartTests {
    [Fact]
    public void PowerPointHtml_RoundTripsBubbleChartLegendLayout() {
        using PowerPointPresentation presentation =
            PowerPointPresentation.Create(new MemoryStream());
        presentation.AddSlide().AddChart(
                OfficeChartKind.Bubble, CreateBubbleData())
            .SetTitle("Positioned")
            .SetLegend(OfficeIMO.PowerPoint.PowerPointChartLegendPosition.Top, overlay: true);
        presentation.AddSlide().AddChart(
                OfficeChartKind.Bubble, CreateBubbleData())
            .SetTitle("Hidden")
            .HideLegend();

        string html = presentation.ToHtml(
            new PowerPointHtmlSaveOptions {
                Profile = OfficeHtmlConversionProfile.PowerPointSemanticSlides
            });

        Assert.Contains("data-officeimo-show-legend=\"true\"", html,
            System.StringComparison.Ordinal);
        Assert.Contains("data-officeimo-legend-position=\"Top\"", html,
            System.StringComparison.Ordinal);
        Assert.Contains("data-officeimo-overlay-legend=\"true\"", html,
            System.StringComparison.Ordinal);
        Assert.Contains("data-officeimo-show-legend=\"false\"", html,
            System.StringComparison.Ordinal);

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult();
        using PowerPointPresentation imported = result.Value;
        Assert.True(imported.Slides[0].Charts.Single().TryGetOfficeSnapshot(
            out OfficeChartSnapshot positionedSnapshot));
        Assert.True(positionedSnapshot.Layout.ShowLegend);
        Assert.Equal(OfficeChartLegendPosition.Top,
            positionedSnapshot.Layout.LegendPosition);
        Assert.True(positionedSnapshot.Layout.OverlayLegend);

        Assert.True(imported.Slides[1].Charts.Single().TryGetOfficeSnapshot(
            out OfficeChartSnapshot hiddenSnapshot));
        Assert.False(hiddenSnapshot.Layout.ShowLegend);
        Assert.DoesNotContain(result.Report.Diagnostics,
            diagnostic =>
                diagnostic.Code ==
                    HtmlConversionDiagnosticCodes.ContentOmitted ||
                diagnostic.Code ==
                    HtmlConversionDiagnosticCodes.ContentApproximated);
    }

    private static OfficeChartData CreateBubbleData() =>
        new(new[] { "1", "2" }, new[] {
            OfficeChartSeries.CreateBubble(
                "Portfolio", new[] { 1D, 2D }, new[] { 3D, 4D },
                new[] { 5D, 6D })
        });
}
