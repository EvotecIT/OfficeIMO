using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Chart edits on a loaded presentation write the same chart part without the SDK as through it, and don't open it.</summary>
public class PowerPointChartsWithoutSdkTests {
    [Fact]
    public void ChartEditsWriteTheSameChartPartWithoutTheSdk() {
        byte[] source = CreateDeck();

        byte[] throughSdk;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            _ = presentation.OpenXmlDocument;
            Edit(presentation.Slides[0].Charts.Single());
            throughSdk = presentation.ToBytes();
        }
        byte[] withoutSdk;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            Edit(presentation.Slides[0].Charts.Single());
            Assert.False(presentation.IsSdkLoaded, string.Join("; ", presentation.Diagnostics.Select(diagnostic => diagnostic.Message)));
            withoutSdk = presentation.ToBytes();
        }

        Assert.Equal(PowerPointTablesWithoutSdkTests.Parts(throughSdk, "ppt/charts/chart"), PowerPointTablesWithoutSdkTests.Parts(withoutSdk, "ppt/charts/chart"));
    }

    private static void Edit(PowerPointChart chart) {
        chart.SetTitle("Revenue\nby quarter")
            .SetTitleTextStyle(fontSize: OfficeLength.Points(18), bold: true, color: OfficeColor.Accent2)
            .SetLegend(OfficeChartLegendPosition.Bottom, overlay: true)
            .SetLegendTextStyle(italic: true, fontFamily: "Arial")
            .SetValueAxisTitle("Amount")
            .SetValueAxisNumberFormat("0.0", sourceLinked: true)
            .SetValueAxisGridlines(showMajor: true, showMinor: true, lineColor: OfficeColor.FromRgb(0xC0, 0xC0, 0xC0), lineWidth: OfficeLength.Points(0.5))
            .SetCategoryAxisLabelRotation(-45)
            .SetValueAxisDisplayUnits(OfficeChartDisplayUnit.Thousands, "Thousands")
            .SetDataLabels(showValue: true, showCategoryName: true)
            .SetDataLabelPosition(OfficeChartDataLabelPosition.OutsideEnd)
            .SetDataLabelSeparator("|")
            .SetSeriesDataLabelForPoint(0, 1, showValue: false, numberFormat: "0%")
            .SetSeriesDataLabelShapeStyle(0, fillColor: OfficeColor.FromRgb(0xFF, 0xF2, 0xCC), noLine: true);
    }

    private static byte[] CreateDeck() {
        using var presentation = PowerPointDocument.Create();
        PowerPointSlide slide = presentation.AddSlide();
        slide.AddChart().SetTitle("Sales");
        return presentation.ToBytes();
    }
}
