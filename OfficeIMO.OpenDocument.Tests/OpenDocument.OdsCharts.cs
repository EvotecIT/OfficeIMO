using System;
using System.IO;
using System.Linq;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdsChartTests {
    [Fact]
    public void ReadsExcelProducedColumnChartWithoutChangingItsPackage() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "microsoft-excel-column-chart.ods");
        OdsDocument document = OdsDocument.Load(path);
        byte[] originalChartPart = document.GetPackageEntryBytes("Object 1/content.xml");

        OdsChart chart = Assert.Single(document.GetSheet("Data")!.Charts);
        Assert.Equal("Chart 1", chart.Name);
        Assert.Equal("chart:bar", chart.ChartClass);
        Assert.Equal("Sales", chart.Title);
        Assert.Equal("Data.$A$2:.$A$3", chart.CategoriesAddress);
        Assert.False(chart.IsStacked);
        Assert.False(chart.IsPercentage);
        Assert.False(chart.IsThreeDimensional);
        Assert.False(chart.VerticalBars);
        OdsChartSeries series = Assert.Single(chart.Series);
        Assert.Equal("Data.$B$2:.$B$3", series.ValuesAddress);
        Assert.Equal("Data.$B$1", series.LabelAddress);
        using var saved = new MemoryStream(document.Serialize().RequireValue());
        OdsDocument reopened = OdsDocument.Load(saved);
        Assert.Single(reopened.GetSheet("Data")!.Charts);
        Assert.Equal(originalChartPart, reopened.GetPackageEntryBytes("Object 1/content.xml"));
        Assert.Equal(OdfCapabilityLevel.Inspected, OdfCapabilityCatalog.Find("advanced-charts")!.Level);
    }
}
