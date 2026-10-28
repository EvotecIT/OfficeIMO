using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlOfficeAdapters {
    [Theory]
    [InlineData(OfficeChartType.Pie)]
    [InlineData(OfficeChartType.Doughnut)]
    public void ExcelHtml_RoundTripsRadialGeometry(OfficeChartType kind) {
        using ExcelDocument document = ExcelDocument.Create();
        document.AddWorksheet("Results").AddChart(kind,
            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeSeries("Status", new[] { 7d, 3d }) }), 1, 1)
            .SetRadialLayout(new OfficeChartRadialLayout(135, 75));
        using ExcelDocument restored = HtmlConversionDocument.Parse(document.ToHtml()).ToExcelDocument();
        var chart = restored.Sheets.SelectMany(s => s.Charts).Single();
        Assert.Equal(135, chart.RadialLayout.FirstSliceAngleDegrees);
        Assert.Equal(kind == OfficeChartType.Doughnut ? 75 : 50, chart.RadialLayout.DoughnutHolePercent);
        Assert.Empty(restored.ValidateDocument());
    }

    [Theory]
    [InlineData(OfficeChartType.Pie)]
    [InlineData(OfficeChartType.Doughnut)]
    public void PowerPointHtml_RoundTripsRadialGeometry(OfficeChartType kind) {
        using PowerPointDocument presentation = PowerPointDocument.Create();
        presentation.AddSlide().AddChart(kind,            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeSeries("Status", new[] { 7d, 3d }) }),global::OfficeIMO.OfficeLength.Points(10),global::OfficeIMO.OfficeLength.Points(10),global::OfficeIMO.OfficeLength.Points(300),global::OfficeIMO.OfficeLength.Points(200))
            .SetRadialLayout(new OfficeChartRadialLayout(135, 75));
        string html = presentation.ToHtml();
        using PowerPointDocument restored = HtmlConversionDocument.Parse(html).ToPowerPointDocument();
        var chart = restored.Slides.SelectMany(s => s.Charts).Single();
        Assert.Equal(135, chart.RadialLayout.FirstSliceAngleDegrees);
        Assert.Equal(kind == OfficeChartType.Doughnut ? 75 : 50, chart.RadialLayout.DoughnutHolePercent);
        Assert.Empty(restored.ValidateDocument());
    }

    [Theory]
    [InlineData("361", "75")]
    [InlineData("90", "9")]
    [InlineData("invalid", "50")]
    public void PowerPointHtml_InvalidRadialMetadataOmitsChartWithDiagnostic(string angle, string hole) {
        using PowerPointDocument presentation = PowerPointDocument.Create();
        presentation.AddSlide().AddChart(OfficeChartType.Doughnut,            new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("Status", new[] { 1d }) }),global::OfficeIMO.OfficeLength.Points(10),global::OfficeIMO.OfficeLength.Points(10),global::OfficeIMO.OfficeLength.Points(300),global::OfficeIMO.OfficeLength.Points(200))
            .SetRadialLayout(new OfficeChartRadialLayout(90, 75));
        string html = presentation.ToHtml().Replace("data-officeimo-first-slice-angle=\"90\"", "data-officeimo-first-slice-angle=\"" + angle + "\"")
            .Replace("data-officeimo-doughnut-hole=\"75\"", "data-officeimo-doughnut-hole=\"" + hole + "\"");
        var result = HtmlConversionDocument.Parse(html).ToPowerPointDocumentResult();
        using PowerPointDocument restored = result.Value;
        Assert.Empty(restored.Slides.SelectMany(s => s.Charts));
        Assert.Contains(result.Report.Diagnostics, d => d.Message.Contains("invalid radial geometry"));
    }
}
