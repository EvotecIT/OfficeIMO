using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeIMO.Notebook.Html;

namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookCartesianChartExportTests {
    [Fact]
    public void DrawsNumericHeatmapGapsAndAlignedSubplotsFromSavedPlotlyOutputs() {
        using var notebook = NotebookDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "AdvancedOutputs",
            "Charts", "plotly-interactive.ipynb"));

        string html = notebook.ToHtml(null, out OfficeConversionReport report);

        XDocument[] figures = Regex.Matches(html, "src=\"data:image/svg\\+xml;base64,([^\"]+)\"").Cast<Match>()
            .Select(match => XDocument.Parse(Encoding.UTF8.GetString(Convert.FromBase64String(match.Groups[1].Value))))
            .ToArray();
        XElement heatmap = Assert.Single(figures, figure => figure.Root!.Value.Contains("Station temperature heatmap", StringComparison.Ordinal)).Root!;
        XElement subplots = Assert.Single(figures, figure => figure.Root!.Value.Contains("Field survey comparison", StringComparison.Ordinal)).Root!;
        Assert.Contains("Sampling hour", heatmap.Value, StringComparison.Ordinal);
        Assert.Contains("Depth", heatmap.Value, StringComparison.Ordinal);
        Assert.Contains("Temperature", heatmap.Value, StringComparison.Ordinal);

        // The independent Plotly.py fixture has four hours at three depths, with the upper row's
        // second cell unmeasured. Paint only the other eleven tiles, with the gap in that column.
        XNamespace svg = heatmap.Name.Namespace;
        XElement[] tiles = heatmap.Descendants(svg + "rect").Where(rectangle =>
            Number(rectangle, "width") is > 70 and < 125 && Number(rectangle, "height") is > 55 and < 110).ToArray();
        Assert.Equal(11, tiles.Length);
        IGrouping<double, XElement>[] rows = tiles.GroupBy(tile => Number(tile, "y")).OrderBy(row => row.Key).ToArray();
        Assert.Equal(new[] { 3, 4, 4 }, rows.Select(row => row.Count()).ToArray());
        double[] columnPositions = rows[1].Select(tile => Number(tile, "x")).OrderBy(x => x).ToArray();
        Assert.DoesNotContain(rows[0], tile => Number(tile, "x") == columnPositions[1]);

        XElement firstTitle = Assert.Single(subplots.Descendants(svg + "text"), text => text.Value == "Water temperature");
        XElement secondTitle = Assert.Single(subplots.Descendants(svg + "text"), text => text.Value == "Samples per station");
        Assert.True(Number(firstTitle, "x") < Number(secondTitle, "x"));
        Assert.Contains("Degrees Celsius", subplots.Value, StringComparison.Ordinal);
        Assert.Contains("Sample count", subplots.Value, StringComparison.Ordinal);
        foreach (string series in new[] { "Surface", "Day 1", "Day 2" }) {
            XElement label = Assert.Single(subplots.Descendants(svg + "text"), text => text.Value == series);
            Assert.True(Number(label, "x") > Number(secondTitle, "x"));
        }
        Assert.Contains(report.Diagnostics, item => item.Code == "notebook-output-chart-static-preview" &&
            item.Location?.Contains("outputs[", StringComparison.Ordinal) == true);
    }

    [Fact]
    public void SingleHeatmapRowKeepsItsSavedCellSpacing() {
        double tall = ReadHeatmapMark("""
            {"data":[{"type":"heatmap","z":[[5]],"y0":10,"dy":10,"showscale":false}],
             "layout":{"yaxis":{"range":[0,20]}}}
            """, attribute: 3);
        double shortRow = ReadHeatmapMark("""
            {"data":[{"type":"heatmap","z":[[5]],"y0":10,"dy":2,"showscale":false}],
             "layout":{"yaxis":{"range":[0,20]}}}
            """, attribute: 3);

        Assert.InRange(tall / shortRow, 4.99, 5.01);
    }

    [Fact]
    public void HeatmapCellGapDoesNotEraseTheVisiblePartOfACroppedCell() {
        double visibleWidth = ReadHeatmapMark("""
            {"data":[{"type":"heatmap","z":[[5]],"x":[0,10],"xgap":4,"showscale":false}],
             "layout":{"width":700,"xaxis":{"range":[9.85,40]}}}
            """, attribute: 2);

        Assert.InRange(visibleWidth, 0.1, 1.5);
    }

    private static double ReadHeatmapMark(string specification, int attribute) {
        using var notebook = NotebookDocument.Create();
        using var spec = JsonDocument.Parse(specification);
        notebook.AddCodeCell(string.Empty).AddOutput(NotebookOutput.CreateDisplayData(new Dictionary<string, JsonElement> {
            ["application/vnd.plotly.v1+json"] = spec.RootElement
        }));
        string html = notebook.ToHtml(new NotebookHtmlExportOptions { IncludeInteractiveCharts = true });
        Match metadata = Regex.Match(html, "<template data-officeimo-chart-marks>([^<]+)</template>");
        Assert.True(metadata.Success);
        using var marks = JsonDocument.Parse(metadata.Groups[1].Value);
        return Assert.Single(marks.RootElement.EnumerateArray())[attribute].GetDouble();
    }

    private static double Number(XElement element, string attribute) =>
        double.TryParse((string?)element.Attribute(attribute), NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
            ? number : double.NaN;
}
