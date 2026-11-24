using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using OfficeIMO.Notebook.Html;

namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookMeshAndMapExportTests {
    [Fact]
    public void AntimeridianRouteClosesAtBothMapEdgesWithoutCrossingTheWorld() {
        using var notebook = CreatePlotlyNotebook("""
            {"data":[{"type":"scattergeo","lon":[170,-170],"lat":[11,13],"mode":"lines"}],
             "layout":{"geo":{"showland":false,"showcoastlines":false,"showframe":false}}}
            """);

        XDocument map = ExtractSvg(notebook.ToHtml());
        XNamespace svg = "http://www.w3.org/2000/svg";
        (double low, double high)[] spans = map.Descendants(svg + "line").Select(line => {
            double from = double.Parse(line.Attribute("x1")!.Value, CultureInfo.InvariantCulture);
            double to = double.Parse(line.Attribute("x2")!.Value, CultureInfo.InvariantCulture);
            return (Math.Min(from, to), Math.Max(from, to));
        }).OrderBy(line => line.Item1).ToArray();

        Assert.Equal(2, spans.Length);
        Assert.All(spans, span => Assert.InRange(span.high - span.low, 12d, 15d));
        Assert.True(spans[1].low - spans[0].high > 450, "The pieces should land at opposite sides of the map.");
    }

    [Fact]
    public void CountryChoroplethUsesEmbeddedShapesAndKeepsRegionValuesInspectable() {
        using var notebook = CreatePlotlyNotebook("""
            {"data":[{"type":"choropleth","locations":["POL","FRA"],"locationmode":"ISO-3","z":[12,28],
              "colorscale":[[0,"#0a3351"],[1,"#e6b63d"]]}],
             "layout":{"geo":{"fitbounds":"locations","showland":false,"showcoastlines":false}}}
            """);

        string html = notebook.ToHtml(new NotebookHtmlExportOptions { IncludeInteractiveCharts = true },
            out OfficeConversionReport report);
        Match metadata = Regex.Match(html, "<template data-officeimo-chart-marks>([^<]+)</template>");
        Assert.True(metadata.Success);
        using var marks = JsonDocument.Parse(metadata.Groups[1].Value);
        Assert.Collection(marks.RootElement.EnumerateArray(),
            mark => Assert.Equal("Poland • Value: 12", mark[4].GetString()),
            mark => Assert.Equal("France • Value: 28", mark[4].GetString()));
        XNamespace svg = "http://www.w3.org/2000/svg";
        Assert.True(ExtractSvg(html).Descendants(svg + "path").Count() >= 2);
        Assert.Contains(report.Diagnostics, item => item.Code == "notebook-output-chart-interactive-preview");
    }

    [Fact]
    public void MeshWithAnOutOfBoundsTriangleUsesItsStoredTextAlternative() {
        using var notebook = CreatePlotlyNotebook("""
            {"data":[{"type":"mesh3d","x":[0,1,0],"y":[0,0,1],"z":[0,1,1],"i":[0],"j":[1],"k":[3]}]}
            """);

        string html = notebook.ToHtml(null, out OfficeConversionReport report);

        Assert.Contains("Stored chart description", html, StringComparison.Ordinal);
        Assert.DoesNotContain("data:image/svg+xml", html, StringComparison.Ordinal);
        Assert.Contains(report.Diagnostics, item => item.Code == "notebook-output-static-fallback");
    }

    [Fact]
    public void ScalarMeshShowsItsSavedColorbarTitleByDefault() {
        using var notebook = CreatePlotlyNotebook("""
            {"data":[{"type":"mesh3d","x":[0,1,0],"y":[0,0,1],"z":[0,1,1],"i":[0],"j":[1],"k":[2],
              "intensity":[2,4,6],"colorbar":{"title":{"text":"Signal intensity"}}}]}
            """);

        XDocument mesh = ExtractSvg(notebook.ToHtml());

        Assert.Contains("Signal intensity", mesh.Root!.Value, StringComparison.Ordinal);
    }

    private static XDocument ExtractSvg(string html) {
        Match svg = Regex.Match(html, "data:image/svg\\+xml;base64,([^\"]+)\"");
        Assert.True(svg.Success);
        return XDocument.Parse(Encoding.UTF8.GetString(Convert.FromBase64String(svg.Groups[1].Value)));
    }

    private static NotebookDocument CreatePlotlyNotebook(string specification) {
        using var spec = JsonDocument.Parse(specification);
        var notebook = NotebookDocument.Create();
        notebook.AddCodeCell(string.Empty).AddOutput(NotebookOutput.CreateDisplayData(new Dictionary<string, JsonElement> {
            ["application/vnd.plotly.v1+json"] = spec.RootElement,
            ["text/plain"] = JsonSerializer.SerializeToElement("Stored chart description")
        }));
        return notebook;
    }
}
