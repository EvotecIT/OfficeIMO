using System.Text.RegularExpressions;
using OfficeIMO.Notebook.Html;

namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookInteractiveChartExportTests {
    [Fact]
    public void ExportsProducerChartsWithIsolatedOfflineViewportsAndReportsTheApproximation() {
        using var notebook = NotebookDocument.Load(Path.Combine(AppContext.BaseDirectory,
            "Fixtures", "AdvancedOutputs", "Charts", "plotly-interactive.ipynb"));

        string html = notebook.ToHtml(new NotebookHtmlExportOptions { IncludeInteractiveCharts = true },
            out OfficeConversionReport report);

        Assert.Equal(5, Regex.Matches(html, "<figure data-officeimo-chart>").Count);
        Assert.Equal(5, Regex.Matches(html, "<image width=\"[^\"]+\" height=\"[^\"]+\" href=\"data:image/svg\\+xml;base64,").Count);
        Assert.Equal(5, Regex.Matches(html, "<template data-officeimo-chart-marks>").Count);
        Assert.Single(Regex.Matches(html, "<script>").Cast<Match>());
        Assert.DoesNotContain(report.Diagnostics, item => item.Code is "notebook-output-unsupported" or "notebook-output-static-fallback");
        Assert.Equal(5, report.Diagnostics.Count(item => item.Code == "notebook-output-chart-interactive-preview"
            && item.LossKind == OfficeConversionLossKind.Approximation));
    }

    [Fact]
    public void PreservesQuotedSavedValuesAsDataInAnInteractiveFragment() {
        using var notebook = NotebookDocument.Create();
        using var spec = JsonDocument.Parse("""
            {"data":[{"type":"scatter","mode":"markers","name":"A \"station\" & B","x":[1,3],"y":[2,4]}],
             "layout":{"title":{"text":"A \"station\" & B"}}}
            """);
        notebook.AddCodeCell(string.Empty).AddOutput(NotebookOutput.CreateDisplayData(new Dictionary<string, JsonElement> {
            ["application/vnd.plotly.v1+json"] = spec.RootElement
        }));

        string html = notebook.ToHtml(new NotebookHtmlExportOptions { IncludeInteractiveCharts = true, Fragment = true });

        Assert.Contains("aria-label=\"A &quot;station&quot; &amp; B\"", html, StringComparison.Ordinal);
        Match markData = Regex.Match(html, "<template data-officeimo-chart-marks>([^<]+)</template>");
        Assert.True(markData.Success);
        using var marks = JsonDocument.Parse(markData.Groups[1].Value);
        Assert.Collection(marks.RootElement.EnumerateArray(),
            point => Assert.Equal("A \"station\" & B • x: 1 • y: 2", point[4].GetString()),
            point => Assert.Equal("A \"station\" & B • x: 3 • y: 4", point[4].GetString()));
    }

    [Fact]
    public void HeatmapAutorangeUsesSavedOrientationAndRecomputesItsBounds() {
        using var notebook = NotebookDocument.Create();
        using var spec = JsonDocument.Parse("""
            {"data":[{"type":"heatmap","x":[0,10],"y":[0],"z":[[1,2]]}],
             "layout":{"xaxis":{"range":[1,0],"autorange":true}}}
            """);
        notebook.AddCodeCell(string.Empty).AddOutput(NotebookOutput.CreateDisplayData(new Dictionary<string, JsonElement> {
            ["application/vnd.plotly.v1+json"] = spec.RootElement
        }));

        string html = notebook.ToHtml(new NotebookHtmlExportOptions { IncludeInteractiveCharts = true });
        Match metadata = Regex.Match(html, "<template data-officeimo-chart-marks>([^<]+)</template>");
        Assert.True(metadata.Success);
        using var marks = JsonDocument.Parse(metadata.Groups[1].Value);
        Assert.Equal(2, marks.RootElement.GetArrayLength());
        Assert.True(marks.RootElement[0][0].GetDouble() > marks.RootElement[1][0].GetDouble());
    }

    [Fact]
    public void SkipsChartInteractivityWhenOutputsAreHiddenOrUseUnsupportedData() {
        using var notebook = NotebookDocument.Create();
        using var spec = JsonDocument.Parse("""
            {"data":[{"type":"choropleth","locations":["one"],"z":[1],"geojson":"https://example.invalid/regions.json"}]}
            """);
        notebook.AddCodeCell(string.Empty).AddOutput(NotebookOutput.CreateDisplayData(new Dictionary<string, JsonElement> {
            ["application/vnd.plotly.v1+json"] = spec.RootElement,
            ["text/plain"] = JsonSerializer.SerializeToElement("Saved map description")
        }));
        var options = new NotebookHtmlExportOptions { IncludeInteractiveCharts = true };

        string fallback = notebook.ToHtml(options, out OfficeConversionReport report);
        Assert.Contains("Saved map description", fallback, StringComparison.Ordinal);
        Assert.DoesNotContain("example.invalid", fallback, StringComparison.Ordinal);
        Assert.DoesNotContain("<script>", fallback, StringComparison.Ordinal);
        Assert.Contains(report.Diagnostics, item => item.Code == "notebook-output-static-fallback");

        using var producer = NotebookDocument.Load(Path.Combine(AppContext.BaseDirectory,
            "Fixtures", "AdvancedOutputs", "Charts", "plotly-interactive.ipynb"));
        options.ProjectionOptions.IncludeOutputs = false;
        string hidden = producer.ToHtml(options);
        Assert.DoesNotContain("data-officeimo-chart", hidden, StringComparison.Ordinal);
        Assert.DoesNotContain("<script>", hidden, StringComparison.Ordinal);
    }
}
