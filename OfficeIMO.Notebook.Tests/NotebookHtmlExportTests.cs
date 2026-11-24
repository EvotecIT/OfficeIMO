using OfficeIMO.Notebook.Html;

namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookHtmlExportTests {
    [Fact]
    public async Task WritesStandaloneNativeMathAndEmbeddedVectorWithoutLoadingActiveContent() {
        using NotebookDocument notebook = NotebookDocument.Load(Encoding.UTF8.GetBytes("""
            {"nbformat":4,"nbformat_minor":5,"metadata":{},"cells":[
              {"id":"math","cell_type":"markdown","metadata":{},
               "source":"# Lab notebook\n\nThe ratio is $\\frac{1}{2}$.\n\n<img src=\"https://example.invalid/steal\">"},
              {"id":"figure","cell_type":"code","metadata":{},"execution_count":1,"source":"example()","outputs":[
                 {"output_type":"display_data","metadata":{},"data":{
                   "image/svg+xml":"<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"20\" height=\"19\"><rect width=\"20\" height=\"19\" fill=\"#cd201f\"/></svg>"}},
                 {"output_type":"display_data","metadata":{},"data":{
                   "application/javascript":"console.log('ACTIVE')","text/plain":"Safe output"}}
              ]}]}
            """));

        using var destination = new MemoryStream();
        OfficeConversionReport report = await notebook.SaveAsHtmlAsync(destination);
        string html = Encoding.UTF8.GetString(destination.ToArray());
        Assert.True(destination.CanWrite);
        Assert.Contains("<!DOCTYPE html>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<mfrac>", html, StringComparison.Ordinal);
        Assert.Contains("src=\"data:image/svg+xml;base64,", html, StringComparison.Ordinal);
        Assert.Contains("Safe output", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<script", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("src=\"https://example.invalid", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "notebook-output-static-fallback" && diagnostic.Location == "cells[1].outputs[1]");
    }

    [Theory]
    [InlineData("application/vnd.plotly.v1+json", """
        {"data":[{"type":"bar","x":["A","B"],"y":[2,3]}],"layout":{"title":"Reconstructable chart"}}
        """)]
    [InlineData("application/vnd.vegalite.v5+json", """
        {"data":{"values":[{"category":"A","value":2}]},"mark":"bar","title":"Reconstructable chart",
         "encoding":{"x":{"field":"category","type":"nominal"},"y":{"field":"value","type":"quantitative"}}}
        """)]
    public async Task PrefersSavedImageEvenWhenTheChartCanBeReconstructed(string chartMime, string specification) {
        using NotebookDocument notebook = NotebookDocument.Create();
        using JsonDocument chart = JsonDocument.Parse(specification);
        NotebookCell cell = notebook.AddCodeCell("plot()");
        cell.AddOutput(NotebookOutput.CreateDisplayData(new Dictionary<string, JsonElement> {
            [chartMime] = chart.RootElement,
            ["image/svg+xml"] = JsonSerializer.SerializeToElement(
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"20\" height=\"19\"><rect width=\"20\" height=\"19\"/></svg>")
        }));

        using var destination = new MemoryStream();
        OfficeConversionReport report = await notebook.SaveAsHtmlAsync(destination);
        string html = Encoding.UTF8.GetString(destination.ToArray());
        Assert.Contains("src=\"data:image/svg+xml;base64,", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Reconstructable chart", html, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == "notebook-output-chart-static-preview");
        Assert.Contains(report.Diagnostics, diagnostic =>
            diagnostic.Code == "notebook-output-static-fallback" &&
            diagnostic.Location == "cells[0].outputs[0]" &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }
}
