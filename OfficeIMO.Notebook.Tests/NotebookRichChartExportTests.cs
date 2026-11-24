using System.Text.RegularExpressions;
using OfficeIMO.Notebook.Html;
using OfficeIMO.Notebook.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookRichChartExportTests {
    [Fact]
    public void ExportsProducerChartGalleryToHtmlAndPdfWithStaticFidelityReports() {
        using NotebookDocument notebook = NotebookDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "rich-charts.ipynb"));

        string html = notebook.ToHtml(null, out OfficeConversionReport htmlReport);
        using var destination = new MemoryStream();
        OfficeConversionReport pdfReport = notebook.SaveAsPdf(destination);
        byte[] pdfBytes = destination.ToArray();

        Match[] figures = Regex.Matches(html, "src=\"data:(image/[^;]+);base64,([^\"]+)\"").Cast<Match>().ToArray();
        Assert.Equal(5, figures.Length);
        string[] vectors = figures.Where(figure => figure.Groups[1].Value == "image/svg+xml")
            .Select(figure => Encoding.UTF8.GetString(Convert.FromBase64String(figure.Groups[2].Value))).ToArray();
        Assert.Contains(vectors, svg => svg.Contains("Monthly revenue", StringComparison.Ordinal));
        Assert.Contains(vectors, svg => svg.Contains("Revenue observations", StringComparison.Ordinal));
        string vega = Assert.Single(vectors, svg => svg.Contains("Revenue by month", StringComparison.Ordinal));
        int january = vega.IndexOf(">Jan<", StringComparison.Ordinal);
        int february = vega.IndexOf(">Feb<", StringComparison.Ordinal);
        int march = vega.IndexOf(">Mar<", StringComparison.Ordinal);
        Assert.True(january >= 0 && january < february && february < march);
        Assert.Contains("<math", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<script", html, StringComparison.OrdinalIgnoreCase);

        using PdfDocument pdf = PdfDocument.Load(pdfBytes);
        Assert.Contains("Revenue and orders notebook", pdf.ExtractText(), StringComparison.Ordinal);
        Assert.NotEmpty(pdf.ExtractImages());
        foreach (OfficeConversionReport report in new[] { htmlReport, pdfReport }) {
            OfficeDiagnostic[] previews = report.Diagnostics.Where(item => item.Code == "notebook-output-chart-static-preview").ToArray();
            Assert.Equal(3, previews.Length);
            Assert.All(previews, item => {
                Assert.StartsWith("cells[", item.Location, StringComparison.Ordinal);
                Assert.Equal(OfficeConversionLossKind.Approximation, item.LossKind);
            });
            Assert.Equal(2, report.Diagnostics.Count(item => item.Code == "notebook-output-static-fallback"));
            Assert.DoesNotContain(report.Diagnostics, item => item.Code is "notebook-output-unsupported" or "unsupported-image");
        }
    }

    [Theory]
    [InlineData("{\"url\":\"https://example.invalid/chart.json\"}", "null")]
    [InlineData("{\"values\":[{\"category\":\"A\",\"value\":1}]}", "42")]
    [InlineData("{\"values\":[{\"category\":\"\\uD800\",\"value\":1}]}", "null")]
    [InlineData("{\"values\":[{\"category\":\"\\uFFFF\",\"value\":1}]}", "null")]
    public void UsesStoredTextForRemoteDataAndMalformedChartProperties(string data, string sort) {
        AssertStoredDescription("application/vnd.vegalite.v6.json", $$$"""
            {"data":{{{data}}},"mark":"bar","encoding":{
              "x":{"field":"category","type":"nominal","sort":{{{sort}}}},
              "y":{"field":"value","type":"quantitative"}
            }}
            """);
    }

    [Fact]
    public void UsesStoredTextWhenAPlotlyTemplateChangesTheActiveAxisScale() {
        AssertStoredDescription("application/vnd.plotly.v1+json", """
            {"data":[{"type":"scatter","mode":"markers","x":[1,10,100],"y":[4,9,5]}],
             "layout":{"template":{"layout":{"xaxis":{"type":"log"}}}}}
            """);
    }

    private static void AssertStoredDescription(string mime, string specification) {
        using var notebook = NotebookDocument.Create();
        using JsonDocument spec = JsonDocument.Parse(specification);
        notebook.AddCodeCell(string.Empty).AddOutput(NotebookOutput.CreateDisplayData(new Dictionary<string, JsonElement> {
            [mime] = spec.RootElement,
            ["text/plain"] = JsonSerializer.SerializeToElement("Stored chart description")
        }));

        string html = notebook.ToHtml(null, out OfficeConversionReport report);

        Assert.Contains("Stored chart description", html, StringComparison.Ordinal);
        Assert.DoesNotContain("https://example.invalid", html, StringComparison.Ordinal);
        Assert.DoesNotContain("data:image/", html, StringComparison.Ordinal);
        Assert.Contains(report.Diagnostics, item => item.Code == "notebook-output-static-fallback");
    }
}
