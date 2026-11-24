using System.Text.RegularExpressions;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeIMO.Notebook.Html;
using OfficeIMO.Notebook.Pdf;
using OfficeIMO.Pdf;

namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookAdvancedOutputTests {
    [Fact]
    public void ExportsProducerPlotly3DAndWebGlFromSavedSpecifications() {
        using NotebookDocument notebook = LoadFixture("Charts", "plotly-advanced.ipynb");
        const string plotly = "application/vnd.plotly.v1+json";
        string[] titles = notebook.Cells.SelectMany(cell => cell.Outputs).Select(output => {
            Assert.Equal(plotly, Assert.Single(output.Data).Key);
            return output.Data[plotly].GetProperty("layout").GetProperty("title").GetProperty("text").GetString()!;
        }).ToArray();
        Assert.Equal(3, titles.Length);

        string html = notebook.ToHtml(null, out OfficeConversionReport htmlReport);
        Match[] previews = Regex.Matches(html, "src=\"data:image/svg\\+xml;base64,([^\"]+)\"").Cast<Match>().ToArray();
        Assert.Equal(titles.Length, previews.Length);
        for (int index = 0; index < titles.Length; index++) {
            string svg = Encoding.UTF8.GetString(Convert.FromBase64String(previews[index].Groups[1].Value));
            XElement scene = XElement.Parse(svg);
            Assert.Contains(titles[index], scene.Value, StringComparison.Ordinal);
            Assert.True(scene.Descendants().Count(element =>
                element.Name.LocalName is "path" or "circle" or "ellipse" or "polygon" or "polyline" or "use") > 5,
                titles[index] + " must contain painted chart geometry.");
        }

        using var destination = new MemoryStream();
        OfficeConversionReport pdfReport = notebook.SaveAsPdf(destination);
        using PdfDocument pdf = PdfDocument.Load(destination.ToArray());
        string pdfText = pdf.ExtractText();
        Assert.All(titles, title => Assert.Contains(title, pdfText, StringComparison.Ordinal));
        foreach (OfficeConversionReport report in new[] { htmlReport, pdfReport }) {
            Assert.Equal(3, report.Diagnostics.Count(item => item.Code == "notebook-output-chart-static-preview"));
            Assert.DoesNotContain(report.Diagnostics, item => item.Code is "notebook-output-unsupported" or "unsupported-image");
        }
    }

    [Fact]
    public void UsesSavedTextWhenAPlotlySurfaceDeclaresTheWrongShape() {
        const string mime = "application/vnd.plotly.v1+json";
        using NotebookDocument producer = LoadFixture("Charts", "plotly-advanced.ipynb");
        JsonElement saved = producer.Cells.Single(cell => cell.Id == "surface").Outputs[0].Data[mime];
        JsonNode specification = JsonNode.Parse(saved.GetRawText())!;
        specification["data"]![0]!["z"]!["shape"] = "7, 8";
        using NotebookDocument notebook = NotebookDocument.Create();
        notebook.AddCodeCell(string.Empty).AddOutput(NotebookOutput.CreateDisplayData(new Dictionary<string, JsonElement> {
            [mime] = JsonSerializer.SerializeToElement(specification),
            ["text/plain"] = JsonSerializer.SerializeToElement("The saved surface is available as a description.")
        }));

        string html = notebook.ToHtml(null, out OfficeConversionReport report);

        Assert.Contains("The saved surface is available as a description.", html, StringComparison.Ordinal);
        Assert.DoesNotContain("data:image/", html, StringComparison.Ordinal);
        Assert.Contains(report.Diagnostics, item => item.Code == "notebook-output-static-fallback");
    }

    [Fact]
    public void IncludesSavedStockControlsAndLinksOnlyInOptInHtml() {
        using NotebookDocument notebook = LoadFixture("Widgets", "widget-list.ipynb");

        string inert = notebook.ToHtml();
        string interactive = notebook.ToHtml(new NotebookHtmlExportOptions {
            IncludeInteractiveWidgets = true
        }, out _);

        Assert.DoesNotContain("<script", inert, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("data-officeimo-widget-id", inert, StringComparison.Ordinal);
        Assert.Contains("<input type=\"range\"", interactive, StringComparison.Ordinal);
        Assert.Contains("writing-mode:vertical-lr;direction:rtl", interactive, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-widget-link-from", interactive, StringComparison.Ordinal);
        Assert.Contains("<script", interactive, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<script src=", interactive, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ReportsExecutableCustomWidgetAndKeepsItsCodeOutOfOptInHtml() {
        using NotebookDocument notebook = LoadFixture("Widgets", "anywidget-counter.ipynb");

        string html = notebook.ToHtml(new NotebookHtmlExportOptions {
            IncludeInteractiveWidgets = true,
            ProjectionOptions = new OfficeIMO.Notebook.Markdown.NotebookMarkdownExportOptions { IncludeCode = false }
        }, out OfficeConversionReport report);

        Assert.DoesNotContain("<script", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("addEventListener", html, StringComparison.Ordinal);
        Assert.Equal(3, Regex.Matches(html, "CounterWidget object").Count);
        Assert.Equal(3, report.Diagnostics.Count(item => item.Code == "notebook-widget-view-unsupported"));
    }

    [Fact]
    public void KeepsControlsButReportsLinksBetweenIncompatibleNumericGrids() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "AdvancedOutputs", "Widgets", "widget-list.ipynb");
        JsonNode saved = JsonNode.Parse(File.ReadAllText(path))!;
        JsonObject state = saved["metadata"]!["widgets"]!["application/vnd.jupyter.widget-state+json"]!["state"]!.AsObject();
        JsonNode link = state.Single(entry => entry.Value!["model_name"]!.GetValue<string>() == "LinkModel").Value!["state"]!;
        JsonNode source = state[link["source"]![0]!.GetValue<string>().Substring("IPY_MODEL_".Length)]!["state"]!;
        JsonNode target = state[link["target"]![0]!.GetValue<string>().Substring("IPY_MODEL_".Length)]!["state"]!;
        source["step"] = 2;
        target["step"] = 3;
        source["value"] = 0;
        target["value"] = 0;
        using NotebookDocument notebook = NotebookDocument.Load(Encoding.UTF8.GetBytes(saved.ToJsonString()));

        string html = notebook.ToHtml(new NotebookHtmlExportOptions { IncludeInteractiveWidgets = true }, out OfficeConversionReport report);

        Assert.Contains("data-officeimo-widget-kind=\"play\"", html, StringComparison.Ordinal);
        Assert.Contains("<input type=\"range\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<span hidden data-officeimo-widget-link-from=", html, StringComparison.Ordinal);
        Assert.Contains(report.Diagnostics, item => item.Code == "notebook-widget-link-unsupported");
    }

    private static NotebookDocument LoadFixture(string family, string file) =>
        NotebookDocument.Load(Path.Combine(AppContext.BaseDirectory, "Fixtures", "AdvancedOutputs", family, file));
}
