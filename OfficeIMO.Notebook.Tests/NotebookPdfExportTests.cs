using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Notebook.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookPdfExportTests {
    [Fact]
    public async Task SavesReadableNotebookPdfAndKeepsLongCodeTailAndResults() {
        string longCode = "codebegin_" + new string('a', 325) + "_codeend";
        using NotebookDocument notebook = LoadNotebook(longCode, """{"text/plain":"Scientific result"}""");
        using var destination = new MemoryStream();
        OfficeConversionReport report = await notebook.SaveAsPdfAsync(destination, new NotebookPdfExportOptions {
            MarkdownOptions = new MarkdownPdfExportOptions {
                ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
                TextFallbacks = PdfTextFallbackFeatures.None
            }
        });
        byte[] bytes = destination.ToArray();
        Assert.True(destination.CanWrite);
        Assert.StartsWith("%PDF", Encoding.ASCII.GetString(bytes, 0, 4), StringComparison.Ordinal);
        using PdfDocument written = PdfDocument.Load(bytes);
        string text = written.ExtractText();
        var lines = written.ExtractStructuredPages().SelectMany(page => page.LinesDetailed).ToArray();
        var first = Assert.Single(lines, line => line.Text.Contains("codebegin_", StringComparison.Ordinal));
        var last = Assert.Single(lines, line => line.Text.Contains("_codeend", StringComparison.Ordinal));
        Assert.NotEqual(first.Y, last.Y);
        Assert.All(lines, line => Assert.InRange(line.XEnd, 0, 550));
        Assert.Contains("codebegin_", text, StringComparison.Ordinal);
        Assert.Contains("_codeend", text, StringComparison.Ordinal);
        Assert.Contains("Scientific result", text, StringComparison.Ordinal);
        Assert.True(report.ItemsWritten > 0);
    }

    [Fact]
    public void StrictConversionLeavesDestinationUntouchedWhenOutputHasNoStaticRepresentation() {
        using NotebookDocument notebook = LoadNotebook("source", """
            {"application/vnd.jupyter.widget-view+json":{"version_major":2,"version_minor":0,"model_id":"store"}}
            """);
        using var destination = new MemoryStream();
        destination.WriteByte(0x41);
        var exception = Assert.Throws<OfficeConversionException>(() => notebook.SaveAsPdf(destination, new NotebookPdfExportOptions { FailOnLoss = true }));
        Assert.Equal(new byte[] { 0x41 }, destination.ToArray());
        Assert.Contains(exception.Report.Diagnostics, diagnostic => diagnostic.Location == "cells[0].outputs[0]" && diagnostic.Code == "notebook-output-unsupported");
    }

    [Fact]
    public void UsesSuppliedOpenTypeMathFaceWithSystemFontsDisabled() {
        using NotebookDocument notebook = LoadNotebook("calculation", """{"text/latex":"\\frac{2}{2}"}""");
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "FixtureMath.ttf"));
        var fontFaces = new OfficeFontFaceCollection().Add("math", font);
        var profile = new OfficeRenderingProfile("notebook-portable-math", fontFaces);
        using var destination = new MemoryStream();
        OfficeConversionReport report = notebook.SaveAsPdf(destination, new NotebookPdfExportOptions {
            IncludeExecutionCounts = false,
            IncludeCode = false,
            MarkdownOptions = new MarkdownPdfExportOptions {
                ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
                TextFallbacks = PdfTextFallbackFeatures.None,
                PdfOptions = new PdfCreateOptions().UseRenderingProfile(profile)
            }
        });

        using PdfDocument written = PdfDocument.Load(destination.ToArray());
        string text = written.ExtractText();
        Assert.Contains("2", text, StringComparison.Ordinal);
        Assert.DoesNotContain("frac", text, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == "math-font-unavailable");
    }

    [Fact]
    public void FitsAnOversizedVectorOutputAndRepeatsSavedTableHeadersAcrossPages() {
        using var notebook = NotebookDocument.Create();
        NotebookCell cell = notebook.AddCodeCell(string.Empty);
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='100' height='1600'>" +
            "<rect width='100' height='1600' fill='navy'/></svg>";
        string rows = string.Concat(Enumerable.Range(1, 30).Select(row => $"<tr><td>Result {row:D2}</td><td>{row}</td></tr>"));
        string table = "<table><thead><tr><th>Measurement</th><th>Value</th></tr></thead><tbody>" + rows + "</tbody></table>";
        using var figureOutput = JsonDocument.Parse(JsonSerializer.Serialize(svg));
        using var tableOutput = JsonDocument.Parse(JsonSerializer.Serialize(table));
        cell.AddOutput(NotebookOutput.CreateDisplayData(new Dictionary<string, JsonElement> { ["image/svg+xml"] = figureOutput.RootElement }));
        cell.AddOutput(NotebookOutput.CreateDisplayData(new Dictionary<string, JsonElement> { ["text/html"] = tableOutput.RootElement }));
        using var destination = new MemoryStream();
        var options = new NotebookPdfExportOptions {
            MarkdownOptions = new MarkdownPdfExportOptions {
                TextFallbacks = PdfTextFallbackFeatures.None,
                PdfOptions = new PdfCreateOptions { PageSetup = { Size = new PdfPageSize(OfficeLength.Points(300), OfficeLength.Points(280)) } }
            }
        };

        OfficeConversionReport report = notebook.SaveAsPdf(destination, options);

        using var pdf = PdfDocument.Load(destination.ToArray());
        var tablePages = pdf.ExtractStructuredPages()
            .Where(page => page.LinesDetailed.Any(line => line.Text.Contains("Result ", StringComparison.Ordinal)))
            .ToArray();
        Assert.True(tablePages.Length > 1);
        Assert.All(tablePages, page => Assert.Contains(page.LinesDetailed, line => line.Text.Contains("Measurement", StringComparison.Ordinal)));
        Assert.Contains("Result 30", pdf.ExtractText(), StringComparison.Ordinal);
        Assert.DoesNotContain(report.Diagnostics, item => item.Code == "unsupported-image" || item.Severity == OfficeDiagnosticSeverity.Error);
    }

    private static NotebookDocument LoadNotebook(string source, string data) => NotebookDocument.Load(Encoding.UTF8.GetBytes("""
        {"nbformat":4,"nbformat_minor":5,"metadata":{},"cells":[
          {"id":"calculation","cell_type":"code","metadata":{},"execution_count":1,"source":"SOURCE","outputs":[
             {"output_type":"display_data","metadata":{},"data":DATA}
          ]}]}
        """.Replace("SOURCE", source, StringComparison.Ordinal).Replace("DATA", data, StringComparison.Ordinal)));
}
