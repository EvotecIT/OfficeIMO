using System.Net;
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

    [Fact]
    public void EmitsVisibleFencesForNestedDelimiterListsFunctionsAndMatrices() {
        using NotebookDocument notebook = NotebookDocument.Create();
        notebook.AddMarkdownCell(@"$2\left(x+\sin{y}\right)$ $\left[a\mathpunct{,}b\right]$" +
            @" $\begin{bmatrix}1&2\\3\end{bmatrix}$ $\begin{cases}1&x\\2&y\end{cases}$");

        string html = notebook.ToHtml(new NotebookHtmlExportOptions { FailOnLoss = true }, out _);

        Assert.DoesNotContain("<mfenced", html, StringComparison.Ordinal);
        Assert.Contains("<mo fence=\"true\" stretchy=\"true\" form=\"prefix\">(</mo><mrow><mi>x</mi>", html, StringComparison.Ordinal);
        Assert.Contains("<mi>sin</mi><mo>⁡</mo><mrow><mo fence=\"true\" stretchy=\"true\" form=\"prefix\">(</mo><mi>y</mi>" +
            "<mo fence=\"true\" stretchy=\"true\" form=\"postfix\">)</mo>", html, StringComparison.Ordinal);
        Assert.Contains("<mi>a</mi><mo separator=\"true\">,</mo><mi>b</mi>", html, StringComparison.Ordinal);
        Assert.Contains("<mo fence=\"true\" stretchy=\"true\" form=\"prefix\">[</mo><mtable>", html, StringComparison.Ordinal);
        Assert.Contains("</mtable><mo fence=\"true\" stretchy=\"true\" form=\"postfix\">]</mo>", html, StringComparison.Ordinal);
        Assert.Contains("<mo fence=\"true\" stretchy=\"true\" form=\"prefix\">{</mo><mtable>", html, StringComparison.Ordinal);
        Assert.Contains("</mtable></mrow>", html, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(192, 1)]
    [InlineData(128, 3)]
    public void PreservesSparseMatrixSourceWhenRectangularExpansionExceedsTheExpressionLimit(int size, int matrixCount) {
        string matrix = @"\begin{bmatrix}x" + new string('&', size - 1) +
            string.Concat(Enumerable.Repeat(@"\\y", size - 1)) + @"\end{bmatrix}";
        string source = string.Join("+", Enumerable.Repeat(@"\left(" + matrix + @"\right)", matrixCount));
        using NotebookDocument notebook = NotebookDocument.Create();
        notebook.AddMarkdownCell("$" + source + "$");

        string html = notebook.ToHtml(null, out OfficeConversionReport report);

        Assert.Contains("<code>" + WebUtility.HtmlEncode(source) + "</code>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<mtable", html, StringComparison.Ordinal);
        Assert.Single(report.Diagnostics, diagnostic => diagnostic.Code == "notebook-html-math-unsupported" &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void PreservesUnrecognizedLatexSourceAndReportsFidelityLoss() {
        const string source = @"\left(\mathbb{R}<x\right)";
        using NotebookDocument notebook = NotebookDocument.Create();
        notebook.AddMarkdownCell(@"A supported $\frac{1}{2}$" + " and an unsupported $" + source + "$ formula.");

        string html = notebook.ToHtml(null, out OfficeConversionReport report);

        Assert.Contains("<mfrac>", html, StringComparison.Ordinal);
        Assert.Contains("<code>" + WebUtility.HtmlEncode(source) + "</code>", html, StringComparison.Ordinal);
        Assert.Single(report.Diagnostics, diagnostic => diagnostic.Code == "notebook-html-math-unsupported");
        Assert.Throws<OfficeConversionException>(() => notebook.ToHtml(new NotebookHtmlExportOptions { FailOnLoss = true }, out _));
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
