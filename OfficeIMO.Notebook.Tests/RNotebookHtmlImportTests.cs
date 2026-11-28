using OfficeIMO.Notebook.Import;
using OfficeIMO.Notebook.Html;

namespace OfficeIMO.Notebook.Tests;

public sealed class RNotebookHtmlImportTests {
    [Fact]
    public void RestoresStaticResultsAndRetainsUnmatchedOutputInASeparateCell() {
        string input = File.ReadAllText(NotebookMarkupImportTests.Fixture("r-saved-results.nb.html"))
            .Replace("<pre><code>[1] 2 &amp; 6</code></pre>", "<pre><code>[1] 2 &amp; 6\n</code></pre>\n" +
                "<!-- rnb-output-end -->\n<!-- rnb-output-begin -->\n<pre><code>  \n</code></pre>\n" +
                "<!-- rnb-output-end -->\n<!-- rnb-output-begin -->\n<span>saved</span> <span>text</span>");
        OfficeConversionResult<NotebookDocument> converted = NotebookImporter.Load(
            Encoding.UTF8.GetBytes(input), NotebookImportFormat.RNotebookHtml);
        using NotebookDocument notebook = converted.Value;
        NotebookCell[] code = notebook.Cells.Where(cell => cell.Kind == NotebookCellKind.Code).ToArray();
        Assert.Equal(4, code.Length);
        Assert.Collection(code[0].Outputs,
            output => Assert.Equal("[1] 2 & 6\n", output.Text),
            output => Assert.Equal("  \n", output.Text),
            output => Assert.Equal("saved text", output.Text),
            output => Assert.Contains("<table", output.Data["text/html"].GetString(), StringComparison.Ordinal),
            output => Assert.True(output.Data.ContainsKey("image/png")));
        Assert.Empty(code[1].Outputs);
        Assert.Empty(code[2].Outputs);
        Assert.Equal(string.Empty, code[3].Source);
        Assert.Contains("Saved text without identifying source", Assert.Single(code[3].Outputs).Text, StringComparison.Ordinal);
        Assert.Contains(converted.Report.Diagnostics, finding => finding.Code == "notebook-import-rnb-unmatched-output"
            && finding.Location == "rnb.chunks[1]");
        Assert.Contains(converted.Report.Diagnostics, finding => finding.Code == "notebook-import-rnb-image");
        Assert.Contains(converted.Report.Diagnostics, finding => finding.Code == "notebook-import-rnb-output");
        string html = notebook.ToHtml();
        Assert.Contains("data:image/png;base64,", html, StringComparison.Ordinal);
        Assert.Contains("<table", html, StringComparison.Ordinal);
        Assert.DoesNotContain("NotebookWidgetMarker", html, StringComparison.Ordinal);
        Assert.DoesNotContain("example.invalid", html, StringComparison.Ordinal);
    }

    [Fact]
    public void ReportsSavedTableResourcesAndUnsupportedElementsRemovedBySanitization() {
        const string nestedOutput = """
            <td>42<img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==">
            <script>window.privateSavedTableWidget = true;</script></td>
            """;
        string input = File.ReadAllText(NotebookMarkupImportTests.Fixture("r-saved-results.nb.html"))
            .Replace("<td>42</td>", nestedOutput);
        OfficeConversionResult<NotebookDocument> converted = NotebookImporter.Load(
            Encoding.UTF8.GetBytes(input), NotebookImportFormat.RNotebookHtml);
        using NotebookDocument notebook = converted.Value;
        NotebookOutput table = Assert.Single(notebook.Cells.SelectMany(cell => cell.Outputs),
            output => output.Data.ContainsKey("text/html"));
        string html = table.Data["text/html"].GetString()!;
        Assert.Contains("42", html, StringComparison.Ordinal);
        Assert.DoesNotContain("base64,", html, StringComparison.Ordinal);
        Assert.DoesNotContain("privateSavedTableWidget", html, StringComparison.Ordinal);
        Assert.Contains(converted.Report.Diagnostics, finding => finding.Code == "notebook-import-rnb-table-resource"
            && finding.Location == "rnb.chunks[0]" && finding.LossKind == OfficeConversionLossKind.Omission);
        Assert.Contains(converted.Report.Diagnostics, finding => finding.Code == "notebook-import-rnb-table-content"
            && finding.Location == "rnb.chunks[0]" && finding.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void RejectsOrdinaryHtmlAndBoundsRecoveredOutputs() {
        byte[] plainHtml = Encoding.UTF8.GetBytes("<!doctype html><html><body><p>Ordinary HTML.</p></body></html>");
        Assert.Throws<InvalidDataException>(() => NotebookImporter.Load(plainHtml, NotebookImportFormat.RNotebookHtml));
        var options = new NotebookImportOptions { Limits = new NotebookLoadOptions { MaxTotalOutputs = 1 } };
        var error = Assert.Throws<InvalidDataException>(() => NotebookImporter.Load(
            NotebookMarkupImportTests.Fixture("r-saved-results.nb.html"), NotebookImportFormat.RNotebookHtml, options));
        Assert.Contains("saved outputs", error.Message, StringComparison.OrdinalIgnoreCase);
    }
}
