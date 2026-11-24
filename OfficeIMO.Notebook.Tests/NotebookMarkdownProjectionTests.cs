using OfficeIMO.Markdown;
using OfficeIMO.Notebook.Markdown;
using System.Text;
using Xunit;

namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookMarkdownProjectionTests {
    [Fact]
    public async Task TextExportPreservesAuthoredCellSpellingAndKeepsSavedOutputs() {
        const string source = "Heading  \r\n=======\r\n\r\n*literal dialect*  \r\n<div>retained Markdown source</div>";
        using NotebookDocument notebook = NotebookDocument.Create();
        notebook.AddMarkdownCell(source);
        notebook.AddCodeCell("generate()").AddOutput(NotebookOutput.CreateStream("stdout", "saved-result"));
        using var destination = new MemoryStream();

        OfficeConversionReport report = await notebook.SaveAsMarkdownAsync(destination, new NotebookMarkdownExportOptions {
            PreserveMarkdownSource = true
        });
        string markdown = Encoding.UTF8.GetString(destination.ToArray());

        Assert.StartsWith(source, markdown, StringComparison.Ordinal);
        Assert.Contains("generate()", markdown, StringComparison.Ordinal);
        Assert.Contains("saved-result", markdown, StringComparison.Ordinal);
        Assert.True(markdown.IndexOf("generate()", StringComparison.Ordinal) < markdown.IndexOf("saved-result", StringComparison.Ordinal));
        Assert.DoesNotContain("=======", notebook.ToMarkdownText(), StringComparison.Ordinal);
        Assert.False(report.HasLoss);
    }

    [Fact]
    public void ProjectsCellLocalAttachmentsAndHtmlTablesWithoutMaterializingExternalImages() {
        const string tinyPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==";
        using NotebookDocument notebook = NotebookDocument.Load(Encoding.UTF8.GetBytes("""
            {"nbformat":4,"nbformat_minor":5,"metadata":{},"cells":[
              {"id":"prose","cell_type":"markdown","metadata":{},
               "source":"# Lab\n\n![Local plot][plot]\n\n[plot]: attachment:one.png\n\n![External](https://example.invalid/figure.png)",
               "attachments":{"one.png":{"image/png":"PAYLOAD"}}},
              {"id":"other","cell_type":"markdown","metadata":{},"source":"![Unavailable](attachment:one.png)"},
              {"id":"report","cell_type":"code","metadata":{},"execution_count":1,"source":"private_code()","outputs":[
                 {"output_type":"display_data","metadata":{},"data":{
                   "text/html":"<table><tr><th>Sample</th><th>Weight</th></tr><tr><td>Alpha</td><td>42</td></tr></table>",
                   "text/plain":"TABLE_FALLBACK"}}
              ]}]}
            """.Replace("PAYLOAD", tinyPng)));

        using MarkdownDocument result = notebook.ToMarkdown(new NotebookMarkdownExportOptions { IncludeCode = false }, out OfficeConversionReport report);
        string markdown = result.ToMarkdown();
        Assert.Contains("data:image/png;base64,", markdown, StringComparison.Ordinal);
        Assert.Contains("Local plot", markdown, StringComparison.Ordinal);
        Assert.Contains("Unavailable", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("example.invalid", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("private_code", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("TABLE_FALLBACK", markdown, StringComparison.Ordinal);
        Assert.Single(result.DescendantsOfType<TableBlock>());
        Assert.Contains("Alpha", markdown, StringComparison.Ordinal);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "notebook-image-reference" && diagnostic.Location == "cells[0]");
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "notebook-attachment-unavailable" && diagnostic.Location == "cells[1]");

        string sourceText = notebook.ToMarkdownText(new NotebookMarkdownExportOptions { PreserveMarkdownSource = true },
            out OfficeConversionReport sourceReport);
        Assert.Contains("[plot]: attachment:one.png", sourceText, StringComparison.Ordinal);
        Assert.DoesNotContain("data:image/png", sourceText, StringComparison.Ordinal);
        Assert.Contains(sourceReport.Diagnostics, diagnostic =>
            diagnostic.Code == "notebook-source-attachments-omitted" && diagnostic.Location == "cells[0]");
    }
}
