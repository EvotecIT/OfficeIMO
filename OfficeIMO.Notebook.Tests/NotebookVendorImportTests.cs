using OfficeIMO.Markdown;
using OfficeIMO.Notebook.Import;
using OfficeIMO.Notebook.Markdown;

namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookVendorImportTests {
    [Fact]
    public void DeepnoteSourceAndSnapshotImportIndependentlyWithStoredOutputs() {
        using NotebookDocument source = NotebookImporter.Load(Vendor("deepnote-hello.deepnote"), NotebookImportFormat.Deepnote).Value;
        OfficeConversionResult<NotebookDocument> imported = NotebookImporter.Load(Vendor("deepnote-hello.snapshot.deepnote"), NotebookImportFormat.Deepnote);
        using NotebookDocument snapshot = imported.Value;
        NotebookCell sourceCell = Assert.Single(source.Cells);
        NotebookCell snapshotCell = Assert.Single(snapshot.Cells);
        Assert.Equal(sourceCell.Id, snapshotCell.Id);
        Assert.Equal("print(\"Hello world!\")", snapshotCell.Source);
        Assert.Equal("python", snapshotCell.Language);
        Assert.Equal("1. Hello World - example", snapshot.Metadata["title"].GetString());
        Assert.Equal(1, snapshotCell.ExecutionCount);
        Assert.Empty(sourceCell.Outputs);
        NotebookOutput saved = Assert.Single(snapshotCell.Outputs);
        Assert.Equal(NotebookOutputKind.Stream, saved.Kind);
        Assert.Equal("Hello world!\n", saved.Text);
        Assert.Contains(imported.Report.Diagnostics, item => item.Code == "notebook-deepnote-project-settings");
    }

    [Fact]
    public void DeepnoteRichSnapshotRoundTripsSavedAlternativesAndKeepsVendorPromptsInert() {
        OfficeConversionResult<NotebookDocument> imported = NotebookImporter.Load(Vendor("deepnote-rich.snapshot.deepnote"), NotebookImportFormat.Deepnote);
        using NotebookDocument snapshot = imported.Value;
        Assert.Throws<OfficeConversionException>(() => snapshot.ToBytes());
        using NotebookDocument saved = NotebookDocument.Load(snapshot.ToBytes(new NotebookSaveOptions { AllowLoss = true }));
        Assert.Equal("# Sales performance", saved.Cells[0].Source);
        NotebookCell table = Assert.Single(saved.Cells, cell => cell.ExecutionCount == 11);
        NotebookOutput result = Assert.Single(table.Outputs);
        Assert.Equal(11, result.ExecutionCount);
        JsonElement vendorTable = result.Data["application/vnd.deepnote.dataframe.v3+json"];
        Assert.Equal(3, vendorTable.GetProperty("column_count").GetInt32());
        Assert.Contains("text/plain", result.Data.Keys);
        NotebookCell figure = Assert.Single(saved.Cells, cell => cell.ExecutionCount == 12);
        Assert.StartsWith("iVBOR", Assert.Single(figure.Outputs).Data["image/png"].GetString(), StringComparison.Ordinal);
        NotebookCell prompt = Assert.Single(saved.Cells, cell =>
            cell.Metadata.TryGetValue("officeimo_deepnote_block_type", out JsonElement type) && type.GetString() == "agent");
        Assert.Empty(prompt.Source);
        Assert.StartsWith("You are a senior revenue analyst", prompt.Metadata["officeimo_deepnote_block_source"].GetString(), StringComparison.Ordinal);
        Assert.Contains("Executive readout", Assert.Single(prompt.Outputs).Text, StringComparison.Ordinal);
        Assert.Contains(imported.Report.Diagnostics, item => item.Code == "notebook-deepnote-block-type");
    }

    [Fact]
    public void DeepnoteLegacyProjectsRequireAnExactNotebookIdAndSortNativeKeys() {
        byte[] bytes = Encoding.UTF8.GetBytes("""
            version: 1.0.0
            project:
              notebooks:
                - id: unused
                  blocks: []
                - id: target
                  blocks:
                    - {id: lower, type: sql, sortingKey: a0, content: select 1}
                    - {id: upper, type: markdown, sortingKey: A0, content: First}
                    - {id: tie, type: code, sortingKey: A0, content: second()}
            """);
        Assert.Throws<FormatException>(() => NotebookImporter.Load(bytes, NotebookImportFormat.Deepnote));
        Assert.Throws<FormatException>(() => NotebookImporter.Load(bytes, NotebookImportFormat.Deepnote,
            new NotebookImportOptions { DeepnoteNotebookId = "Target" }));
        using NotebookDocument notebook = NotebookImporter.Load(bytes, NotebookImportFormat.Deepnote,
            new NotebookImportOptions { DeepnoteNotebookId = "target" }).Value;
        Assert.Equal(new[] { "upper", "tie", "lower" }, notebook.Cells.Select(cell => cell.Id));
        Assert.Equal("sql", notebook.Cells[2].Language);
        Assert.Equal("select 1", notebook.Cells[2].Source);
    }

    [Fact]
    public void DeepnoteRejectsAliasesInThePortableProfile() {
        byte[] bytes = Encoding.UTF8.GetBytes("""
            version: 1.0.0
            project:
              notebooks:
                - id: notebook
                  blocks:
                    - {type: code, sortingKey: a0, content: &source print(1)}
                    - {type: code, sortingKey: a1, content: *source}
            """);
        Assert.Throws<FormatException>(() => NotebookImporter.Load(bytes, NotebookImportFormat.Deepnote));
    }

    [Fact]
    public void DeepnoteOnlyProjectsImageBlocksWithEmbeddedContent() {
        byte[] bytes = Encoding.UTF8.GetBytes("""
            version: 1.0.0
            project:
              notebooks:
                - id: notebook
                  blocks:
                    - type: image
                      sortingKey: a0
                      metadata:
                        deepnote_img_src: data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==
                    - type: image
                      sortingKey: a1
                      metadata:
                        deepnote_img_src: neighboring-file.png
            """);
        OfficeConversionResult<NotebookDocument> imported = NotebookImporter.Load(bytes, NotebookImportFormat.Deepnote);
        using NotebookDocument notebook = imported.Value;
        Assert.Single(notebook.Cells[0].Attachments);
        Assert.Empty(notebook.Cells[1].Attachments);
        using MarkdownDocument exported = notebook.ToMarkdown();
        Assert.Single(exported.Blocks, block => block is ImageBlock);
        Assert.Contains(imported.Report.Diagnostics, item => item.Code == "notebook-deepnote-image");
    }

    [Fact]
    public void ZeppelinNativeSavedSqlResultsProjectAsTablesAndKeepExactTsv() {
        using NotebookDocument notebook = NotebookImporter.Load(Vendor("zeppelin-spark.zpln"), NotebookImportFormat.Zeppelin).Value;
        Assert.Equal("2. Spark Basic Features", notebook.Metadata["title"].GetString());
        Assert.Equal(NotebookCellKind.Markdown, notebook.Cells[0].Kind);
        Assert.Equal("scala", notebook.Cells[1].Language);
        NotebookCell query = notebook.Cells[2];
        Assert.Equal("sql", query.Language);
        Assert.StartsWith("select age, count(1)", query.Source, StringComparison.Ordinal);
        NotebookOutput table = Assert.Single(query.Outputs);
        Assert.StartsWith("age\tvalue\n19\t4", table.Data["text/plain"].GetString(), StringComparison.Ordinal);
        using MarkdownDocument exported = notebook.ToMarkdown(new NotebookMarkdownExportOptions { IncludeCode = false });
        Assert.Contains(exported.Blocks, block => block is TableBlock);
        string markdown = exported.ToMarkdown();
        Assert.Contains("| age | value |", markdown, StringComparison.Ordinal);
        Assert.Contains("| 19 | 4 |", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void ZeppelinNativeRasterActivatesTheNotebookImageProjectionWithoutExecution() {
        using NotebookDocument notebook = NotebookImporter.Load(Vendor("zeppelin-image.zpln"), NotebookImportFormat.Zeppelin).Value;
        NotebookCell cell = Assert.Single(notebook.Cells, entry => entry.Id == "paragraph_1623916874799_812799753");
        Assert.Equal("python", cell.Language);
        Assert.Contains("%matplotlib inline", cell.Source, StringComparison.Ordinal);
        Assert.StartsWith("iVBOR", Assert.Single(cell.Outputs).Data["image/png"].GetString(), StringComparison.Ordinal);
        using MarkdownDocument exported = notebook.ToMarkdown(new NotebookMarkdownExportOptions { IncludeCode = false });
        Assert.Contains(exported.Blocks, block => block is ImageBlock);
    }

    [Fact]
    public void ZeppelinReportsUnsupportedSavedResultsAndCountsThemAgainstImportLimits() {
        byte[] bytes = Encoding.UTF8.GetBytes("""
            {"id":"synthetic","name":"Bounded result","paragraphs":[{
              "text":"%sh\nthis is saved source","status":"ERROR",
              "results":{"code":"ERROR","msg":[
                {"type":"ANGULAR","data":"<script>unexpected()</script>"},
                {"type":"TEXT","data":"previous run failed"}
              ]}
            }]}
            """);
        OfficeConversionResult<NotebookDocument> imported = NotebookImporter.Load(bytes, NotebookImportFormat.Zeppelin);
        using NotebookDocument notebook = imported.Value;
        NotebookCell cell = Assert.Single(notebook.Cells);
        Assert.Equal("this is saved source", cell.Source);
        NotebookOutput failure = Assert.Single(cell.Outputs);
        Assert.Equal(NotebookOutputKind.Error, failure.Kind);
        Assert.Equal("previous run failed", failure.ErrorValue);
        Assert.Contains(imported.Report.Diagnostics, item => item.Code == "notebook-zeppelin-result" &&
            item.Location == "paragraphs[0].results.msg[0]");
        Assert.Throws<InvalidDataException>(() => NotebookImporter.Load(bytes, NotebookImportFormat.Zeppelin,
            new NotebookImportOptions { Limits = new NotebookLoadOptions { MaxOutputsPerCell = 1 } }));
        Assert.Throws<FormatException>(() => NotebookImporter.Load(Encoding.UTF8.GetBytes("{\"paragraphs\":[]}"),
            NotebookImportFormat.Zeppelin));
    }

    private static string Vendor(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "Vendors", name);
}
