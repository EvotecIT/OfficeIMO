using OfficeIMO.Notebook.Import;

namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookMarkupImportTests {
    [Fact]
    public void QuartoLeavesExampleFencesInProseAndSeparatesOnlyExecutableCells() {
        OfficeConversionResult<NotebookDocument> converted = NotebookImporter.Load(Fixture("quarto-analysis.qmd"), NotebookImportFormat.Quarto);
        using NotebookDocument notebook = converted.Value;
        Assert.Equal("Offline Quarto lab", notebook.Metadata["title"].GetString());
        Assert.Equal("python", notebook.Language);
        Assert.Equal(new[] { NotebookCellKind.Markdown, NotebookCellKind.Code, NotebookCellKind.Markdown },
            notebook.Cells.Select(cell => cell.Kind));
        Assert.Contains("this_is_documentation()", notebook.Cells[0].Source, StringComparison.Ordinal);
        NotebookCell code = notebook.Cells[1];
        Assert.Equal("python", code.Language);
        Assert.Contains("measurements", code.Source, StringComparison.Ordinal);
        Assert.DoesNotContain("#|", code.Source, StringComparison.Ordinal);
        Assert.Contains("label: fig-run", code.Metadata["notebook_import_options"].GetString(), StringComparison.Ordinal);
        Assert.Empty(code.Outputs);
        Assert.Contains(converted.Report.Diagnostics, finding => finding.Code == "notebook-import-markup-directive");
    }

    [Fact]
    public void QuartoIndentedFenceRemovesOnlyFenceIndentFromCodeAndOptions() {
        const string source = "  ```{python}\r\n  #| label: fenced\r\n  if True:\r\n    print(1)\r\n"
            + "\tprint(2)\r\n print(3)\r\n ```\r\n";
        var result = NotebookImporter.Load(Encoding.UTF8.GetBytes(source), NotebookImportFormat.Quarto);

        using NotebookDocument notebook = result.Value;
        NotebookCell cell = Assert.Single(notebook.Cells);
        Assert.Equal(NotebookCellKind.Code, cell.Kind);
        Assert.Equal("if True:\r\n  print(1)\r\n  print(2)\r\nprint(3)\r\n", cell.Source);
        Assert.Equal("#| label: fenced\r\n", cell.Metadata["notebook_import_options"].GetString());
        Assert.Contains(result.Report.Diagnostics, finding => finding.Code == "notebook-import-chunk-options");
    }

    [Fact]
    public void QuartoRetainsObservableAndKnitrOptionsAsMetadata() {
        const string source = """
            ```{ojs}
            //| echo: false
            result = 3
            ```
            ```{r, echo = FALSE}
            summary(results)
            ```
            ```{r}
            plot(results)
            ```
            """;
        var imported = NotebookImporter.Load(Encoding.UTF8.GetBytes(source), NotebookImportFormat.Quarto);
        using NotebookDocument notebook = imported.Value;

        Assert.Collection(notebook.Cells,
            cell => {
                Assert.Equal(("ojs", "result = 3\n"), (cell.Language, cell.Source));
                Assert.Equal("//| echo: false\n", cell.Metadata["notebook_import_options"].GetString());
            },
            cell => Assert.Equal("{r, echo = FALSE}", cell.Metadata["notebook_import_header"].GetString()),
            cell => Assert.Equal("{r}", cell.Metadata["notebook_import_header"].GetString()));
        Assert.Equal(new[] { "lines[1]", "lines[5]" }, imported.Report.Diagnostics
            .Where(finding => finding.Code == "notebook-import-chunk-options").Select(finding => finding.Location));
        Assert.True(imported.HasLoss);
    }

    [Fact]
    public void MySTKeepsExecutableFencesInsideHtmlCommentsAsProse() {
        const string comment = """
             <!-- Disabled submission
            ```{code-cell} python
            print(1)
            ```
            -->
            """;
        const string source = comment + "\n\n```{code-cell} python\nprint(2)\n```";
        var imported = NotebookImporter.Load(Encoding.UTF8.GetBytes(source), NotebookImportFormat.MySTMarkdown);
        using NotebookDocument notebook = imported.Value;

        Assert.False(imported.HasLoss);
        Assert.Collection(notebook.Cells,
            cell => Assert.Equal((NotebookCellKind.Markdown, comment + "\n\n"), (cell.Kind, cell.Source)),
            cell => Assert.Equal((NotebookCellKind.Code, "print(2)\n"), (cell.Kind, cell.Source)));
    }

    [Fact]
    public void MySTRestoresInheritedAndExplicitLanguagesAndKeepsDirectiveOptions() {
        OfficeConversionResult<NotebookDocument> converted = NotebookImporter.Load(Fixture("myst-analysis.md"), NotebookImportFormat.MySTMarkdown);
        using NotebookDocument notebook = converted.Value;
        NotebookCell[] code = notebook.Cells.Where(cell => cell.Kind == NotebookCellKind.Code).ToArray();
        Assert.Equal(new[] { "python", "julia" }, code.Select(cell => cell.Language));
        Assert.Contains(":label: cells-summary", code[0].Metadata["notebook_import_options"].GetString(), StringComparison.Ordinal);
        Assert.Contains("skip-execution", code[1].Metadata["notebook_import_options"].GetString(), StringComparison.Ordinal);
        Assert.StartsWith("sum", code[1].Source, StringComparison.Ordinal);
        Assert.All(code, cell => Assert.Empty(cell.Outputs));
        Assert.Equal(2, converted.Report.Diagnostics.Count(finding => finding.Code == "notebook-import-markup-directive"));
        Assert.Contains("{eval}`sum(values)`", notebook.Cells.Last().Source, StringComparison.Ordinal);
    }

    [Fact]
    public void RMarkdownSourceContainsChunkProgramsWithoutSavedExecutionResults() {
        OfficeConversionResult<NotebookDocument> converted = NotebookImporter.Load(Fixture("r-analysis.Rmd"), NotebookImportFormat.RMarkdown);
        using NotebookDocument notebook = converted.Value;
        NotebookCell[] code = notebook.Cells.Where(cell => cell.Kind == NotebookCellKind.Code).ToArray();
        Assert.Equal("r", notebook.Language);
        Assert.Equal(3, code.Length);
        Assert.Equal("{r measurements, echo=TRUE}", code[0].Metadata["notebook_import_header"].GetString());
        Assert.Contains("summary(measurements)", code[0].Source, StringComparison.Ordinal);
        Assert.All(code, cell => Assert.Empty(cell.Outputs));
        Assert.Contains(converted.Report.Diagnostics, finding => finding.Code == "notebook-import-markup-directive");
    }

    internal static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "Markup", name);
}
