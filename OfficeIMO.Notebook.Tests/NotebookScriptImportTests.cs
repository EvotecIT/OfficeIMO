using OfficeIMO.Notebook.Import;

namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookScriptImportTests {
    [Fact]
    public void PolyglotProducerFixturePreservesOrderLanguagesAndKernelMetadataThroughJupyter() {
        var imported = NotebookImporter.Load(Fixture("Polyglot.dib"), NotebookImportFormat.PolyglotDib);
        using NotebookDocument notebook = imported.Value;
        Assert.False(imported.HasLoss);
        Assert.Equal("csharp", notebook.Language);
        Assert.Equal("csharp", notebook.Metadata["kernelInfo"].GetProperty("defaultKernelName").GetString());
        Assert.Collection(notebook.Cells,
            cell => AssertCell(cell, NotebookCellKind.Markdown, null, "* Markdown code"),
            cell => AssertCell(cell, NotebookCellKind.Code, "csharp", "// C# code"),
            cell => AssertCell(cell, NotebookCellKind.Code, "fsharp", "// F# code"),
            cell => AssertCell(cell, NotebookCellKind.Code, "pwsh", "# PowerShell code"),
            cell => AssertCell(cell, NotebookCellKind.Code, "javascript", "// JavaScript code"),
            cell => AssertCell(cell, NotebookCellKind.Code, "mermaid", "%% Mermaid code"));

        using NotebookDocument reopened = NotebookDocument.Load(notebook.ToBytes());
        Assert.Equal(notebook.Cells.Select(cell => (cell.Id, cell.Kind, cell.Language, cell.Source)),
            reopened.Cells.Select(cell => (cell.Id, cell.Kind, cell.Language, cell.Source)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PolyglotUsesDeclaredDefaultAndAliasWhileRetainingMagicSource(bool metadataFollowsSubmission) {
        const string metadataAndSource = """
            #!meta

            {"kernelInfo":{"defaultKernelName":"custom","items":[{"name":"custom","aliases":["cust"]}]}}

            #!cust

            call source
            #!import other.dib
            #!undeclared

            #!md

            # Notes

            #!cust

            second
            """;
        string source = metadataFollowsSubmission ? "#!csharp\n\nFirstSubmission();\n\n" + metadataAndSource : metadataAndSource;
        using NotebookDocument notebook = NotebookImporter.Load(Encoding.UTF8.GetBytes(source.Replace("\r\n", "\n").Replace("\n", "\r\n")),
            NotebookImportFormat.PolyglotDib).Value;

        Assert.Equal("custom", notebook.Language);
        var cells = notebook.Cells.AsEnumerable();
        if (metadataFollowsSubmission) {
            AssertCell(notebook.Cells[0], NotebookCellKind.Code, "csharp", "FirstSubmission();");
            cells = cells.Skip(1);
        }
        Assert.Collection(cells,
            cell => AssertCell(cell, NotebookCellKind.Code, "custom", "call source\n#!import other.dib\n#!undeclared"),
            cell => AssertCell(cell, NotebookCellKind.Markdown, null, "# Notes"),
            cell => AssertCell(cell, NotebookCellKind.Code, "custom", "second"));
    }

    [Fact]
    public void JupytextPreservesExplicitPairResultsAndStructuredMetadataThroughJupyter() {
        NotebookDocument imported;
        using (NotebookDocument companion = NotebookDocument.Load(Fixture("Paired.ipynb"))) {
            var result = NotebookImporter.Load(Fixture("Paired.py"), NotebookImportFormat.JupytextPercent,
                new NotebookImportOptions { PairedNotebook = companion });
            Assert.False(result.HasLoss);
            imported = result.Value;
            Assert.Equal(companion.Cells.Select(cell => (cell.Kind, cell.Source)),
                imported.Cells.Select(cell => (cell.Kind, cell.Source)));
        }

        using (imported)
        using (NotebookDocument reopened = NotebookDocument.Load(imported.ToBytes())) {
            Assert.Equal("Python 3", reopened.Metadata["kernelspec"].GetProperty("display_name").GetString());
            Assert.Equal("1.19.7", reopened.Metadata["jupytext"].GetProperty("text_representation").GetProperty("jupytext_version").GetString());
            Assert.Equal("10", reopened.Cells[1].Metadata["attributes"].GetProperty("n").GetString());
            Assert.True(reopened.Cells[1].Metadata["trusted"].GetBoolean());
            Assert.Equal(3, reopened.Cells[1].ExecutionCount);
            Assert.Collection(reopened.Cells[1].Outputs,
                output => Assert.Equal("Saved companion stream\n", output.Text),
                output => Assert.Equal(4, output.Data["application/json"].GetProperty("scalar").GetInt32()));
            Assert.StartsWith("%%time\n", reopened.Cells[2].Source);
            Assert.Equal("Saved second result\n", Assert.Single(reopened.Cells[2].Outputs).Text);
        }
    }

    [Fact]
    public void JupytextDiscoversNoCompanionAndRejectsStalePairAsAUnit() {
        string original = File.ReadAllText(Fixture("Paired.py"));
        using NotebookDocument sourceOnly = NotebookImporter.Load(Fixture("Paired.py"), NotebookImportFormat.JupytextPercent).Value;
        Assert.All(sourceOnly.Cells, cell => Assert.Empty(cell.Outputs));
        using NotebookDocument companion = NotebookDocument.Load(Fixture("Paired.ipynb"));
        string edited = original.Replace("second cell", "changed cell");
        var result = NotebookImporter.Load(Encoding.UTF8.GetBytes(edited), NotebookImportFormat.JupytextPercent,
            new NotebookImportOptions { PairedNotebook = companion });
        using NotebookDocument notebook = result.Value;

        Assert.True(result.HasLoss);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "notebook-companion-mismatch");
        Assert.Contains("changed cell", notebook.Cells[2].Source);
        Assert.All(notebook.Cells, cell => {
            Assert.Empty(cell.Outputs);
            Assert.Null(cell.ExecutionCount);
        });
    }

    [Fact]
    public void JupytextRequiresMatchingExplicitIdsForRepeatedSourceWithSavedResults() {
        using NotebookDocument companion = NotebookDocument.Load(Encoding.UTF8.GetBytes("""
            {"nbformat":4,"nbformat_minor":5,"metadata":{},"cells":[
              {"cell_type":"code","id":"first","source":"print(1)","metadata":{},"execution_count":1,"outputs":[]},
              {"cell_type":"code","id":"second","source":"print(1)","metadata":{},"execution_count":2,
               "outputs":[{"output_type":"stream","name":"stdout","text":"second result"}]}
            ]}
            """));
        var options = new NotebookImportOptions { PairedNotebook = companion };
        var ambiguous = NotebookImporter.Load(Encoding.UTF8.GetBytes("# %%\nprint(1)\n\n# %%\nprint(1)"),
            NotebookImportFormat.JupytextPercent, options);
        using NotebookDocument missingIds = ambiguous.Value;
        Assert.True(ambiguous.HasLoss);
        Assert.All(missingIds.Cells, cell => Assert.Null(cell.ExecutionCount));

        var identified = NotebookImporter.Load(Encoding.UTF8.GetBytes("# %% id=\"first\"\nprint(1)\n\n# %% id=\"second\"\nprint(1)"),
            NotebookImportFormat.JupytextPercent, options);
        using NotebookDocument withIds = identified.Value;
        Assert.False(identified.HasLoss);
        Assert.Equal("second", withIds.Cells[1].Id);
        Assert.Equal(2, withIds.Cells[1].ExecutionCount);
        Assert.Equal("second result", Assert.Single(withIds.Cells[1].Outputs).Text);
    }

    [Fact]
    public void JupytextExplicitIdsTakePrecedenceOverGeneratedIdsDuringPairing() {
        using NotebookDocument companion = NotebookDocument.Load(Encoding.UTF8.GetBytes("""
            {"nbformat":4,"nbformat_minor":5,"metadata":{},"cells":[
              {"cell_type":"code","id":"other","source":"print(1)","metadata":{},"execution_count":null,"outputs":[]},
              {"cell_type":"code","id":"cell-1","source":"print(2)","metadata":{},"execution_count":2,
               "outputs":[{"output_type":"stream","name":"stdout","text":"paired result"}]}
            ]}
            """));
        var result = NotebookImporter.Load(Encoding.UTF8.GetBytes("# %%\nprint(1)\n\n# %% id=\"cell-1\"\nprint(2)"),
            NotebookImportFormat.JupytextPercent, new NotebookImportOptions { PairedNotebook = companion });
        using NotebookDocument notebook = result.Value;
        Assert.False(result.HasLoss);
        Assert.Equal("cell-1", notebook.Cells[1].Id);
        Assert.NotEqual(notebook.Cells[1].Id, notebook.Cells[0].Id);
        Assert.Equal("paired result", Assert.Single(notebook.Cells[1].Outputs).Text);
    }

    [Theory]
    [InlineData("# %% language=\"bash\"\n# exit", "exit", "bash", null, false)]
    [InlineData("# %% language=\"R\" magic_args=\"-i values\"\n# sum(values)", "%%R -i values\nsum(values)", "R", null, true)]
    [InlineData("# %%\n# %%R\nprint(1)", "%%R\nprint(1)", "python", null, true)]
    [InlineData("# %% language=\"python\"\n# %%R\nprint(1)", "%%R\nprint(1)", "python", "python", true)]
    [InlineData("# %% language=\"R\"\n# sum(values)", "%%R\nsum(values)", "R", "python", true)]
    [InlineData("# %%\n# %%time\nprint(1)", "%%time\nprint(1)", "python", "r", false)]
    [InlineData("# %% {\"language\":null}\n1 + 1", "1 + 1", "python", "r", false)]
    [InlineData("# %% {\"language\":\" \"}\n# ordinary comment\n1 + 1", "# ordinary comment\n1 + 1", "python", "r", false)]
    public void JupytextPairsSavedResultsUsingEffectiveCellLanguages(string script, string source, string language,
        string? companionLanguage, bool matches) {
        using JsonDocument kernel = JsonDocument.Parse("""{"language":"python"}""");
        using NotebookDocument original = NotebookDocument.Create();
        original.SetMetadata("kernelspec", kernel.RootElement);
        NotebookCell saved = companionLanguage == null ? original.AddCodeCell(source) : original.AddCodeCell(source, companionLanguage);
        saved.ExecutionCount = 3;
        saved.AddOutput(NotebookOutput.CreateStream("stdout", "saved value"));
        using NotebookDocument companion = NotebookDocument.Load(original.ToBytes());

        var result = NotebookImporter.Load(Encoding.UTF8.GetBytes(script), NotebookImportFormat.JupytextPercent,
            new NotebookImportOptions { PairedNotebook = companion });
        using NotebookDocument notebook = result.Value;
        NotebookCell cell = Assert.Single(notebook.Cells);
        AssertCell(cell, NotebookCellKind.Code, language, source);
        if (matches) {
            Assert.False(result.HasLoss);
            Assert.Equal(3, cell.ExecutionCount);
            Assert.Equal("saved value", Assert.Single(cell.Outputs).Text);
        } else {
            Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "notebook-companion-mismatch");
            Assert.Null(cell.ExecutionCount);
            Assert.Empty(cell.Outputs);
        }
    }

    [Fact]
    public void JupytextSameLanguageAliasPreservesCodeComments() {
        const string code = "# This is a comment\nprint(1)";
        var result = NotebookImporter.Load(Encoding.UTF8.GetBytes("# %% language=\"py\"\n" + code),
            NotebookImportFormat.JupytextPercent);
        using NotebookDocument notebook = result.Value;
        Assert.False(result.HasLoss);
        AssertCell(Assert.Single(notebook.Cells), NotebookCellKind.Code, "py", code);
    }

    [Fact]
    public void JupytextRecognizesPinnedRJuliaAndTripleQuotedMarkdownExamples() {
        using NotebookDocument latex = NotebookImporter.Load(Fixture("Latex.py"), NotebookImportFormat.JupytextPercent).Value;
        Assert.Equal(3, latex.Cells.Count);
        Assert.All(latex.Cells, cell => Assert.Equal(NotebookCellKind.Markdown, cell.Kind));
        Assert.Contains(@"\begin{align}", latex.Cells[2].Source);
        Assert.Equal("r'''", latex.Cells[2].Metadata["cell_marker"].GetString());

        using NotebookDocument r = NotebookImporter.Load(Fixture("RNotebook.R"), NotebookImportFormat.JupytextPercent).Value;
        Assert.Equal("R", r.Language);
        Assert.Equal(4, r.Cells.Count);
        AssertCell(r.Cells[1], NotebookCellKind.Code, "r", "sum(1:10)");
        AssertCell(r.Cells[3], NotebookCellKind.Code, "r", string.Empty);

        using NotebookDocument julia = NotebookImporter.Load(Fixture("JuliaNotebook.jl"), NotebookImportFormat.JupytextPercent).Value;
        Assert.Equal(3, julia.Cells.Count);
        Assert.Contains("# %% not a boundary", julia.Cells[1].Source);
        AssertCell(julia.Cells[2], NotebookCellKind.Code, "julia", "sum(1:5)");
    }

    [Theory]
    [InlineData("a'")]
    [InlineData("a[1:end]'")]
    [InlineData("α̂'")]
    [InlineData("flag = true; value = !flag'")]
    [InlineData("for c in'\"'\n    println(c)\nend")]
    [InlineData(":(c isa'\"')")]
    [InlineData(":(T where'\"')")]
    [InlineData("for c in'''\n    println(c)\nend")]
    [InlineData("quote_char = '''")]
    [InlineData("fields = (; in = a); fields.in'")]
    [InlineData("value = in'")]
    public void JupytextJuliaApostrophesPreserveCellBoundariesAndCompanionResults(string code) {
        const string next = "println(22)";
        const string prefix = """"
            a = [1 2]
            α̂ = a
            label = '#'; description = """
            # %% remains inside the string
            """
            #=
            # %% remains inside the comment
            =#
            """";
        string source = prefix.Replace("\r\n", "\n") + "\n" + code;
        using NotebookDocument companion = NotebookDocument.Create();
        companion.AddCodeCell(source, "julia");
        NotebookCell saved = companion.AddCodeCell(next, "julia");
        saved.ExecutionCount = 2;
        saved.AddOutput(NotebookOutput.CreateStream("stdout", "22\n"));
        byte[] script = Encoding.UTF8.GetBytes("# %%\n" + source + "\n\n# %%\n" + next);

        var result = NotebookImporter.Load(script, NotebookImportFormat.JupytextPercent,
            new NotebookImportOptions { ScriptLanguage = "julia", PairedNotebook = companion });
        using NotebookDocument notebook = result.Value;
        Assert.False(result.HasLoss);
        Assert.Collection(notebook.Cells,
            cell => AssertCell(cell, NotebookCellKind.Code, "julia", source),
            cell => {
                AssertCell(cell, NotebookCellKind.Code, "julia", next);
                Assert.Equal(2, cell.ExecutionCount);
                Assert.Equal("22\n", Assert.Single(cell.Outputs).Text);
            });
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void JupytextOnlyKeepsPercentCommentsInsideClosedMarkdownFences(bool closeFence) {
        string fence = closeFence ? "# ```mermaid\n#     ```\n# %% inside chart\n#    ```" : "# ```mermaid";
        const string code = """"
            message = """
            ```
            # ```
            # %% still in the string
            """
            """";
        string source = "# %% [markdown]\n" + fence + "\n\n# %%\n" + code + "\n";

        using NotebookDocument notebook = NotebookImporter.Load(Encoding.UTF8.GetBytes(source), NotebookImportFormat.JupytextPercent).Value;
        Assert.Collection(notebook.Cells,
            cell => AssertCell(cell, NotebookCellKind.Markdown, null,
                closeFence ? "```mermaid\n    ```\n%% inside chart\n   ```" : "```mermaid"),
            cell => AssertCell(cell, NotebookCellKind.Code, "python", code.Replace("\r\n", "\n")));
    }

    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "Scripts", name);

    private static void AssertCell(NotebookCell cell, NotebookCellKind kind, string? language, string source) {
        Assert.Equal(kind, cell.Kind);
        Assert.Equal(language, cell.Language);
        Assert.Equal(source, cell.Source);
    }
}
