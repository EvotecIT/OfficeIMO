namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookProducerCompatibilityTests {
    [Theory]
    [InlineData("jupyterlab", 5, 2, "kernelspec", 1, "tags", "print('multiple')", null, "multiple\n")]
    [InlineData("colab", 0, 9, "colab", 1, "colab_type", "# Using Google Colab with GitHub", null, null)]
    [InlineData("vscode", 5, 2, "language_info", 1, "vscode", "Write-Output", "powershell", "Saved from VS Code\n")]
    [InlineData("kaggle", 4, 3, "kaggle", 1, "_kg_hide-input", "import os", null, "/kaggle/input\n/kaggle/working\n")]
    [InlineData("polyglot", 2, 6, "kernelspec", 4, "dotnet_interactive", "LaTeXString", "fsharp", null)]
    [InlineData("deepnote", 0, 1, "deepnote_notebook_id", 0, "deepnote_sorting_key", "print(\"Hello world!\")", null, "Hello world!\n")]
    [InlineData("databricks", 0, 1, "application/vnd.databricks.v1+notebook", 0, "application/vnd.databricks.v1+cell", "teste", null, null)]
    [InlineData("marimo", 5, 2, "marimo", 1, null, "print(", null, "Saved marimo witness 42\n")]
    public void OpensProducerSourceAndKeepsMetadataAndOutputsAfterEditing(string producer, int minor, int cellCount,
        string notebookMetadata, int cellIndex, string? cellMetadata, string sourcePrefix, string? language, string? stdout) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "Producers", producer + ".ipynb");
        using JsonDocument original = JsonDocument.Parse(File.ReadAllBytes(path));
        using NotebookDocument notebook = NotebookDocument.Load(path);

        Assert.Equal(minor, notebook.NbFormatMinor);
        Assert.Equal(cellCount, notebook.Cells.Count);
        Assert.Contains(notebookMetadata, notebook.Metadata.Keys);
        NotebookCell witness = notebook.Cells[cellIndex];
        Assert.StartsWith(sourcePrefix, witness.Source, StringComparison.Ordinal);
        if (cellMetadata != null) Assert.Contains(cellMetadata, witness.Metadata.Keys);
        if (language != null) Assert.Equal(language, witness.Language);
        if (stdout != null) {
            Assert.Contains(witness.Outputs, output => output.Kind == NotebookOutputKind.Stream &&
                output.Name == "stdout" && output.Text == stdout);
        }
        if (producer == "marimo") {
            Assert.Contains("text/markdown", Assert.Single(notebook.Cells[0].Outputs).Data.Keys);
        }

        string editedSource = witness.Source + "\n";
        witness.Source = editedSource;
        using JsonDocument edited = JsonDocument.Parse(notebook.ToBytes());
        Assert.Equal(5, edited.RootElement.GetProperty("nbformat_minor").GetInt32());
        Assert.True(JsonElement.DeepEquals(original.RootElement.GetProperty("metadata"), edited.RootElement.GetProperty("metadata")));

        JsonElement before = original.RootElement.GetProperty("cells");
        JsonElement after = edited.RootElement.GetProperty("cells");
        Assert.Equal(editedSource, after[cellIndex].GetProperty("source").GetString());
        if (producer == "deepnote") {
            string? sourceGroup = before[0].GetProperty("block_group").GetString();
            Assert.Equal(sourceGroup, after[0].GetProperty("metadata").GetProperty("deepnote_block_group").GetString());
            Assert.False(after[0].TryGetProperty("block_group", out _));
        }
        for (int index = 0; index < before.GetArrayLength(); index++) {
            Assert.True(JsonElement.DeepEquals(before[index].GetProperty("metadata"), after[index].GetProperty("metadata")));
            if (before[index].TryGetProperty("outputs", out JsonElement outputs)) {
                Assert.True(JsonElement.DeepEquals(outputs, after[index].GetProperty("outputs")));
            }
        }

        var ids = after.EnumerateArray().Select(cell => cell.GetProperty("id").GetString()).ToArray();
        Assert.Equal(cellCount, ids.Distinct(StringComparer.Ordinal).Count());
        Assert.All(ids, id => Assert.Matches("^[A-Za-z0-9_-]{1,64}$", id!));
    }
}
