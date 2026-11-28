namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookCellLanguageTests {
    [Theory]
    [InlineData("  ", """{"language":"python"}""", "python")]
    [InlineData(null, """{"language":"\t","polyglot_notebook":{"kernelName":"python"}}""", "python")]
    [InlineData(null, """{"polyglot_notebook":{"kernelName":""},"dotnet_interactive":{"language":"python"}}""", "python")]
    [InlineData(null, """{"vscode":{"languageId":"  "}}""", null)]
    [InlineData("python", """{"language":"powershell","vscode":{"languageId":"pwsh"}}""", "python")]
    public void BlankHintsFallThroughWithoutRewritingTheSourceLanguage(string? sourceLanguage, string metadata, string? expected) {
        string source = $$"""
            {
              "nbformat":4,"nbformat_minor":5,"metadata":{},
              "cells":[{"cell_type":"code","id":"example","language":{{JsonSerializer.Serialize(sourceLanguage)}},
                        "metadata":{{metadata}},"source":"'example'","outputs":[],"execution_count":null}]
            }
            """;
        using NotebookDocument notebook = NotebookDocument.Load(Encoding.UTF8.GetBytes(source));
        Assert.Equal(expected, Assert.Single(notebook.Cells).Language);

        byte[] saved = notebook.ToBytes();
        using JsonDocument savedJson = JsonDocument.Parse(saved);
        using JsonDocument originalMetadata = JsonDocument.Parse(metadata);
        JsonElement savedCell = savedJson.RootElement.GetProperty("cells")[0];
        Assert.Equal(sourceLanguage, savedCell.GetProperty("language").GetString());
        Assert.True(JsonElement.DeepEquals(originalMetadata.RootElement, savedCell.GetProperty("metadata")));
        using NotebookDocument reopened = NotebookDocument.Load(saved);
        Assert.Equal(expected, Assert.Single(reopened.Cells).Language);
    }

    [Fact]
    public void SettingLanguageMetadataUpdatesAnUntypedCellBeforeAndAfterSaving() {
        using NotebookDocument notebook = NotebookDocument.Create();
        notebook.SetMetadata("language_info", JsonSerializer.SerializeToElement(new { name = "python" }));
        NotebookCell cell = notebook.AddCodeCell("'hello from PowerShell'");
        Assert.Equal("python", cell.Language ?? notebook.Language);

        cell.SetMetadata("language", JsonSerializer.SerializeToElement("powershell"));
        Assert.Equal("powershell", cell.Language);

        using NotebookDocument reopened = NotebookDocument.Load(notebook.ToBytes());
        Assert.Equal("python", reopened.Language);
        NotebookCell reopenedCell = Assert.Single(reopened.Cells);
        Assert.Equal("powershell", reopenedCell.Language);

        reopenedCell.SetMetadata("polyglot_notebook", JsonSerializer.SerializeToElement(new { kernelName = "python" }));
        Assert.Equal("powershell", reopenedCell.Language);
        reopenedCell.SetMetadata("language", JsonSerializer.SerializeToElement(" \t"));
        Assert.Equal("python", reopenedCell.Language);
    }
}
