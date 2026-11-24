namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookCellLanguageTests {
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
        Assert.Equal("powershell", Assert.Single(reopened.Cells).Language);
    }
}
