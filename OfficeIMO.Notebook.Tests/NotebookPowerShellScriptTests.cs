using OfficeIMO.Notebook.Import;

namespace OfficeIMO.Notebook.Tests;

public sealed class NotebookPowerShellScriptTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PlainPowerShellScriptPreservesSourceAndOrdinaryMarkerComments(bool utf16) {
        const string source = "#requires -Version 5.1\r\n# %% ordinary commentary\r\nWrite-Output 'zażółć'\r\n";
        Encoding encoding = utf16 ? Encoding.Unicode : new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);
        byte[] input = encoding.GetPreamble().Concat(encoding.GetBytes(source)).ToArray();

        var imported = NotebookImporter.Load(input, NotebookImportFormat.PowerShellScript);
        using NotebookDocument notebook = imported.Value;
        NotebookCell cell = Assert.Single(notebook.Cells);

        Assert.False(imported.HasLoss);
        Assert.Equal(NotebookCellKind.Code, cell.Kind);
        Assert.Equal("powershell", notebook.Language);
        Assert.Equal("powershell", cell.Language);
        Assert.Equal(source, cell.Source);
        Assert.Empty(cell.Outputs);
    }

    [Fact]
    public void ExplicitCommentBoundariesPreserveCellKindsAndIdsWithWindowsLineEndings() {
        const string script = "# OfficeIMO PowerShell Notebook v1 7\r\n" +
            "# %% [7] markdown intro\r\n# ## Observations\r\n# \r\n# Samples from the lab.\r\n" +
            "# %% [7] code code-id\r\nGet-Item .\r\n";

        using NotebookDocument notebook = NotebookImporter.Load(Encoding.UTF8.GetBytes(script),
            NotebookImportFormat.PowerShellScript).Value;
        Assert.Collection(notebook.Cells,
            cell => {
                Assert.Equal(("intro", NotebookCellKind.Markdown, "## Observations\r\n\r\nSamples from the lab."),
                    (cell.Id, cell.Kind, cell.Source));
            },
            cell => Assert.Equal(("code-id", NotebookCellKind.Code, "Get-Item ."), (cell.Id, cell.Kind, cell.Source)));
    }

    [Fact]
    public async Task ScriptFileEncodingAndRoundTripPreserveCodeAndMarkdownAcrossBoundaryCollisions() {
        const string code = "$value = @'\r\n# %% [0] code marker\r\n中文\r\n'@\r";
        const string markdown = "## 文档\r\n<#\r#>\n# %% [1] markdown not-a-cell\n";
        using NotebookDocument original = NotebookDocument.Create();
        NotebookCell entry = original.AddCodeCell(code, "pwsh");
        NotebookCell note = original.AddMarkdownCell(markdown);
        NotebookCell blank = original.AddCodeCell(string.Empty, "powershell");
        using var output = new MemoryStream();

        OfficeConversionReport report = await NotebookPowerShellScript.SaveAsync(original, output, failOnLoss: true);
        Assert.False(report.HasLoss);
        Assert.True(output.CanRead);
        Assert.Equal(0L, output.Position);
        Assert.Equal(new byte[] { 0xEF, 0xBB, 0xBF }, output.ToArray().Take(3));

        using NotebookDocument reopened = (await NotebookImporter.LoadAsync(output, NotebookImportFormat.PowerShellScript)).Value;
        Assert.Collection(reopened.Cells,
            cell => Assert.Equal((entry.Id, NotebookCellKind.Code, code), (cell.Id, cell.Kind, cell.Source)),
            cell => Assert.Equal((note.Id, NotebookCellKind.Markdown, markdown), (cell.Id, cell.Kind, cell.Source)),
            cell => Assert.Equal((blank.Id, NotebookCellKind.Code, string.Empty), (cell.Id, cell.Kind, cell.Source)));
    }

    [Fact]
    public async Task ScriptExportReportsLossAndHonorsStrictModeBeforeReplacingAnExistingFile() {
        using NotebookDocument notebook = NotebookDocument.Create();
        NotebookCell code = notebook.AddCodeCell("Get-Item .", "PowerShell");
        code.AddOutput(NotebookOutput.CreateStream("stdout", "saved stdout"));
        code.ExecutionCount = 1;
        notebook.AddCodeCell("print('mixed kernel')", "python");
        notebook.AddRawCell("raw presentation data");
        using JsonDocument attachment = JsonDocument.Parse("\"data\"");
        notebook.AddMarkdownCell("![Graph](attachment:graph.png)").SetAttachment("graph.png",
            new Dictionary<string, JsonElement> { ["image/png"] = attachment.RootElement });
        notebook.SetMetadata("description", attachment.RootElement);
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-script-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "lab.ps1");
            File.WriteAllText(path, "existing script");
            OfficeConversionException error = await Assert.ThrowsAsync<OfficeConversionException>(
                () => NotebookPowerShellScript.SaveAsync(notebook, path, failOnLoss: true));
            Assert.Equal("existing script", File.ReadAllText(path));
            Assert.Contains(error.Report.Diagnostics, item => item.Code == "notebook-powershell-outputs-omitted" &&
                item.Location == "cells[0].outputs");
            Assert.Contains(error.Report.Diagnostics, item => item.Code == "notebook-powershell-execution-count-omitted");
            Assert.Contains(error.Report.Diagnostics, item => item.Code == "notebook-powershell-language-omitted");
            Assert.Contains(error.Report.Diagnostics, item => item.Code == "notebook-powershell-cell-omitted");
            Assert.Contains(error.Report.Diagnostics, item => item.Code == "notebook-powershell-attachments-omitted");
            Assert.Contains(error.Report.Diagnostics, item => item.Code == "notebook-powershell-notebook-metadata-omitted");

            OfficeConversionReport published = NotebookPowerShellScript.Save(notebook, path);
            Assert.Equal(path, published.OutputPath);
            using NotebookDocument reopened = NotebookImporter.Load(path, NotebookImportFormat.PowerShellScript).Value;
            Assert.Collection(reopened.Cells,
                cell => {
                    Assert.Equal("Get-Item .", cell.Source);
                    Assert.Empty(cell.Outputs);
                    Assert.Null(cell.ExecutionCount);
                },
                cell => Assert.Equal("![Graph](attachment:graph.png)", cell.Source));
            Assert.Single(Directory.GetFiles(directory));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData("# OfficeIMO PowerShell Notebook v2 7\n", "unsupported version")]
    [InlineData("# OfficeIMO PowerShell Notebook v1 7\n# %% [7] markdown note\nWrite-Output 'text'\n", "must begin")]
    public void MalformedMarkedScriptsFailInsteadOfSilentlyLosingContent(string source, string expectedMessage) {
        var error = Assert.Throws<FormatException>(() => NotebookImporter.Load(Encoding.UTF8.GetBytes(source),
            NotebookImportFormat.PowerShellScript));
        Assert.Contains(expectedMessage, error.Message, StringComparison.Ordinal);
    }
}
