using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Word;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class WordObjectExportWorkflowTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-word-object-export-" + Guid.NewGuid().ToString("N"));
    public WordObjectExportWorkflowTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);

    [Theory]
    [InlineData(true, true)]
    [InlineData(false, false)]
    public void ObjectsBecomeAnEditableModernTableWithOptionalHeader(bool pipeline, bool header) {
        string path = Path.Combine(_root, "Services [current].docx");
        using var runspace = Open(path);
        using var shell = Shell(runspace);
        string command = pipeline ? "$rows | Export-OfficeWord" : "Export-OfficeWord -InputObject $rows";
        shell.AddScript("""
            $rows = @(
                [pscustomobject]@{ Name='Directory'; Count=2; Tags=@('Core','Identity') },
                [pscustomobject]@{ Name='Messaging'; Count=3; Tags=@('Mail') }
            )
            """ + "\n" + command + " -Path $target -Title 'Services' -PassThru" + (header ? "" : " -NoHeader"));
        var output = shell.Invoke();
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Equal(path, Assert.IsType<FileInfo>(Assert.Single(output).BaseObject).FullName);
        using WordDocument saved = WordDocument.Load(path);
        Assert.Empty(saved.ValidateDocument());
        Assert.Equal("Services", Assert.Single(saved.Paragraphs).Text);
        WordTable table = Assert.Single(saved.Tables);
        Assert.Equal(WordTableStyle.OfficeIMOModern, table.Style);
        Assert.Equal(header, table.RepeatHeaderRowAtTheTopOfEachPage);
        Assert.Equal(header ? 3 : 2, table.Rows.Count);
        int firstDataRow = header ? 1 : 0;
        if (header) {
            Assert.Equal(new[] { "Name", "Count", "Tags" }, table.Rows[0].Cells.Select(cell => cell.Paragraphs[0].Text));
        }
        Assert.Equal("Directory", table.Rows[firstDataRow].Cells[0].Paragraphs[0].Text);
        Assert.Equal("2", table.Rows[firstDataRow].Cells[1].Paragraphs[0].Text);
        Assert.Equal("Core, Identity", table.Rows[firstDataRow].Cells[2].Paragraphs[0].Text);
        Assert.Equal("Messaging", table.Rows[firstDataRow + 1].Cells[0].Paragraphs[0].Text);
    }

    [Fact]
    public void DataTableUsesTheSharedTabularProjection() {
        string path = Path.Combine(_root, "Data.docx");
        using var runspace = Open(path);
        using var shell = Shell(runspace);
        shell.AddScript("""
            $data = [System.Data.DataTable]::new()
            [void]$data.Columns.Add('Name', [string])
            [void]$data.Columns.Add('Total', [int])
            [void]$data.Rows.Add('North', 42)
            Export-OfficeWord -InputObject $data -Path $target
            """);
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using WordDocument saved = WordDocument.Load(path);
        WordTable table = Assert.Single(saved.Tables);
        Assert.Equal(new[] { "Name", "Total" }, table.Rows[0].Cells.Select(cell => cell.Paragraphs[0].Text));
        Assert.Equal(new[] { "North", "42" }, table.Rows[1].Cells.Select(cell => cell.Paragraphs[0].Text));
    }

    [Fact]
    public void SelectedColumnsUseDotNetFormatsAndExplicitOrientation() {
        string path = Path.Combine(_root, "Formatted.docx");
        using var runspace = Open(path);
        using var shell = Shell(runspace);
        shell.AddScript("""
            [pscustomobject]@{ Name='North'; Count=1234.6; Date=[datetime]'2026-10-30' } |
                Export-OfficeWord -Path $target -Property Date,Count -ColumnFormat @{ Date='yyyy-MM-dd'; Count='N0' } -Orientation Landscape
            """);
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using WordDocument saved = WordDocument.Load(path);
        Assert.Equal(OfficeIMO.OfficePageOrientation.Landscape, saved.Sections[0].PageOrientation);
        WordTable table = Assert.Single(saved.Tables);
        Assert.Equal(new[] { "Date", "Count" }, table.Rows[0].Cells.Select(cell => cell.Paragraphs[0].Text));
        Assert.Equal(new[] { "2026-10-30", "1,235" }, table.Rows[1].Cells.Select(cell => cell.Paragraphs[0].Text));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void WideTablesChooseLandscapeUnlessPortraitIsExplicit(bool automatic) {
        string path = Path.Combine(_root, "Wide.docx");
        using var runspace = Open(path);
        using var shell = Shell(runspace);
        shell.AddScript("""
            $row = [ordered]@{}
            foreach ($index in 1..12) { $row['LongColumnName' + $index] = 'Some content' }
            [pscustomobject]$row | Export-OfficeWord -Path $target
            """ + (automatic ? "" : " -Orientation Portrait"));
        Assert.Empty(shell.Invoke());
        Assert.False(shell.HadErrors, Errors(shell));
        using WordDocument saved = WordDocument.Load(path);
        Assert.Equal(automatic ? OfficeIMO.OfficePageOrientation.Landscape : OfficeIMO.OfficePageOrientation.Portrait,
            saved.Sections[0].PageOrientation);
        Assert.Equal(12, Assert.Single(saved.Tables).Rows[0].Cells.Count);
    }

    [Theory]
    [InlineData("-Property Missing")]
    [InlineData("-Property Name,Name")]
    [InlineData("-ColumnFormat @{ Missing='N0' }")]
    [InlineData("-ColumnFormat @{ Count='Q' }")]
    public void InvalidColumnSettingsLeaveTheDestinationUnchanged(string arguments) {
        string path = Path.Combine(_root, "Existing.docx");
        byte[] original = { 4, 2 };
        File.WriteAllBytes(path, original);
        using var runspace = Open(path);
        using var shell = Shell(runspace);
        shell.AddScript("[pscustomobject]@{Name='North'; Count=42} | Export-OfficeWord -Path $target " + arguments);
        Exception? failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors);
        Assert.Equal(original, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFileSystemEntries(_root));
    }

    [Fact]
    public void NoClobberPreservesExistingBytes() {
        string path = Path.Combine(_root, "Existing.docx");
        byte[] original = { 4, 2 };
        File.WriteAllBytes(path, original);
        using var runspace = Open(path);
        using var shell = Shell(runspace);
        shell.AddScript("[pscustomobject]@{Name='North'} | Export-OfficeWord -Path $target -NoClobber");
        Exception? failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors);
        Assert.Equal(original, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFileSystemEntries(_root));
    }

    [Theory]
    [InlineData(".docx", "-WhatIf", false)]
    [InlineData(".pdf", "", true)]
    public void WhatIfAndInvalidExtensionsNeverCreateTheDestination(string extension, string arguments, bool error) {
        string path = Path.Combine(_root, "absent", "Report" + extension);
        using var runspace = Open(path);
        using var shell = Shell(runspace);
        shell.AddScript("[pscustomobject]@{Name='North'} | Export-OfficeWord -Path $target " + arguments);
        Exception? failure = Record.Exception(() => shell.Invoke());
        Assert.Equal(error, failure != null || shell.HadErrors);
        Assert.Empty(Directory.GetFileSystemEntries(_root));
    }

    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error);
    private static PowerShell Shell(Runspace runspace) {
        var shell = PowerShell.Create();
        shell.Runspace = runspace;
        return shell;
    }
    private static Runspace Open(string path) {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Export-OfficeWord", typeof(ExportOfficeWordCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        runspace.SessionStateProxy.SetVariable("target", path);
        return runspace;
    }
}
