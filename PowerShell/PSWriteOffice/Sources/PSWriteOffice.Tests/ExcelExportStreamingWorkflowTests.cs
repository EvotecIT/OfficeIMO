using System.Data;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Excel;
using PSWriteOffice.Cmdlets.Excel;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ExcelExportStreamingWorkflowTests {
    [Fact]
    public void PipelineCopiesEachCurrentRowAndKeepsLaterColumns() {
        InDirectory(path => {
            using var runspace = CreateRunspace(path);
            Invoke(runspace, """
                $row = [pscustomobject]@{ Name = 'first'; Amount = 7 }
                & {
                    $row
                    $row.Name = 'second'
                    $row | Add-Member -NotePropertyName Later -NotePropertyValue 19
                    $row
                } | Export-OfficeExcel -Path $destination -Sheet Orders -Address B3 -Title 'Daily orders' -DocumentProperty @{ Author = 'Finance' }
                """);
            using var saved = ExcelDocument.Load(path);
            var sheet = saved["Orders"];
            Assert.Equal("Daily orders", sheet.Cell("B3").GetValue<string>());
            Assert.Equal("Name", sheet.Cell("B4").GetValue<string>());
            Assert.Equal("Later", sheet.Cell("D4").GetValue<string>());
            Assert.Equal("first", sheet.Cell("B5").GetValue<string>());
            Assert.Equal("second", sheet.Cell("B6").GetValue<string>());
            Assert.True(string.IsNullOrEmpty(sheet.Cell("D5").GetValue<string>()));
            Assert.Equal(19, sheet.Cell("D6").GetValue<int>());
            Assert.Equal("Finance", saved.Properties.Creator);
            Assert.Single(sheet.GetTables());
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LatePropertyOrUpstreamFailureLeavesTheDestinationUnchanged(bool upstreamFailure) {
        InDirectory(path => {
            byte[] original = [1, 2, 3, 4];
            File.WriteAllBytes(path, original);
            using var runspace = CreateRunspace(path);
            string producer = upstreamFailure
                ? "[pscustomobject]@{ Name = 'first' }; throw 'upstream failure'"
                : """
                  [pscustomobject]@{ Name = 'first' }
                  $row = [pscustomobject]@{ Name = 'second' }
                  $row | Add-Member -MemberType ScriptProperty -Name Later -Value { throw 'late property failure' }
                  $row
                  """;
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("$ErrorActionPreference = 'Stop'; & { " + producer + " } | Export-OfficeExcel -Path $destination -FlattenOptions @{ PropertyConversionErrorAction = 'Stop' }");
            try { shell.Invoke(); } catch (RuntimeException) { }
            Assert.True(shell.HadErrors);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Equal(new[] { path }, Directory.GetFiles(Path.GetDirectoryName(path)!));
        });
    }

    [Theory]
    [InlineData("-WhatIf")]
    [InlineData("-NoClobber")]
    public void DeclinedCreateDoesNotConsumeOrCloseBorrowedReader(string policy) {
        InDirectory(path => {
            File.WriteAllBytes(path, [1, 2, 3]);
            using var table = new DataTable();
            table.Columns.Add("Name");
            table.Rows.Add("first");
            using var reader = table.CreateDataReader();
            using var runspace = CreateRunspace(path);
            runspace.SessionStateProxy.SetVariable("reader", reader);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("Export-OfficeExcel -Path $destination -InputObject $reader " + policy);
            shell.Invoke();
            Assert.Equal(policy == "-NoClobber", shell.HadErrors);
            Assert.False(reader.IsClosed);
            Assert.True(reader.Read());
            Assert.Equal("first", reader.GetString(0));
            Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(path));
        });
    }

    [Fact]
    public void EmptyBorrowedReaderRetainsItsSchemaWithoutInventingDataRows() {
        InDirectory(path => {
            using var table = new DataTable();
            table.Columns.Add("Name");
            table.Columns.Add("Amount", typeof(decimal));
            using var reader = table.CreateDataReader();
            using var runspace = CreateRunspace(path);
            runspace.SessionStateProxy.SetVariable("reader", reader);
            Invoke(runspace, "Export-OfficeExcel -Path $destination -InputObject $reader -NoTable");
            Assert.False(reader.IsClosed);
            using var saved = ExcelDocument.Load(path);
            Assert.Equal("Name", saved["Sheet1"].Cell("A1").GetValue<string>());
            Assert.Equal("Amount", saved["Sheet1"].Cell("B1").GetValue<string>());
            Assert.True(string.IsNullOrEmpty(saved["Sheet1"].Cell("A2").GetValue<string>()));
        });
    }

    private static Runspace CreateRunspace(string destination) {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Export-OfficeExcel", typeof(ExportOfficeExcelCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        runspace.SessionStateProxy.SetVariable("destination", destination);
        return runspace;
    }

    private static void Invoke(Runspace runspace, string script) {
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript(script);
        shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
    }

    private static void InDirectory(Action<string> action) {
        string directory = Path.Combine(Path.GetTempPath(), "PSWriteOffice-stream-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try { action(Path.Combine(directory, "report.xlsx")); }
        finally { Directory.Delete(directory, recursive: true); }
    }
}
