using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Excel;
using PSWriteOffice.Cmdlets.Excel;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ExcelOpenLifecycleWorkflowTests {
    [Theory]
    [InlineData(null)]
    [InlineData("test-secret")]
    public void LocalBlocksPublishOnlyAfterSuccessAndPreserveSourceEncryption(string? password) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-excel-open-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "Review[final].xlsx");
            using (var source = ExcelDocument.Create()) {
                source.AddSheet("Data").CellValue(1, 1, "Original");
                source.Save(path, new ExcelSaveOptions { Password = password });
            }
            using var runspace = Open();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("Open-OfficeExcel").AddParameter("LiteralPath", path)
                .AddParameter("Password", password).AddParameter("PassThru")
                .AddParameter("SaveOptions", new System.Collections.Hashtable { ["ValidateOpenXml"] = true })
                .AddParameter("Content", ScriptBlock.Create("ExcelSheet 'Data' { ExcelCell -Address A1 -Value 'Reviewed' }"));
            FileInfo result = Assert.IsType<FileInfo>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.False(shell.HadErrors, Errors(shell));
            Assert.Equal(path, result.FullName);
            using (var saved = ExcelDocument.Load(path, new ExcelLoadOptions { Password = password })) {
                Assert.True(saved.Sheets[0].TryGetCellText(1, 1, out string text));
                Assert.Equal("Reviewed", text);
            }
            Assert.Equal(password == null ? new byte[] { 0x50, 0x4B, 0x03, 0x04 } : new byte[] { 0xD0, 0xCF, 0x11, 0xE0 },
                File.ReadAllBytes(path).Take(4));
            byte[] original = File.ReadAllBytes(path);
            shell.Commands.Clear();
            shell.AddCommand("Open-OfficeExcel").AddParameter("LiteralPath", path).AddParameter("Password", password)
                .AddParameter("Content", ScriptBlock.Create("ExcelSheet 'Data' { ExcelCell -Address A1 -Value 'Discarded' }; throw 'authoring-failed'"));
            Exception? failure = Record.Exception(() => shell.Invoke());
            Assert.True(failure != null || shell.HadErrors);
            Assert.Contains("authoring-failed", failure?.Message ?? Errors(shell));
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Single(Directory.GetFiles(root));

            foreach (string policy in new[] { "ReadOnly", "WhatIf" }) {
                shell.Commands.Clear();
                shell.Streams.Error.Clear();
                runspace.SessionStateProxy.SetVariable("entered", false);
                shell.AddCommand("Open-OfficeExcel").AddParameter("LiteralPath", path).AddParameter("Password", password)
                    .AddParameter(policy).AddParameter("Content", ScriptBlock.Create("$global:entered = $true"));
                failure = Record.Exception(() => shell.Invoke());
                if (policy == "ReadOnly") Assert.True(failure != null || shell.HadErrors);
                else { Assert.Null(failure); Assert.False(shell.HadErrors, Errors(shell)); }
                Assert.False((bool)runspace.SessionStateProxy.GetVariable("entered"));
                Assert.Equal(original, File.ReadAllBytes(path));
            }
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void RemoteInputRejectsAuthoringAtBindingBeforeAnyDownload() {
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("try { Open-OfficeExcel -Uri 'https://127.0.0.1:1/report.xlsx' -Content {} -ErrorAction Stop } catch { $_.FullyQualifiedErrorId }");
        Assert.StartsWith("AmbiguousParameterSet", Assert.IsType<string>(Assert.Single(shell.Invoke()).BaseObject));
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.ImportPSModule(new[] { typeof(OpenOfficeExcelCommand).Assembly.Location });
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }

    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error);
}
