using OfficeIMO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;
using PSWriteOffice.Cmdlets.Conversion;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class GeneralConversionCmdletTests {
    [Fact]
    public void SingleAndSelectedFileJobsShareRoutingAndPreserveConflictingOutput() {
        // Batch checkpoints reject symlink ancestors, including macOS's /var temp alias.
        string root = Path.Combine(AppContext.BaseDirectory, "officeimo-conversion-command-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "literal.txt");
            string output = Path.Combine(root, "literal.pdf");
            File.WriteAllText(source, "# Literal text remains text");
            var state = InitialSessionState.CreateDefault();
            state.ImportPSModule(new[] { typeof(ConvertToOfficeDocumentCommand).Assembly.Location });
            using var runspace = RunspaceFactory.CreateRunspace(state);
            runspace.Open();
            runspace.SessionStateProxy.SetVariable("sourcePath", source);
            runspace.SessionStateProxy.SetVariable("outputPath", output);
            runspace.SessionStateProxy.SetVariable("batchPath", Path.Combine(root, "batch"));
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                ConvertTo-OfficeDocument -InputPath $sourcePath -OutputPath $outputPath -WhatIf
                if (Test-Path -LiteralPath $outputPath) { throw 'WhatIf wrote a PDF' }
                $single = ConvertTo-OfficeDocument -InputPath $sourcePath -OutputPath $outputPath -ConflictPolicy fail -OutputProfile faithful
                Get-Item -LiteralPath $sourcePath | ConvertTo-OfficeDocument -OutputDirectory $batchPath -TargetExtension pdf -ConflictPolicy rEpLaCe -OutputProfile fAiThFuL -ItemResults -SummaryVariable batchSummary
                """);
            var items = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(OfficeBatchItemStatus.Completed, Assert.IsType<OfficeBatchItemReport>(Assert.Single(items).BaseObject).Status);
            var single = Assert.IsType<OfficeConversionReport>(PSObject.AsPSObject(runspace.SessionStateProxy.GetVariable("single")).BaseObject);
            Assert.Equal("Pdf", single.Target);
            Assert.True(single.ItemsWritten > 0);
            Assert.Contains("# Literal text", PdfDocument.Load(output).Read().Text);
            byte[] original = File.ReadAllBytes(output);
            shell.Commands.Clear();
            shell.AddCommand("ConvertTo-OfficeDocument").AddParameter("InputPath", source).AddParameter("OutputPath", output);
            shell.Invoke();
            Assert.True(shell.HadErrors);
            Assert.Equal(original, File.ReadAllBytes(output));
        } finally {
            Directory.Delete(root, true);
        }
    }
}
