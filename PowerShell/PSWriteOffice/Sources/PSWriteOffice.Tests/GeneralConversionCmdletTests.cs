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
        string root = Path.Combine(Path.GetTempPath(), "officeimo-conversion-command-" + Guid.NewGuid().ToString("N"));
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
                $single = ConvertTo-OfficeDocument -InputPath $sourcePath -OutputPath $outputPath
                Get-Item -LiteralPath $sourcePath | ConvertTo-OfficeDocument -OutputDirectory $batchPath -TargetExtension pdf -ItemResults -SummaryVariable batchSummary
                """);
            var items = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(OfficeWorkflowStatus.Completed, Assert.IsType<OfficeConversionBatchItemResult>(Assert.Single(items).BaseObject).Status);
            var single = Assert.IsType<OfficeWorkflowResult>(PSObject.AsPSObject(runspace.SessionStateProxy.GetVariable("single")).BaseObject);
            Assert.True(single.Succeeded);
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
