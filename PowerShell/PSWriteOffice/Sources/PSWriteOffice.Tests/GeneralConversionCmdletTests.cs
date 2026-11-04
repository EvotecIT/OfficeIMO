using OfficeIMO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;
using PSWriteOffice.Cmdlets.Conversion;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class GeneralConversionCmdletTests {
    [Theory]
    [InlineData("single")]
    [InlineData("directory")]
    [InlineData("job")]
    public void ConversionEntrypointsKeepDotNetOnlyImageCodecsUnavailable(string entrypoint) {
        string root = Path.Combine(AppContext.BaseDirectory, "officeimo-dotnet-image-codec-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(root, "input");
        string output = Path.Combine(root, "output");
        Directory.CreateDirectory(input);
        try {
            string source = Path.Combine(input, "still.heic");
            // The unsupported extension must never reach image decoding in this host.
            File.WriteAllBytes(source, new byte[] { 0 });
            OfficeWorkflow.RegisterConversions();
            string job = Path.Combine(root, "image.officejob.json");
            new OfficeBatchOptions { InputPaths = new[] { source }, OutputDirectory = output, TargetExtension = ".pdf" }.Save(job);
            var state = InitialSessionState.CreateDefault();
            state.ImportPSModule(new[] { typeof(ConvertOfficeDocumentCommand).Assembly.Location });
            using var runspace = RunspaceFactory.CreateRunspace(state);
            runspace.Open();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            if (entrypoint == "job")
                shell.AddCommand("Convert-OfficeDocument").AddParameter("JobPath", job).AddParameter("Plan");
            else if (entrypoint == "directory")
                shell.AddCommand("Convert-OfficeDocument").AddParameter("LiteralPath", input)
                    .AddParameter("OutputDirectory", output).AddParameter("To", "pdf").AddParameter("Plan");
            else
                shell.AddCommand("Convert-OfficeDocument").AddParameter("LiteralPath", source)
                    .AddParameter("OutputPath", Path.Combine(output, "still.pdf")).AddParameter("To", "pdf").AddParameter("WhatIf");
            if (entrypoint == "single") {
                var error = Assert.Throws<CmdletInvocationException>(() => shell.Invoke());
                Assert.IsType<PSNotSupportedException>(error.InnerException);
                Assert.Contains("not available in PSWriteOffice", error.Message);
            } else {
                var result = Assert.Single(shell.Invoke()).BaseObject;
                var plan = Assert.IsType<OfficeBatchPlan>(result);
                Assert.Equal(0, plan.Selected);
                Assert.Equal(1, plan.Skipped);
                Assert.Equal(0, plan.Failed);
                Assert.True(shell.HadErrors);
            }
            if (Directory.Exists(output)) Assert.Empty(Directory.GetFiles(output));
            Assert.Equal(new byte[] { 0 }, File.ReadAllBytes(source));
        } finally {
            Directory.Delete(root, true);
        }
    }

    [Theory]
    [InlineData("directory")]
    [InlineData("job")]
    public void MixedBatchesConvertSupportedInputsAndSkipDotNetOnlyImageCodecs(string entrypoint) {
        string root = Path.Combine(AppContext.BaseDirectory, "officeimo-mixed-image-codec-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(root, "input"), output = Path.Combine(root, "output");
        Directory.CreateDirectory(input);
        try {
            File.WriteAllText(Path.Combine(input, "report.txt"), "Supported mixed-batch text");
            File.WriteAllBytes(Path.Combine(input, "still.heic"), new byte[] { 0 });
            OfficeWorkflow.RegisterConversions();
            string job = Path.Combine(root, "mixed.officejob.json");
            new OfficeBatchOptions { InputDirectory = input, OutputDirectory = output, TargetExtension = ".pdf" }.Save(job);
            var state = InitialSessionState.CreateDefault();
            state.ImportPSModule(new[] { typeof(ConvertOfficeDocumentCommand).Assembly.Location });
            using var runspace = RunspaceFactory.CreateRunspace(state);
            runspace.Open();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            if (entrypoint == "job")
                shell.AddCommand("Convert-OfficeDocument").AddParameter("JobPath", job);
            else
                shell.AddCommand("Convert-OfficeDocument").AddParameter("LiteralPath", input)
                    .AddParameter("OutputDirectory", output).AddParameter("To", "pdf");
            shell.AddParameter("SummaryVariable", "summary");
            shell.Invoke();
            var summary = Assert.IsType<OfficeBatchSummary>(runspace.SessionStateProxy.GetVariable("summary"));
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(1, summary.Completed);
            Assert.Equal(1, summary.Skipped);
            Assert.Equal(0, summary.Failed);
            Assert.Null(summary.FailureSummary);
            Assert.Contains("Supported mixed-batch text", PdfDocument.Load(Path.Combine(output, "report.pdf")).Read().Text);
            Assert.False(File.Exists(Path.Combine(output, "still.pdf")));
        } finally {
            Directory.Delete(root, true);
        }
    }

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
            state.ImportPSModule(new[] { typeof(ConvertOfficeDocumentCommand).Assembly.Location });
            using var runspace = RunspaceFactory.CreateRunspace(state);
            runspace.Open();
            runspace.SessionStateProxy.SetVariable("sourcePath", source);
            runspace.SessionStateProxy.SetVariable("outputPath", output);
            runspace.SessionStateProxy.SetVariable("batchPath", Path.Combine(root, "batch"));
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                Convert-OfficeDocument -LiteralPath $sourcePath -OutputPath $outputPath -WhatIf
                if (Test-Path -LiteralPath $outputPath) { throw 'WhatIf wrote a PDF' }
                $single = Convert-OfficeDocument -LiteralPath $sourcePath -OutputPath $outputPath -NoClobber -OutputProfile Faithful
                Get-Item -LiteralPath $sourcePath | Convert-OfficeDocument -OutputDirectory $batchPath -To Pdf -OutputProfile Faithful -SummaryVariable batchSummary
                """);
            var items = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal("Pdf", Assert.IsType<OfficeConversionReport>(Assert.Single(items).BaseObject).Target);
            var single = Assert.IsType<OfficeConversionReport>(PSObject.AsPSObject(runspace.SessionStateProxy.GetVariable("single")).BaseObject);
            Assert.Equal("Pdf", single.Target);
            Assert.True(single.ItemsWritten > 0);
            Assert.Contains("# Literal text", PdfDocument.Load(output).Read().Text);
            byte[] original = File.ReadAllBytes(output);
            shell.Commands.Clear();
            shell.AddCommand("Convert-OfficeDocument").AddParameter("LiteralPath", source).AddParameter("OutputPath", output).AddParameter("NoClobber");
            var skipped = Assert.IsType<OfficeBatchItemReport>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.True(skipped.Skipped);
            Assert.False(shell.HadErrors);
            Assert.Equal(original, File.ReadAllBytes(output));
        } finally {
            Directory.Delete(root, true);
        }
    }
}
