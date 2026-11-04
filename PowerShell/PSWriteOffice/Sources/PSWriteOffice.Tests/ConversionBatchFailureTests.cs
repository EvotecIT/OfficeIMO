using OfficeIMO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;
using PSWriteOffice.Cmdlets.Conversion;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class ConversionBatchFailureTests {
    [Fact]
    public void DiscoveryAbortRetainsCompletedArtifactAndSummaryBeforeErrorActionStop() {
        string root = Path.Combine(AppContext.BaseDirectory, "officeimo-batch-discovery-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(root, "input");
        Directory.CreateDirectory(input);
        try {
            File.WriteAllText(Path.Combine(input, "first.txt"), "Retained completed artifact");
            File.WriteAllText(Path.Combine(input, "second.txt"), "Retained completed artifact");
            using var runspace = CreateRunspace();
            runspace.SessionStateProxy.SetVariable("inputRoot", input);
            string output = Path.Combine(root, "output");
            runspace.SessionStateProxy.SetVariable("outputRoot", output);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("$continued = $false; $caught = $null; try { Convert-OfficeDocument -LiteralPath $inputRoot -To Pdf -OutputDirectory $outputRoot" +
                " -MaximumFiles 1 -ThrottleLimit 1 -SummaryVariable summary -ErrorAction Stop; $continued = $true } catch { $caught = $_.FullyQualifiedErrorId }");
            shell.Invoke();
            Assert.False((bool)runspace.SessionStateProxy.GetVariable("continued"));
            Assert.StartsWith("OfficeConversionJobFailed", Assert.IsType<string>(runspace.SessionStateProxy.GetVariable("caught")));
            var summary = Assert.IsType<OfficeBatchSummary>(runspace.SessionStateProxy.GetVariable("summary"));
            Assert.Equal(1, summary.Completed);
            Assert.Equal(0, summary.Failed);
            Assert.NotNull(summary.FailureSummary);
            string artifact = Assert.Single(Directory.GetFiles(output, "*.pdf"));
            Assert.Contains("Retained completed artifact", PdfDocument.Load(artifact).Read().Text);
        } finally {
            Directory.Delete(root, true);
        }
    }

    [Fact]
    public void ErrorActionStopRetainsSummaryAndCompletedArtifacts() {
        string root = Path.Combine(AppContext.BaseDirectory, "officeimo-batch-errors-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(root, "input");
        Directory.CreateDirectory(input);
        try {
            File.WriteAllText(Path.Combine(input, "good.txt"), "Completed before reporting batch failures");
            File.WriteAllText(Path.Combine(input, "broken.docx"), "This is not an Open XML package.");
            using var runspace = CreateRunspace();
            runspace.SessionStateProxy.SetVariable("inputRoot", input);
            string output = Path.Combine(root, "output");
            runspace.SessionStateProxy.SetVariable("outputRoot", output);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("$continued = $false; $caught = $null; try { Convert-OfficeDocument -LiteralPath $inputRoot -To Pdf -OutputDirectory $outputRoot" +
                " -ThrottleLimit 1 -SummaryVariable summary -ErrorAction Stop; $continued = $true } catch { $caught = $_.FullyQualifiedErrorId }");
            shell.Invoke();
            Assert.False((bool)runspace.SessionStateProxy.GetVariable("continued"));
            Assert.StartsWith("OfficeConversionJobFailed", Assert.IsType<string>(runspace.SessionStateProxy.GetVariable("caught")));
            var summary = Assert.IsType<OfficeBatchSummary>(runspace.SessionStateProxy.GetVariable("summary"));
            Assert.Equal(1, summary.Completed);
            Assert.Equal(1, summary.Failed);
            Assert.Contains("Completed before reporting", PdfDocument.Load(Path.Combine(output, "good.pdf")).Read().Text);
            Assert.False(File.Exists(Path.Combine(output, "broken.pdf")));
        } finally {
            Directory.Delete(root, true);
        }
    }

    [Fact]
    public void AllSkippedIsAnErrorUnlessEmptySelectionIsExplicit() {
        string root = Path.Combine(AppContext.BaseDirectory, "officeimo-batch-empty-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(root, "input");
        Directory.CreateDirectory(input);
        try {
            File.WriteAllText(Path.Combine(input, "unsupported.bin"), "Unsupported source");
            using var runspace = CreateRunspace();
            runspace.SessionStateProxy.SetVariable("inputRoot", input);
            runspace.SessionStateProxy.SetVariable("outputRoot", Path.Combine(root, "output"));
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            string script = "Convert-OfficeDocument -LiteralPath $inputRoot -To Pdf -OutputDirectory $outputRoot -SummaryVariable summary";
            shell.AddScript(script);
            var output = shell.Invoke();
            Assert.StartsWith("OfficeConversionJobEmpty", Assert.Single(shell.Streams.Error).FullyQualifiedErrorId);
            var summary = Assert.IsType<OfficeBatchSummary>(runspace.SessionStateProxy.GetVariable("summary"));
            Assert.Equal(0, summary.Selected);
            Assert.Equal(1, summary.Skipped);
            shell.Commands.Clear();
            shell.Streams.Error.Clear();
            shell.AddScript(script + " -AllowEmpty");
            Assert.Single(shell.Invoke());
            Assert.Empty(shell.Streams.Error);
        } finally {
            Directory.Delete(root, true);
        }
    }

    private static Runspace CreateRunspace() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Convert-OfficeDocument", typeof(ConvertOfficeDocumentCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
