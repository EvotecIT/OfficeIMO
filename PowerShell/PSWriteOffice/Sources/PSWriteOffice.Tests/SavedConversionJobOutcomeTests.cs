using OfficeIMO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Workflows;
using PSWriteOffice.Services;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class SavedConversionJobOutcomeTests {
    [Fact]
    public async Task ErrorActionStopPreservesPartialConversionAndFinalRecordingFailureTogether() {
        string root = Path.Combine(AppContext.BaseDirectory, "officeimo-combined-job-failure-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(root, "input");
        Directory.CreateDirectory(input);
        try {
            File.WriteAllText(Path.Combine(input, "one.txt"), "Completed output is still available");
            File.WriteAllText(Path.Combine(input, "two.docx"), "Invalid document package");
            string jobPath = Path.Combine(root, "job.json"), reportPath = Path.Combine(root, "run.json");
            var job = new OfficeBatchOptions { InputDirectory = "input", OutputDirectory = "output", MaxParallelism = 1 };
            job.Save(jobPath);
            job = OfficeBatchOptions.Load(jobPath);
            job.ReportPath = reportPath;
            bool replaced = false;
            var report = await OfficeConverter.ConvertAsync(job,
                progress: new Observer(_ => {
                    if (replaced) return;
                    File.WriteAllText(reportPath, "External replacement must survive");
                    replaced = true;
                }));
            Assert.Equal(1, report.Summary.Completed);
            Assert.Equal(1, report.Summary.Failed);
            Assert.NotNull(report.RecordingError);

            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Test-RecordedJobOutcome", typeof(RecordedJobOutcomeCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state);
            runspace.Open();
            runspace.SessionStateProxy.SetVariable("report", report);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("$continued = $false; $caught = $null; try { Test-RecordedJobOutcome -Report $report -SummaryVariable summary -ErrorAction Stop; $continued = $true } catch { $caught = $_ }");
            Assert.Empty(shell.Invoke());
            Assert.False((bool)runspace.SessionStateProxy.GetVariable("continued"));
            var caught = Assert.IsType<ErrorRecord>(runspace.SessionStateProxy.GetVariable("caught"));
            Assert.StartsWith("OfficeConversionJobReportFailed", caught.FullyQualifiedErrorId);
            Assert.Contains("1 conversion(s) failed", caught.Exception.Message);
            Assert.Contains("run report could not be finalized", caught.Exception.Message);
            Assert.Same(report, caught.TargetObject);
            Assert.Same(report.Summary, runspace.SessionStateProxy.GetVariable("summary"));
            Assert.Equal("External replacement must survive", File.ReadAllText(reportPath));
            Assert.Contains("Completed output", OfficeIMO.Pdf.PdfDocument.Load(Path.Combine(root, "output", "one.pdf")).Read().Text);
        } finally {
            Directory.Delete(root, true);
        }
    }

    private sealed class Observer(Action<OfficeBatchItemReport> action) : IProgress<OfficeBatchItemReport> {
        public void Report(OfficeBatchItemReport value) => action(value);
    }
}

/// <summary>Hosts the production saved-job completion adapter in a real PowerShell pipeline.</summary>
[Cmdlet(VerbsDiagnostic.Test, "RecordedJobOutcome")]
public sealed class RecordedJobOutcomeCommand : AsyncPSCmdlet {
    [Parameter(Mandatory = true)]
    public OfficeBatchReport Report { get; set; } = null!;
    [Parameter]
    public string? SummaryVariable { get; set; }
    protected override Task ProcessRecordAsync() {
        OfficeWorkflowCommandUtilities.WriteSavedJobOutcome(this, Report.Summary, Report, SummaryVariable, false, false);
        return Task.CompletedTask;
    }
}
