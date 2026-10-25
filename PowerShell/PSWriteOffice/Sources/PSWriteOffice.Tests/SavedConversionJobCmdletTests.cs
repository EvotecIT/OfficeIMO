using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;
using PSWriteOffice.Cmdlets.Conversion;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class SavedConversionJobCmdletTests {
    [Fact]
    public void SavedJobPlansWithoutWritesAndRunsRelativeToItsFileThenResumes() {
        string root = Path.Combine(AppContext.BaseDirectory, "officeimo-saved-command-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(root, "input");
        string jobs = Path.Combine(root, "jobs");
        Directory.CreateDirectory(input);
        Directory.CreateDirectory(jobs);
        try {
            File.WriteAllText(Path.Combine(input, "report.txt"), "Shared saved conversion job");
            string output = Path.Combine(root, "output");
            string checkpoint = Path.Combine(root, "checkpoint");
            string path = Path.Combine(jobs, "monthly.officejob.json");
            using var runspace = CreateRunspace();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("Export-OfficeConversionJob").AddParameter("InputDirectory", input)
                .AddParameter("OutputDirectory", output).AddParameter("CheckpointDirectory", checkpoint).AddParameter("Path", path);
            Assert.IsType<FileInfo>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            OfficeConversionJob saved = OfficeConversionJobSerializer.ReadFile(path);
            Assert.False(Path.IsPathRooted(saved.InputDirectory!));
            Assert.False(Path.IsPathRooted(saved.OutputDirectory));
            shell.Commands.Clear();
            shell.AddCommand("Invoke-OfficeConversionJob").AddParameter("Path", path).AddParameter("Plan");
            var plan = Assert.IsType<OfficeConversionJobPlan>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.True(plan.CanRun);
            Assert.Equal(1, plan.Selected);
            Assert.False(Directory.Exists(output));
            Assert.False(Directory.Exists(checkpoint));
            shell.Commands.Clear();
            string whatIfReport = Path.Combine(jobs, "whatif-run.json");
            shell.AddCommand("Invoke-OfficeConversionJob").AddParameter("Path", path).AddParameter("WhatIf").AddParameter("ReportPath", whatIfReport);
            Assert.Empty(shell.Invoke());
            Assert.False(File.Exists(whatIfReport));
            Assert.False(Directory.Exists(output));
            Assert.False(Directory.Exists(checkpoint));
            shell.Commands.Clear();
            string firstReportPath = Path.Combine(jobs, "first-run.json");
            shell.AddCommand("Invoke-OfficeConversionJob").AddParameter("Path", path).AddParameter("ReportPath", firstReportPath);
            var first = Assert.IsType<OfficeConversionBatchResult>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.Equal(1, first.Completed);
            string pdf = Path.Combine(output, "report.txt.pdf");
            Assert.Contains("Shared saved conversion job", PdfDocument.Load(pdf).Read().Text);
            byte[] firstBytes = File.ReadAllBytes(pdf);
            byte[] firstReportBytes = File.ReadAllBytes(firstReportPath);
            shell.Commands.Clear();
            shell.AddCommand("Get-OfficeConversionJobReport").AddParameter("Path", firstReportPath);
            var report = Assert.IsType<OfficeConversionJobRunReport>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.Equal(1, report.Batch!.Summary.Completed);
            Assert.Single(report.Batch.Items);
            shell.Commands.Clear();
            shell.AddCommand("Invoke-OfficeConversionJob").AddParameter("Path", path).AddParameter("SummaryVariable", "summary")
                .AddParameter("ReportPath", Path.Combine(jobs, "resumed-run.json"));
            var resumed = Assert.IsType<OfficeConversionBatchResult>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.Equal(1, resumed.Reused);
            Assert.Same(resumed, runspace.SessionStateProxy.GetVariable("summary"));
            Assert.Equal(firstBytes, File.ReadAllBytes(pdf));
            Assert.Equal(firstReportBytes, File.ReadAllBytes(firstReportPath));
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        } finally {
            Directory.Delete(root, true);
        }
    }

    [Fact]
    public void ExportWhatIfDoesNotWriteAndOnlyEnvironmentNamesAreSaved() {
        string root = Path.Combine(AppContext.BaseDirectory, "officeimo-job-export-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "report.txt");
            File.WriteAllText(source, "Selected-file saved job");
            string path = Path.Combine(root, "selected.officejob.json");
            using var runspace = CreateRunspace();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            void AddExport() => shell.AddCommand("Export-OfficeConversionJob")
                .AddParameter("InputPath", new[] { new FileInfo(source) })
                .AddParameter("OutputDirectory", Path.Combine(root, "output"))
                .AddParameter("SourcePasswordEnvironmentVariable", "OFFICEIMO_TEST_INPUT_PASSWORD")
                .AddParameter("Path", path);
            AddExport();
            shell.AddParameter("WhatIf");
            Assert.Empty(shell.Invoke());
            Assert.False(File.Exists(path));
            shell.Commands.Clear();
            AddExport();
            Assert.Single(shell.Invoke());
            OfficeConversionJob job = OfficeConversionJobSerializer.ReadFile(path);
            Assert.Single(job.InputPaths!);
            Assert.Equal("OFFICEIMO_TEST_INPUT_PASSWORD", job.Secrets.SourcePasswordEnvironmentVariable);
            byte[] original = File.ReadAllBytes(path);
            shell.Commands.Clear();
            AddExport();
            Assert.Throws<CmdletInvocationException>(() => shell.Invoke());
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally {
            Directory.Delete(root, true);
        }
    }

    [Fact]
    public void SavedJobErrorActionStopRetainsSummaryAndStreamsItemEvidence() {
        string root = Path.Combine(AppContext.BaseDirectory, "officeimo-job-failed-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(root, "input");
        Directory.CreateDirectory(input);
        try {
            File.WriteAllText(Path.Combine(input, "broken.docx"), "Malformed package");
            string jobPath = Path.Combine(root, "broken.officejob.json");
            OfficeConversionJobSerializer.SaveFile(OfficeConversionJob.FromRequest(new OfficeConversionBatchRequest {
                InputDirectory = input,
                OutputDirectory = Path.Combine(root, "output")
            }, root), jobPath);
            using var runspace = CreateRunspace();
            runspace.SessionStateProxy.SetVariable("jobPath", jobPath);
            string reportPath = Path.Combine(root, "failed-run.json");
            runspace.SessionStateProxy.SetVariable("reportPath", reportPath);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("$caught = $null; try { Invoke-OfficeConversionJob -Path $jobPath -ReportPath $reportPath -ItemResults -SummaryVariable summary -ErrorAction Stop } catch { $caught = $_.FullyQualifiedErrorId }");
            var results = shell.Invoke();
            Assert.Equal(OfficeWorkflowStatus.Failed, Assert.IsType<OfficeConversionBatchItemResult>(Assert.Single(results).BaseObject).Status);
            Assert.StartsWith("OfficeConversionJobFailed", Assert.IsType<string>(runspace.SessionStateProxy.GetVariable("caught")));
            Assert.Equal(1, Assert.IsType<OfficeConversionBatchResult>(runspace.SessionStateProxy.GetVariable("summary")).Failed);
            Assert.Equal(1, OfficeConversionJobRunReportSerializer.ReadFile(reportPath).Batch!.Summary.Failed);
        } finally {
            Directory.Delete(root, true);
        }
    }

    [Fact]
    public void EmptyPlanCanBeExplicitlyAcceptedWithoutWritingReportOrOutput() {
        string root = Path.Combine(AppContext.BaseDirectory, "officeimo-empty-job-plan-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(root, "input");
        Directory.CreateDirectory(input);
        try {
            File.WriteAllText(Path.Combine(input, "unsupported.bin"), "No executable route");
            string output = Path.Combine(root, "output");
            string jobPath = Path.Combine(root, "empty.json");
            OfficeConversionJobSerializer.SaveFile(new OfficeConversionJob { InputDirectory = "input", OutputDirectory = "output" }, jobPath);
            using var runspace = CreateRunspace();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("Invoke-OfficeConversionJob").AddParameter("Path", jobPath).AddParameter("Plan").AddParameter("AllowEmpty");
            var plan = Assert.IsType<OfficeConversionJobPlan>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.Equal(0, plan.Selected);
            Assert.Equal(0, plan.Failed);
            Assert.False(plan.CanRun);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.False(Directory.Exists(output));
            shell.Commands.Clear();
            shell.AddCommand("Invoke-OfficeConversionJob").AddParameter("Path", jobPath).AddParameter("Plan");
            Assert.Single(shell.Invoke());
            Assert.StartsWith("OfficeConversionJobPlanInvalid", Assert.Single(shell.Streams.Error).FullyQualifiedErrorId);
        } finally {
            Directory.Delete(root, true);
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    public void ExplicitEmptyReportPathFailsBeforeConversion(string reportPath) {
        string root = Path.Combine(AppContext.BaseDirectory, "officeimo-empty-report-path-" + Guid.NewGuid().ToString("N"));
        string source = Path.Combine(root, "input");
        Directory.CreateDirectory(source);
        try {
            File.WriteAllText(Path.Combine(source, "one.txt"), "Do not silently skip requested recording");
            string job = Path.Combine(root, "job.json");
            string output = Path.Combine(root, "output");
            OfficeConversionJobSerializer.SaveFile(new OfficeConversionJob { InputDirectory = "input", OutputDirectory = "output" }, job);
            using var runspace = CreateRunspace();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("Invoke-OfficeConversionJob").AddParameter("Path", job).AddParameter("ReportPath", reportPath);
            Assert.Throws<CmdletInvocationException>(() => shell.Invoke());
            Assert.False(Directory.Exists(output));
            Assert.Equal("Do not silently skip requested recording", File.ReadAllText(Path.Combine(source, "one.txt")));
        } finally {
            Directory.Delete(root, true);
        }
    }

    private static Runspace CreateRunspace() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Export-OfficeConversionJob", typeof(ExportOfficeConversionJobCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Invoke-OfficeConversionJob", typeof(InvokeOfficeConversionJobCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Get-OfficeConversionJobReport", typeof(GetOfficeConversionJobReportCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
