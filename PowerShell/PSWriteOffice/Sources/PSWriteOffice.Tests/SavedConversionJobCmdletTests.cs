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
            shell.AddCommand("Invoke-OfficeConversionJob").AddParameter("Path", path).AddParameter("WhatIf");
            Assert.Empty(shell.Invoke());
            Assert.False(Directory.Exists(output));
            Assert.False(Directory.Exists(checkpoint));
            shell.Commands.Clear();
            shell.AddCommand("Invoke-OfficeConversionJob").AddParameter("Path", path);
            var first = Assert.IsType<OfficeConversionBatchResult>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.Equal(1, first.Completed);
            string pdf = Path.Combine(output, "report.txt.pdf");
            Assert.Contains("Shared saved conversion job", PdfDocument.Load(pdf).Read().Text);
            byte[] firstBytes = File.ReadAllBytes(pdf);
            shell.Commands.Clear();
            shell.AddCommand("Invoke-OfficeConversionJob").AddParameter("Path", path).AddParameter("SummaryVariable", "summary");
            var resumed = Assert.IsType<OfficeConversionBatchResult>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.Equal(1, resumed.Reused);
            Assert.Same(resumed, runspace.SessionStateProxy.GetVariable("summary"));
            Assert.Equal(firstBytes, File.ReadAllBytes(pdf));
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
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("$caught = $null; try { Invoke-OfficeConversionJob -Path $jobPath -ItemResults -SummaryVariable summary -ErrorAction Stop } catch { $caught = $_.FullyQualifiedErrorId }");
            var results = shell.Invoke();
            Assert.Equal(OfficeWorkflowStatus.Failed, Assert.IsType<OfficeConversionBatchItemResult>(Assert.Single(results).BaseObject).Status);
            Assert.StartsWith("OfficeConversionJobFailed", Assert.IsType<string>(runspace.SessionStateProxy.GetVariable("caught")));
            Assert.Equal(1, Assert.IsType<OfficeConversionBatchResult>(runspace.SessionStateProxy.GetVariable("summary")).Failed);
        } finally {
            Directory.Delete(root, true);
        }
    }

    private static Runspace CreateRunspace() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Export-OfficeConversionJob", typeof(ExportOfficeConversionJobCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Invoke-OfficeConversionJob", typeof(InvokeOfficeConversionJobCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
