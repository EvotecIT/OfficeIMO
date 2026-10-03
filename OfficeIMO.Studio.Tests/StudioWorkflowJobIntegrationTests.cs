using Avalonia;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Workflows;
using System.Text.Json;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioWorkflowJobIntegrationTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProvenanceReportExportRechecksPermissionAndCancellationBeforeCommit(bool cancel) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string output = Path.Combine(services.Paths.Root, "output");
            Directory.CreateDirectory(output);
            string source = Path.Combine(services.Paths.Root, "source.html");
            const string original = "<html><body>Preserved source</body></html>";
            await File.WriteAllTextAsync(source, original);
            ProvenanceWorkbenchViewModel? workspace = null;
            int checks = 0;
            var guard = new StudioWorkflowPublicationGuard((path, _) => {
                if (++checks == 1) return true;
                Assert.Single(Directory.GetFiles(output));
                Assert.False(File.Exists(path));
                if (cancel) workspace!.CancelCommand.Execute(null);
                return cancel;
            });
            using var provenance = new ProvenanceWorkbenchViewModel(_ => Task.FromResult<string?>(null),
                _ => Task.FromResult<string?>(null), null, guard, services.Jobs);
            workspace = provenance;
            provenance.InputPath = source; provenance.OutputFolder = output;
            await provenance.AssessCommand.ExecuteAsync(null);

            await provenance.ExportReportCommand.ExecuteAsync(null);

            Assert.Equal(2, checks);
            Assert.Empty(provenance.ReportPath);
            Assert.Empty(Directory.GetFiles(output));
            Assert.Equal(cancel ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed, services.Jobs.Entries[0].Outcome);
            Assert.Equal(original, await File.ReadAllTextAsync(source));
            Assert.False(services.Jobs.HasActive);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ProvenanceReportExportHonorsProtectedRecoveryDestinations() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            Directory.CreateDirectory(services.Paths.WorkflowRecoveryRoot);
            string source = Path.Combine(services.Paths.Root, "source.html");
            const string original = "<html><body>Preserved source</body></html>";
            await File.WriteAllTextAsync(source, original);
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            var provenance = shell.ProvenanceWorkbench;
            provenance.InputPath = source;
            provenance.OutputFolder = services.Paths.WorkflowRecoveryRoot;
            await provenance.AssessCommand.ExecuteAsync(null);
            Assert.True(provenance.CanExportReport, provenance.Status);

            await provenance.ExportReportCommand.ExecuteAsync(null);

            Assert.Empty(provenance.ReportPath);
            Assert.Empty(Directory.GetFiles(services.Paths.WorkflowRecoveryRoot));
            Assert.Equal(original, await File.ReadAllTextAsync(source));
            Assert.Equal(OfficeWorkflowStatus.Failed, services.Jobs.Entries[0].Outcome);
            Assert.False(services.Jobs.HasActive);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ProvenanceReportExportWaitsForBudgetCancelsFromJobsAndRetainsAValidReport() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string output = Path.Combine(services.Paths.Root, "output");
            Directory.CreateDirectory(output);
            string source = Path.Combine(services.Paths.Root, "source.html");
            const string original = "<html><body>Preserved source</body></html>";
            await File.WriteAllTextAsync(source, original);
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            var provenance = shell.ProvenanceWorkbench;
            provenance.InputPath = source; provenance.OutputFolder = output;
            await provenance.AssessCommand.ExecuteAsync(null);
            using var first = await services.Jobs.EnterAsync(default);
            using var second = await services.Jobs.EnterAsync(default);

            Task pending = provenance.ExportReportCommand.ExecuteAsync(null);
            var queued = services.Jobs.Entries[0];
            Assert.True(queued.IsActive);
            Assert.False(pending.IsCompleted);
            Assert.Empty(Directory.GetFiles(output));
            queued.CancelCommand.Execute(null);
            await pending.WaitAsync(TimeSpan.FromSeconds(30));
            Assert.Equal(OfficeWorkflowStatus.Cancelled, queued.Outcome);
            Assert.False(queued.IsActive);
            Assert.Empty(Directory.GetFiles(output));
            Assert.True(provenance.CanExportReport);

            first.Dispose();
            await provenance.ExportReportCommand.ExecuteAsync(null).WaitAsync(TimeSpan.FromSeconds(30));
            var exported = services.Jobs.Entries[0];
            Assert.True(exported.IsSucceeded, exported.Summary);
            Assert.True(exported.HasOutput);
            Assert.Equal(provenance.ReportPath, exported.OutputPath);
            Assert.Equal(provenance.ReportPath, Assert.Single(Directory.GetFiles(output)));
            using var report = JsonDocument.Parse(await File.ReadAllTextAsync(provenance.ReportPath));
            Assert.Equal("officeimo.provenance.result.v2", report.RootElement.GetProperty("schema").GetString());
            Assert.Equal(provenance.InputHash, report.RootElement.GetProperty("inputSha256").GetString());
            Assert.Equal(original, await File.ReadAllTextAsync(source));
            Assert.False(services.Jobs.HasActive);
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FolderExportAndProvenanceWaitForSharedBudgetAndCanCancelFromJobs(bool folderExport) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string input = Path.Combine(services.Paths.Root, "input"), output = Path.Combine(services.Paths.Root, "output");
            Directory.CreateDirectory(input); Directory.CreateDirectory(output);
            string source = Path.Combine(input, folderExport ? "source.txt" : "source.html");
            string original = folderExport ? "Shared job budget" : "<html><head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>Preserved text</body></html>";
            await File.WriteAllTextAsync(source, original);
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            var batch = shell.ConversionWorkbench.BatchExport;
            batch.InputDirectory = input; batch.OutputDirectory = output;
            var provenance = shell.ProvenanceWorkbench;
            provenance.InputPath = source; provenance.OutputFolder = output;
            using var first = await services.Jobs.EnterAsync(default);
            using var second = await services.Jobs.EnterAsync(default);
            Task Run() => folderExport ? batch.RunCommand.ExecuteAsync(null) : provenance.AssessCommand.ExecuteAsync(null);
            Task pending = Run();
            var queued = Assert.Single(services.Jobs.Entries);
            Assert.True(queued.IsActive);
            Assert.False(pending.IsCompleted);
            Assert.Empty(Directory.GetFiles(output));
            Assert.Empty(provenance.Findings);
            queued.CancelCommand.Execute(null);
            await pending.WaitAsync(TimeSpan.FromSeconds(30));
            Assert.Equal(OfficeWorkflowStatus.Cancelled, queued.Outcome);
            Assert.False(queued.IsActive);
            Assert.Empty(Directory.GetFiles(output));

            first.Dispose();
            await Run().WaitAsync(TimeSpan.FromSeconds(30));
            var completed = services.Jobs.Entries[0];
            Assert.True(completed.IsSucceeded, completed.Summary);
            if (folderExport) {
                Assert.True(completed.HasOutput);
                Assert.Equal(output, completed.OutputPath);
                Assert.True(File.Exists(Path.Combine(output, "source.txt.pdf")));
            } else {
                Assert.False(completed.HasOutput);
                provenance.RemoveReferences = true;
                await provenance.CreateCopyCommand.ExecuteAsync(null);
                var copy = services.Jobs.Entries[0];
                Assert.True(copy.IsSucceeded, copy.Summary);
                Assert.True(copy.HasOutput);
                Assert.Equal(provenance.OutputPath, copy.OutputPath);
                Assert.DoesNotContain("c2pa-manifest", await File.ReadAllTextAsync(copy.OutputPath!));
            }
            Assert.Equal(original, await File.ReadAllTextAsync(source));
            Assert.False(services.Jobs.HasActive);
            return true;
        }, CancellationToken.None);
    }
}
