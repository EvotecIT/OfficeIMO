using Avalonia;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioWorkflowJobIntegrationTests {
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
