using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioJobHistoryTests {
    [Fact]
    public void RetentionAndClearNeverRemoveActiveJobsOrDeleteOutputs() {
        var history = new StudioJobHistory(StudioLocalization.Current);
        int cancelled = 0;
        var active = history.Start("Active", "input", "output", () => cancelled++);
        for (int index = 0; index < StudioJobHistory.MaximumEntries + 5; index++) {
            history.Start("Finished", "input", "output", () => { }).Complete(OfficeWorkflowStatus.Completed, null, "Report ready");
        }
        Assert.Equal(StudioJobHistory.MaximumEntries, history.Entries.Count);
        Assert.Contains(active, history.Entries);
        Assert.Equal(1, history.ActiveCount);
        history.ClearFinished();
        Assert.Same(active, Assert.Single(history.Entries));
        active.CancelCommand.Execute(null);
        Assert.Equal(1, cancelled);
        active.Complete(OfficeWorkflowStatus.Cancelled, null, "Cancelled");
        active.CancelCommand.Execute(null);
        Assert.Equal(1, cancelled);
        history.ClearFinished();
        Assert.Empty(history.Entries);
    }

    [Fact]
    public async Task SharedBudgetLimitsActualRunsAndCancelsWaitingWorkWithoutStartingIt() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            string input = Path.Combine(services.Paths.Root, "source.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(200, 300))).Save(input);
            var runner = new BlockingRunner();
            DocumentHealthViewModel Create() => new(_ => Task.FromResult<string?>(input),
                _ => Task.FromResult<string?>(services.Paths.Root), runner, jobHistory: services.Jobs);
            using var first = Create();
            using var second = Create();
            using var third = Create();
            await first.ChooseInputCommand.ExecuteAsync(null);
            await second.ChooseInputCommand.ExecuteAsync(null);
            await third.ChooseInputCommand.ExecuteAsync(null);
            Task firstRun = first.RunCommand.ExecuteAsync(null);
            Task secondRun = second.RunCommand.ExecuteAsync(null);
            await runner.TwoStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));
            Task thirdRun = third.RunCommand.ExecuteAsync(null);
            Assert.Equal(2, runner.Started);
            Assert.Equal(3, services.Jobs.ActiveCount);
            StudioJobRecord waiting = services.Jobs.Entries[0];
            waiting.CancelCommand.Execute(null);
            await thirdRun.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.False(waiting.IsActive);
            Assert.Equal("Cancelled", waiting.Status);
            Assert.Equal(2, runner.Started);
            runner.Release.TrySetResult();
            await Task.WhenAll(firstRun, secondRun).WaitAsync(TimeSpan.FromSeconds(10));
            Assert.Equal(0, services.Jobs.ActiveCount);
            Assert.Equal(2, runner.Peak);
            Assert.Equal(2, services.Jobs.Entries.Count(entry => entry.Status == "Completed"));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task FinishedOutputActionsCheckAvailabilityAndClearingHistoryPreservesFiles() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-jobs-output-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string output = Path.Combine(root, "result.pdf");
            File.WriteAllText(output, "output retained");
            var history = new StudioJobHistory(StudioLocalization.Current);
            string? opened = null;
            using var model = new StudioJobsViewModel(history, (path, _) => { opened = path; return Task.CompletedTask; });
            var record = history.Start("Conversion", "source.html", output, () => { });
            record.Complete(OfficeWorkflowStatus.Completed, output, "Ready");
            await model.OpenOutputCommand.ExecuteAsync(record);
            Assert.Equal(output, opened);
            model.ClearFinishedCommand.Execute(null);
            Assert.Empty(history.Entries);
            Assert.Equal("output retained", File.ReadAllText(output));
            File.Delete(output);
            opened = null;
            await model.OpenOutputCommand.ExecuteAsync(record);
            Assert.Null(opened);
            Assert.Contains("no longer available", model.ActionError);
        } finally { Directory.Delete(root, recursive: true); }
    }

    private sealed class BlockingRunner : IOfficeWorkflowRunner {
        private readonly OfficeWorkflowRunner _owner = new();
        private int _active;
        public int Started { get; private set; }
        public int Peak { get; private set; }
        public TaskCompletionSource TwoStarted { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource Release { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public async Task<OfficeWorkflowResult> RunAsync(OfficeWorkflowRequest request,
            IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default) {
            Started++;
            Peak = Math.Max(Peak, ++_active);
            if (Started == 2) TwoStarted.TrySetResult();
            try {
                await Release.Task.WaitAsync(cancellationToken);
                return await _owner.RunAsync(request, progress, cancellationToken);
            } finally { _active--; }
        }
        public Task<IReadOnlyList<OfficeWorkflowResult>> RunBatchAsync(IEnumerable<OfficeWorkflowRequest> requests,
            IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default) =>
            _owner.RunBatchAsync(requests, progress, cancellationToken);
    }
}
