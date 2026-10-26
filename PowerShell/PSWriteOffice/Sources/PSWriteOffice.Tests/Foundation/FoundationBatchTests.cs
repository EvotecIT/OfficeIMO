using System.Management.Automation;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Tests.Foundation;

public sealed class FoundationBatchTests {
    [Fact]
    public async Task ThrottleLimitsActualWorkersAndReportsEveryCompletion() {
        var pipeline = new RecordingPipeline();
        int active = 0, maximum = 0;
        await OfficeBatch.RunAsync(Enumerable.Range(1, 12), async (input, token) => {
            int count = Interlocked.Increment(ref active);
            int seen;
            do { seen = maximum; } while (count > seen && Interlocked.CompareExchange(ref maximum, count, seen) != seen);
            try { await Task.Delay(5, token); return (object?)input; }
            finally { Interlocked.Decrement(ref active); }
        }, 2, CancellationToken.None, pipeline, "Batch", "batch-failed");
        Assert.Equal(2, maximum);
        Assert.Equal(12, pipeline.Output.Count);
        Assert.Equal(13, pipeline.Progress.Count);
        Assert.Equal(ProgressRecordType.Completed, pipeline.Progress[^1].RecordType);
        Assert.Empty(pipeline.Errors);
    }
    [Fact]
    public async Task FailedItemsKeepTheirOriginalTargetAndDoNotAbortTheBatch() {
        var pipeline = new RecordingPipeline();
        await OfficeBatch.RunAsync(new[] { "first", "bad", "last" }, (input, _) => input == "bad"
            ? Task.FromException<object?>(new FileNotFoundException("Missing input.")) : Task.FromResult<object?>(input),
            2, CancellationToken.None, pipeline, "Batch", "input-unreadable");
        Assert.Equal(2, pipeline.Output.Count);
        var error = Assert.Single(pipeline.Errors);
        Assert.Equal("bad", error.TargetObject);
        Assert.Equal("input-unreadable", error.FullyQualifiedErrorId);
        Assert.Equal(ErrorCategory.ObjectNotFound, error.CategoryInfo.Category);
    }
    [Fact]
    public async Task CancellationJoinsAllWorkersBeforeReturning() {
        var pipeline = new RecordingPipeline();
        using var stop = new CancellationTokenSource();
        var entered = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        int active = 0;
        Task run = OfficeBatch.RunAsync(new[] { 1, 2, 3 }, async (input, token) => {
            Interlocked.Increment(ref active); entered.TrySetResult(true);
            try { await Task.Delay(Timeout.Infinite, token); return input; }
            finally { Interlocked.Decrement(ref active); }
        }, 2, stop.Token, pipeline, "Batch", "cancelled");
        await entered.Task.WaitAsync(TimeSpan.FromSeconds(10));
        stop.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => run);
        Assert.Equal(0, active);
        Assert.Empty(pipeline.Output);
        Assert.Empty(pipeline.Errors);
    }
    private sealed class RecordingPipeline : IAsyncCmdletPipeline {
        internal List<object?> Output { get; } = new();
        internal List<ProgressRecord> Progress { get; } = new();
        internal List<ErrorRecord> Errors { get; } = new();
        public void WriteObject(object? value) => Output.Add(value);
        public void WriteObject(object? value, bool enumerateCollection) => Output.Add(value);
        public void WriteError(ErrorRecord errorRecord) => Errors.Add(errorRecord);
        public void WriteProgress(ProgressRecord record) => Progress.Add(record);
        public void ThrowTerminatingError(ErrorRecord errorRecord) => throw new InvalidOperationException();
        public void WriteWarning(string message) => throw new InvalidOperationException();
        public void WriteVerbose(string message) => throw new InvalidOperationException();
        public void WriteDebug(string message) => throw new InvalidOperationException();
        public void WriteInformation(InformationRecord informationRecord) => throw new InvalidOperationException();
        public void WriteInformation(object messageData, string[] tags) => throw new InvalidOperationException();
        public bool ShouldProcess(string? target, string action) => throw new InvalidOperationException();
        public bool ShouldContinue(string query, string caption) => throw new InvalidOperationException();
        public PSCredential? PromptForCredential(string caption, string message, string userName, string targetName) => throw new InvalidOperationException();
    }
}
