using Microsoft.JSInterop;
using OfficeIMO.Web.Converter.Services;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class ConverterObjectUrlBatchTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public async Task WorkspaceChangeDuringAnyUrlCreationRevokesTheWholeUnpublishedResult(int completedUrls) {
        var js = new ControlledJs();
        var interop = new ConverterInterop(js);
        bool current = true;
        await using (var batch = new ConverterObjectUrlBatch(interop, () => current)) {
            for (int index = 0; index < completedUrls; index++)
                await batch.CreateAsync([1], "application/octet-stream");
            js.Pending = new(TaskCreationOptions.RunContinuationsAsynchronously);
            Task<string> pending = batch.CreateAsync([2], "application/json").AsTask();
            current = false;
            await interop.DisposeAsync();
            js.Pending.SetResult("blob:late");
            await Assert.ThrowsAsync<OperationCanceledException>(() => pending);
            Assert.Throws<OperationCanceledException>(() => batch.Commit());
        }
        Assert.Equal(Enumerable.Range(0, completedUrls).Select(index => $"blob:{index}").Append("blob:late"), js.Revoked);
    }

    [Fact]
    public async Task FailedCompanionCreationRevokesEarlierOutputWhileCommitRetainsDownloads() {
        var js = new ControlledJs();
        var interop = new ConverterInterop(js);
        await using (var batch = new ConverterObjectUrlBatch(interop, () => true)) {
            await batch.CreateAsync([1], "application/pdf");
            js.Pending = new(TaskCreationOptions.RunContinuationsAsynchronously);
            js.Pending.SetException(new JSException("Creation failed"));
            await Assert.ThrowsAsync<JSException>(() => batch.CreateAsync([2], "application/json").AsTask());
        }
        Assert.Equal(["blob:0"], js.Revoked);
        js.Pending = null;
        await using (var batch = new ConverterObjectUrlBatch(interop, () => true)) {
            await batch.CreateAsync([1], "application/pdf");
            await batch.CreateAsync([2], "application/json");
            batch.Commit();
        }
        Assert.Equal(["blob:0"], js.Revoked);
    }

    private sealed class ControlledJs : IJSRuntime, IJSObjectReference {
        internal TaskCompletionSource<string>? Pending;
        internal List<string> Revoked { get; } = [];
        private int _next;
        public ValueTask<TValue> InvokeAsync<TValue>(string identifier, object?[]? args) => InvokeAsync<TValue>(identifier, default, args);
        public async ValueTask<TValue> InvokeAsync<TValue>(string identifier, CancellationToken cancellationToken, object?[]? args) {
            if (identifier == "import") return (TValue)(object)this;
            if (identifier == ConverterInterop.CreateObjectUrlMethod)
                return (TValue)(object)(Pending is null ? $"blob:{_next++}" : await Pending.Task);
            if (identifier == "URL.revokeObjectURL") { Revoked.Add((string)args![0]!); return default!; }
            throw new InvalidOperationException(identifier);
        }
        public ValueTask DisposeAsync() => ValueTask.CompletedTask;
    }
}
