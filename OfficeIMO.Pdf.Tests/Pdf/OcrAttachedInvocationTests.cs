using System;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class OcrAttachedInvocationTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task AttachedInvocationWaitsForProviderCleanupAfterCancellationOrDeadline(bool concurrent, bool deadline) {
        var entered = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var cleaning = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var cancellation = new CancellationTokenSource();
        var engine = new DelegateOcrEngine("attached-fixture", async (_, token) => {
            entered.TrySetResult(true);
            try { await Task.Delay(Timeout.Infinite, token); }
            finally { cleaning.TrySetResult(true); await release.Task; }
            return new OcrResult();
        }, new OcrEngineCapabilities { SupportsConcurrentRequests = concurrent });
        Task<OcrResult> operation = OcrEngineRunner.CreateExecution(engine).RecognizeAttachedAsync(new(),
            deadline ? TimeSpan.FromMilliseconds(250) : TimeSpan.FromSeconds(30), new(1, 1, 0), cancellation.Token);
        try {
            Assert.Same(entered.Task, await Task.WhenAny(entered.Task, Task.Delay(TimeSpan.FromSeconds(10)))); await entered.Task;
            if (!deadline) cancellation.Cancel();
            Assert.Same(cleaning.Task, await Task.WhenAny(cleaning.Task, Task.Delay(TimeSpan.FromSeconds(10)))); await cleaning.Task;
            await Task.Delay(100);
            Assert.False(operation.IsCompleted);
            release.TrySetResult(true);
            if (deadline) await Assert.ThrowsAsync<OcrEngineTimeoutException>(() => operation);
            else await Assert.ThrowsAnyAsync<OperationCanceledException>(() => operation);
        } finally {
            release.TrySetResult(true); cancellation.Cancel();
            try { await operation; } catch (OperationCanceledException) { } catch (OcrEngineTimeoutException) { }
        }
    }

    [Fact]
    public async Task AttachedInvocationAlsoWaitsForCancellationCallbacksAfterTheProviderReturns() {
        var entered = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var callbackEntered = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var returnProvider = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var providerReturned = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        using var releaseCallback = new ManualResetEventSlim();
        using var cancellation = new CancellationTokenSource();
        var engine = new DelegateOcrEngine("attached-callback", async (_, token) => {
            // This registration belongs to the runner-owned invocation token and is released with that scope.
            token.Register(() => { callbackEntered.TrySetResult(true); releaseCallback.Wait(); });
            entered.TrySetResult(true); await returnProvider.Task; providerReturned.TrySetResult(true);
            return new OcrResult();
        });
        var operation = OcrEngineRunner.CreateExecution(engine).RecognizeAttachedAsync(new(), TimeSpan.FromSeconds(30), new(1, 1, 0), cancellation.Token);
        try {
            Assert.Same(entered.Task, await Task.WhenAny(entered.Task, Task.Delay(TimeSpan.FromSeconds(10)))); await entered.Task; cancellation.Cancel();
            Assert.Same(callbackEntered.Task, await Task.WhenAny(callbackEntered.Task, Task.Delay(TimeSpan.FromSeconds(10)))); await callbackEntered.Task;
            returnProvider.TrySetResult(true); Assert.Same(providerReturned.Task, await Task.WhenAny(providerReturned.Task, Task.Delay(TimeSpan.FromSeconds(10)))); await providerReturned.Task;
            await Task.Delay(100); Assert.False(operation.IsCompleted);
            releaseCallback.Set(); await Assert.ThrowsAnyAsync<OperationCanceledException>(() => operation);
        } finally {
            returnProvider.TrySetResult(true); releaseCallback.Set(); cancellation.Cancel();
            try { await operation; } catch (OperationCanceledException) { }
        }
    }

    [Fact]
    public async Task AttachedCancellationWhileWaitingForExistingEngineGateDoesNotStartAnotherProvider() {
        int calls = 0;
        var entered = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var engine = new DelegateOcrEngine("attached-admission", async (_, _) => {
            Interlocked.Increment(ref calls); entered.TrySetResult(true); await release.Task; return new OcrResult();
        });
        var execution = OcrEngineRunner.CreateExecution(engine);
        Task<OcrResult> first = execution.RecognizeAsync(new(), TimeSpan.FromSeconds(30));
        try {
            Assert.Same(entered.Task, await Task.WhenAny(entered.Task, Task.Delay(TimeSpan.FromSeconds(10)))); await entered.Task;
            using var cancellation = new CancellationTokenSource();
            Task<OcrResult> waiting = execution.RecognizeAttachedAsync(new(), TimeSpan.FromSeconds(30), new(1, 1, 0), cancellation.Token);
            cancellation.Cancel();
            Assert.Same(waiting, await Task.WhenAny(waiting, Task.Delay(TimeSpan.FromSeconds(5))));
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => waiting);
            Assert.Equal(1, Volatile.Read(ref calls));
        } finally { release.TrySetResult(true); await first; }
    }
}
