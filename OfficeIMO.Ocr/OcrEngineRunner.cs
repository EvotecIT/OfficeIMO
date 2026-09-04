using System;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Ocr;

/// <summary>Coordinates bounded calls to a shared OCR engine instance.</summary>
public static class OcrEngineRunner {
    /// <summary>Maximum accepted length of an OCR engine identifier before surrounding whitespace is removed.</summary>
    public const int MaximumEngineIdCharacters = 256;

    private static readonly ConditionalWeakTable<IOcrEngine, SemaphoreSlim> NonConcurrentEngineGates =
        new ConditionalWeakTable<IOcrEngine, SemaphoreSlim>();

    /// <summary>Reads and validates the stable identifier exposed by an OCR engine.</summary>
    public static string GetValidatedEngineId(IOcrEngine engine) {
        if (engine == null) throw new ArgumentNullException(nameof(engine));
        string? rawEngineId = engine.Id;
        if (string.IsNullOrEmpty(rawEngineId)) throw new ArgumentException("OCR engine id cannot be empty.", nameof(engine));
        if (rawEngineId.Length > MaximumEngineIdCharacters) {
            throw new ArgumentException(
                "OCR engine id cannot exceed " + MaximumEngineIdCharacters + " characters.",
                nameof(engine));
        }

        string engineId = rawEngineId.Trim();
        if (engineId.Length == 0) throw new ArgumentException("OCR engine id cannot be empty.", nameof(engine));
        return engineId;
    }

    /// <summary>Captures one validated identity and capability snapshot for a logical OCR operation.</summary>
    public static OcrEngineExecution CreateExecution(IOcrEngine engine) {
        if (engine == null) throw new ArgumentNullException(nameof(engine));
        string engineId = GetValidatedEngineId(engine);
        OcrEngineCapabilities capabilities = (engine.Capabilities ?? new OcrEngineCapabilities()).Clone();
        return new OcrEngineExecution(engine, engineId, capabilities);
    }

    /// <summary>
    /// Runs one recognition request within a total timeout. Engines that do not advertise concurrent-request
    /// support are serialized across every Reader, PDF, and caller integration using this method.
    /// </summary>
    /// <remarks>
    /// If a provider ignores cancellation after a timeout, its shared gate remains held until that provider call
    /// actually settles. A later caller can still time out while waiting for the occupied gate.
    /// </remarks>
    public static Task<OcrResult> RecognizeAsync(
        IOcrEngine engine,
        OcrRequest request,
        TimeSpan timeout,
        CancellationToken cancellationToken = default) =>
        CreateExecution(engine).RecognizeAsync(request, timeout, cancellationToken);

    internal static async Task<OcrResult> RecognizeAsync(
        OcrEngineExecution execution,
        OcrRequest request,
        TimeSpan timeout,
        CancellationToken cancellationToken) {
        if (execution == null) throw new ArgumentNullException(nameof(execution));
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (timeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(timeout));
        IOcrEngine engine = execution.Engine;
        string engineId = execution.Id;

        SemaphoreSlim? gate = execution.SupportsConcurrentRequests
            ? null
            : NonConcurrentEngineGates.GetValue(engine, static _ => new SemaphoreSlim(1, 1));
        bool gateHeld = false;
        CancellationTokenSource? providerCancellation = null;
        ProviderInvocation? providerInvocation = null;
        Stopwatch elapsed = Stopwatch.StartNew();

        try {
            if (gate != null) {
                gateHeld = await WaitForGateAsync(gate, timeout, elapsed, engineId, cancellationToken)
                    .ConfigureAwait(false);
            }

            cancellationToken.ThrowIfCancellationRequested();
            TimeSpan remaining = GetRemaining(timeout, elapsed, engineId, providerCallStarted: false);
            providerCancellation = new CancellationTokenSource();
            using var deadlineCancellation = new CancellationTokenSource();
            Task deadlineTask = DelayLongAsync(remaining, deadlineCancellation.Token);
            providerInvocation = StartProviderCall(
                engine,
                request,
                providerCancellation.Token,
                elapsed,
                timeout,
                cancellationToken);
            OcrResult result = await WaitForProviderAsync(
                    providerInvocation,
                    deadlineTask,
                    deadlineCancellation,
                    timeout,
                    engineId,
                    providerCancellation,
                    cancellationToken)
                .ConfigureAwait(false);
            return result ?? throw new InvalidOperationException("OCR engine returned a null result.");
        } finally {
            Task<OcrResult>? providerTask = providerInvocation?.Task;
            if (providerTask != null) ObserveBackgroundFailure(providerTask);
            DisposeCancellationSourceWhenSettled(
                providerCancellation,
                providerTask,
                providerInvocation?.CancellationTask);
            if (gateHeld && gate != null) {
                ReleaseGateWhenSettled(gate, providerTask, providerInvocation?.CancellationTask);
            }
        }
    }

    private static async Task<bool> WaitForGateAsync(
        SemaphoreSlim gate,
        TimeSpan timeout,
        Stopwatch elapsed,
        string engineId,
        CancellationToken cancellationToken) {
        TimeSpan remaining = GetRemaining(timeout, elapsed, engineId, providerCallStarted: false);
        using var gateCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        using var deadlineCancellation = new CancellationTokenSource();
        Task gateTask = gate.WaitAsync(gateCancellation.Token);
        Task deadlineTask = DelayLongAsync(remaining, deadlineCancellation.Token);
        Task completed = await WaitForCompletionAsync(gateTask, deadlineTask, cancellationToken).ConfigureAwait(false);

        if (completed == gateTask) {
            deadlineCancellation.Cancel();
            await gateTask.ConfigureAwait(false);
            return true;
        }

        deadlineCancellation.Cancel();
        TryCancel(gateCancellation);
        bool acquired = false;
        try {
            await gateTask.ConfigureAwait(false);
            acquired = true;
        } catch (OperationCanceledException) {
        }

        if (acquired) gate.Release();
        cancellationToken.ThrowIfCancellationRequested();
        throw new OcrEngineTimeoutException(engineId, timeout, providerCallStarted: false);
    }

    private static ProviderInvocation StartProviderCall(
        IOcrEngine engine,
        OcrRequest request,
        CancellationToken providerCancellationToken,
        Stopwatch elapsed,
        TimeSpan timeout,
        CancellationToken callerCancellationToken) {
        var invocation = new ProviderInvocation(elapsed, timeout, callerCancellationToken);
        invocation.Start(engine, request, providerCancellationToken);
        return invocation;
    }

    private static async Task<OcrResult> WaitForProviderAsync(
        ProviderInvocation invocation,
        Task deadlineTask,
        CancellationTokenSource deadlineCancellation,
        TimeSpan timeout,
        string engineId,
        CancellationTokenSource providerCancellation,
        CancellationToken cancellationToken) {
        Task<OcrResult> providerTask = invocation.Task;
        if (providerTask.IsCompleted) {
            if (!invocation.HasStarted) ThrowSuppressedProvider(engineId, timeout, cancellationToken);
            deadlineCancellation.Cancel();
            return await providerTask.ConfigureAwait(false);
        }

        Task completed = await WaitForCompletionAsync(providerTask, deadlineTask, cancellationToken).ConfigureAwait(false);
        if (completed == providerTask || providerTask.IsCompleted) {
            if (!invocation.HasStarted) ThrowSuppressedProvider(engineId, timeout, cancellationToken);
            deadlineCancellation.Cancel();
            return await providerTask.ConfigureAwait(false);
        }

        invocation.SuppressIfNotStarted();
        invocation.RequestCancellation(providerCancellation);
        deadlineCancellation.Cancel();
        if (cancellationToken.IsCancellationRequested) {
            cancellationToken.ThrowIfCancellationRequested();
        }

        throw new OcrEngineTimeoutException(engineId, timeout, invocation.HasStarted);
    }

    private static void ThrowSuppressedProvider(
        string engineId,
        TimeSpan timeout,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        throw new OcrEngineTimeoutException(engineId, timeout, providerCallStarted: false);
    }

    private static TimeSpan GetRemaining(TimeSpan timeout, Stopwatch elapsed, string engineId, bool providerCallStarted) {
        TimeSpan remaining = timeout - elapsed.Elapsed;
        if (remaining > TimeSpan.Zero) return remaining;
        throw new OcrEngineTimeoutException(engineId, timeout, providerCallStarted);
    }

    private static async Task DelayLongAsync(TimeSpan delay, CancellationToken cancellationToken) {
        TimeSpan remaining = delay;
        TimeSpan maximumSlice = TimeSpan.FromDays(20);
        while (remaining > TimeSpan.Zero) {
            TimeSpan slice = remaining > maximumSlice ? maximumSlice : remaining;
            await Task.Delay(slice, cancellationToken).ConfigureAwait(false);
            remaining -= slice;
        }
    }

    private static async Task<Task> WaitForCompletionAsync(
        Task operation,
        Task deadline,
        CancellationToken cancellationToken) {
        if (!cancellationToken.CanBeCanceled) {
            return await Task.WhenAny(operation, deadline).ConfigureAwait(false);
        }

        var canceled = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        using (cancellationToken.Register(static state => ((TaskCompletionSource<object?>)state!).TrySetResult(null), canceled)) {
            return await Task.WhenAny(operation, deadline, canceled.Task).ConfigureAwait(false);
        }
    }

    private static void TryCancel(CancellationTokenSource cancellation) {
        try {
            cancellation.Cancel();
        } catch (AggregateException) {
            // Provider cancellation callbacks must not replace the timeout result.
        } catch (ObjectDisposedException) {
        }
    }

    private static void DisposeCancellationSourceWhenSettled(
        CancellationTokenSource? cancellation,
        Task? providerTask,
        Task? cancellationTask) {
        if (cancellation == null) return;
        if ((providerTask == null || providerTask.IsCompleted) &&
            (cancellationTask == null || cancellationTask.IsCompleted)) {
            if (cancellationTask?.IsFaulted == true) _ = cancellationTask.Exception;
            cancellation.Dispose();
            return;
        }

        Task[] pending = cancellationTask == null
            ? new[] { providerTask! }
            : new[] { providerTask!, cancellationTask };
        _ = Task.WhenAll(pending).ContinueWith(
            static (completed, state) => {
                _ = completed.Exception;
                ((CancellationTokenSource)state!).Dispose();
            },
            cancellation,
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);
    }

    private static void ReleaseGateWhenSettled(
        SemaphoreSlim gate,
        Task? providerTask,
        Task? cancellationTask) {
        if ((providerTask == null || providerTask.IsCompleted) &&
            (cancellationTask == null || cancellationTask.IsCompleted)) {
            gate.Release();
            return;
        }

        Task settlement = cancellationTask == null
            ? providerTask!
            : Task.WhenAll(providerTask!, cancellationTask);
        _ = settlement.ContinueWith(
            static (completed, state) => {
                _ = completed.Exception;
                ((SemaphoreSlim)state!).Release();
            },
            gate,
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);
    }

    private static void ObserveBackgroundFailure(Task providerTask) {
        _ = providerTask.ContinueWith(
            static completed => { _ = completed.Exception; },
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously | TaskContinuationOptions.OnlyOnFaulted,
            TaskScheduler.Default);
    }

    private sealed class ProviderInvocation {
        private readonly OcrProviderEntryGate _entryGate;
        private int _cancellationRequested;
        private Task? _cancellationTask;

        internal ProviderInvocation(
            Stopwatch elapsed,
            TimeSpan timeout,
            CancellationToken callerCancellationToken) {
            _entryGate = new OcrProviderEntryGate(elapsed, timeout, callerCancellationToken);
        }

        internal Task<OcrResult> Task { get; private set; } = null!;

        internal Task? CancellationTask => Volatile.Read(ref _cancellationTask);

        internal bool HasStarted => _entryGate.HasStarted;

        internal void Start(IOcrEngine engine, OcrRequest request, CancellationToken cancellationToken) {
            Task = TaskFactory(engine, request, cancellationToken);
        }

        internal void SuppressIfNotStarted() {
            _entryGate.SuppressIfNotStarted();
        }

        internal void RequestCancellation(CancellationTokenSource cancellation) {
            if (Interlocked.CompareExchange(ref _cancellationRequested, 1, 0) != 0) return;
            System.Threading.Tasks.Task cancellationTask = System.Threading.Tasks.Task.Factory.StartNew(
                static state => TryCancel((CancellationTokenSource)state!),
                cancellation,
                CancellationToken.None,
                TaskCreationOptions.LongRunning | TaskCreationOptions.DenyChildAttach,
                TaskScheduler.Default);
            Volatile.Write(ref _cancellationTask, cancellationTask);
        }

        private Task<OcrResult> TaskFactory(
            IOcrEngine engine,
            OcrRequest request,
            CancellationToken cancellationToken) => System.Threading.Tasks.Task.Factory.StartNew(
                () => {
                    if (!_entryGate.TryStart()) {
                        throw new OperationCanceledException("OCR provider invocation was suppressed before it started.");
                    }
                    return engine.RecognizeAsync(request, cancellationToken);
                },
                CancellationToken.None,
                TaskCreationOptions.LongRunning | TaskCreationOptions.DenyChildAttach,
                TaskScheduler.Default).Unwrap();
    }
}

/// <summary>Thrown when an OCR engine cannot complete within the caller's total execution timeout.</summary>
public sealed class OcrEngineTimeoutException : TimeoutException {
    internal OcrEngineTimeoutException(string engineId, TimeSpan timeout, bool providerCallStarted)
        : base("OCR engine '" + engineId + "' exceeded its execution timeout (" + timeout + ").") {
        EngineId = engineId;
        Timeout = timeout;
        ProviderCallStarted = providerCallStarted;
    }

    /// <summary>Identifier of the engine whose execution timed out.</summary>
    public string EngineId { get; }

    /// <summary>Total timeout applied to the execution.</summary>
    public TimeSpan Timeout { get; }

    /// <summary>Whether the provider method had started before the timeout expired.</summary>
    public bool ProviderCallStarted { get; }
}
