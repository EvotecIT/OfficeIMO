using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;

namespace OfficeIMO.Html;

/// <summary>
/// Supplies the stack that <see cref="HtmlRenderOptions.MaxLayoutDepth"/> permits. A stack overflow
/// ends the process and no caller can catch it, and worker threads can start with as little as 512 KB.
/// </summary>
internal static class HtmlExecutionStack {
    // Each layout level costs up to about 30 KB of unoptimized frames.
    private const int StackBytesPerLayoutLevel = 64 * 1024;
    internal const int MinimumStackBytes = 16 * 1024 * 1024;
    private const int MaximumStackBytes = 1024 * 1024 * 1024;

    internal static int StackBytesFor(int layoutDepth) =>
        (int)Math.Min(MaximumStackBytes, Math.Max(MinimumStackBytes, (long)layoutDepth * StackBytesPerLayoutLevel));

    internal static bool HasRoom() {
#if NET8_0_OR_GREATER
        return RuntimeHelpers.TryEnsureSufficientExecutionStack();
#else
        try {
            RuntimeHelpers.EnsureSufficientExecutionStack();
            return true;
        } catch (InsufficientExecutionStackException) {
            return false;
        }
#endif
    }

    /// <summary>Runs <paramref name="work"/> on a dedicated stack while the caller waits, so execution stays sequential.</summary>
    /// <returns><see langword="false"/> without running the work when the host cannot start a thread.</returns>
    internal static bool TryRun<T>(Func<T> work, int stackBytes, out T result) {
        if (!StackWorker.TryRent(stackBytes, out StackWorker? worker)) {
            result = default!;
            return false;
        }
        T completed = default!;
        ExceptionDispatchInfo? failure = null;
        // The work sees the caller's cultures and flowed context, as it would on a new thread.
        CultureInfo culture = CultureInfo.CurrentCulture;
        CultureInfo uiCulture = CultureInfo.CurrentUICulture;
        ExecutionContext? context = CaptureWithoutSynchronizationContext();
        worker!.Run(() => {
            try {
                Thread current = Thread.CurrentThread;
                current.CurrentCulture = culture;
                current.CurrentUICulture = uiCulture;
                if (context == null) {
                    completed = work();
                } else {
                    ExecutionContext.Run(context, static state => {
                        var job = (Action)state!;
                        job();
                    }, (Action)(() => completed = work()));
                }
            } catch (Exception exception) {
                failure = ExceptionDispatchInfo.Capture(exception);
            }
        });
        StackWorker.Return(worker);
        failure?.Throw();
        result = completed;
        return true;
    }

    /// <summary>
    /// .NET Framework captures the caller's <see cref="SynchronizationContext"/> with the flowed context,
    /// and a new thread never carried it. The caller blocks while layout runs, so work posted back to it
    /// would deadlock; capture the flowed values without it, as <see cref="Thread.Start()"/> does.
    /// </summary>
    private static ExecutionContext? CaptureWithoutSynchronizationContext() {
        SynchronizationContext? callerContext = SynchronizationContext.Current;
        if (callerContext == null) return ExecutionContext.Capture();
        SynchronizationContext.SetSynchronizationContext(null);
        try {
            return ExecutionContext.Capture();
        } finally {
            SynchronizationContext.SetSynchronizationContext(callerContext);
        }
    }

    /// <summary>
    /// A rendering thread kept for the next conversion. Starting a thread with a large stack for
    /// every document costs more than many small conversions spend on layout, so a few idle
    /// workers wait briefly for further work before they exit.
    /// </summary>
    private sealed class StackWorker {
        private const int MaximumIdleWorkers = 4;
        private const int MaximumRetainedStackBytes = 4 * MinimumStackBytes;
        private const int IdleLifetimeMilliseconds = 30_000;
        private static readonly object Gate = new object();
        private static readonly List<StackWorker> Idle = new List<StackWorker>();

        private readonly int _stackBytes;
        private readonly AutoResetEvent _start = new AutoResetEvent(false);
        private readonly AutoResetEvent _completed = new AutoResetEvent(false);
        private Action? _job;
        private bool _stopping;

        private StackWorker(int stackBytes) => _stackBytes = stackBytes;

        internal static bool TryRent(int stackBytes, out StackWorker? worker) {
            lock (Gate) {
                for (int index = Idle.Count - 1; index >= 0; index--) {
                    if (Idle[index]._stackBytes != stackBytes) continue;
                    worker = Idle[index];
                    Idle.RemoveAt(index);
                    return true;
                }
            }
            var created = new StackWorker(stackBytes);
            var thread = new Thread(created.Loop, stackBytes) {
                IsBackground = true,
                Name = "OfficeIMO HTML rendering"
            };
            try {
                // A worker serves later callers, so it must not keep the first caller's context.
                // Each job applies its own caller's context instead.
                if (ExecutionContext.IsFlowSuppressed()) {
                    thread.Start();
                } else {
                    using (ExecutionContext.SuppressFlow()) thread.Start();
                }
            } catch (Exception exception) when (exception is PlatformNotSupportedException or ThreadStartException or OutOfMemoryException) {
                created.Dispose();
                worker = null;
                return false;
            }
            worker = created;
            return true;
        }

        internal static void Return(StackWorker worker) {
            if (worker._stackBytes <= MaximumRetainedStackBytes) {
                lock (Gate) {
                    if (Idle.Count < MaximumIdleWorkers) {
                        Idle.Add(worker);
                        return;
                    }
                }
            }
            worker.Stop();
        }

        internal void Run(Action job) {
            _job = job;
            _start.Set();
            _completed.WaitOne();
        }

        private void Stop() {
            _stopping = true;
            _start.Set();
        }

        private void Loop() {
            try {
                while (true) {
                    if (!_start.WaitOne(IdleLifetimeMilliseconds)) {
                        lock (Gate) {
                            if (Idle.Remove(this)) return;
                        }
                        // A caller rented this worker as it timed out; its job is being handed over.
                        _start.WaitOne();
                    }
                    if (_stopping) return;
                    RunPendingJob();
                    _completed.Set();
                }
            } finally {
                Dispose();
            }
        }

        // The job's closure holds the rendered result. Keep it in this frame only, so an idle
        // worker does not retain the previous document while it waits.
        [MethodImpl(MethodImplOptions.NoInlining)]
        private void RunPendingJob() {
            Action job = _job!;
            _job = null;
            job();
        }

        private void Dispose() {
            _start.Dispose();
            _completed.Dispose();
        }
    }
}
