using AngleSharp;
using AngleSharp.Browser;
using AngleSharp.Common;
using AngleSharp.Js;
using Jint;

namespace OfficeIMO.Html.Runtime.Worker;

// Retain the provider's task transport, but finish promise jobs after every native
// task, including timers and lifecycle callbacks that do not pass through a command.
// MutationHost still queues its notification as a separate normal task; this adapter
// does not change that provider behavior or claim a unified mutation microtask queue.
internal sealed class RuntimeEventLoop(IBrowsingContext context, Func<Engine?> getEngine, RuntimeScriptErrors errors) : IEventLoop, IDisposable {
    private readonly IEventLoop _inner = new JsEventLoop(context);

    public ICancellable Enqueue(Action<CancellationToken> action, TaskPriority priority) => _inner.Enqueue(token => {
        if (token.IsCancellationRequested) return;
        var engine = getEngine();
        if (engine == null) { action(token); return; }
        lock (engine) {
            try { action(token); }
            finally {
                try {
                    engine.Advanced.ProcessTasks();
                    errors.ThrowIfFailed();
                } catch (Exception error) {
                    // Latch failures even while no host command is active. Do not throw
                    // through provider completion callbacks and strand document loading.
                    errors.Report(error.Message);
                }
            }
        }
    }, priority);

    public void Spin() => _inner.Spin();
    public void CancelAll() => _inner.CancelAll();
    public void Dispose() => ((IDisposable)_inner).Dispose();
}
