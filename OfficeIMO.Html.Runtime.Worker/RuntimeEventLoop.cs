using AngleSharp;
using AngleSharp.Browser;
using AngleSharp.Common;
using AngleSharp.Js;
using Jint;
using Jint.Native;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

// Retain the provider's task transport, but finish promise jobs after every native
// task, including timers and lifecycle callbacks that do not pass through a command.
// The pinned DOM provider exposes mutation notification enqueue at its source,
// allowing the notification to enter the same FIFO queue as promise reactions.
internal sealed class RuntimeEventLoop(IBrowsingContext context, Func<Engine?> getEngine, RuntimeScriptErrors errors,
    object sessionSync, Action checkpoint, RuntimeScriptEntry scriptEntry) : IEventLoop, IMutationMicrotaskScheduler, IDisposable {
    private readonly IEventLoop _inner = new JsEventLoop(context);
    private readonly object _sync = sessionSync;
    private JsValue? _enqueueMicrotask;
    private volatile bool _cancelled;

    internal void InitializeMicrotasks(Engine engine) => _enqueueMicrotask=engine.Evaluate("(() => {const resolved=Promise.resolve(),then=Promise.prototype.then,apply=Reflect.apply;Object.defineProperty(resolved,'constructor',{value:Object.freeze({[Symbol.species]:Promise})});return callback=>apply(then,resolved,[callback]);})()");

    void IMutationMicrotaskScheduler.EnqueueMutationMicrotask(Action notification) => EnqueueMicrotask(notification);

    internal void EnqueueMicrotask(Action notification) {
        var engine=getEngine();
        if(engine==null || _enqueueMicrotask==null) {
            // An unscripted document can still be observed by another realm. Its
            // native notification must wait for the active script to release _sync.
            _inner.Enqueue(_=>{
                lock(_sync) {
                    if(!_cancelled) notification();
                }
            },TaskPriority.Microtask);
            return;
        }
        lock(_sync) {
            var callback=new ClrFunction(engine,"notifyMutations",(_,_)=>{
                if(!_cancelled) {
                    try {notification();}
                    catch(Exception error){errors.Report(error.Message);}
                }
                return JsValue.Undefined;
            });
            engine.Invoke(_enqueueMicrotask,new JsValue[]{callback});
        }
    }

    internal bool TryEnqueue(Action<CancellationToken> action, TaskPriority priority, CancellationToken lifetime) {
        lock (_sync) {
            if (_cancelled || lifetime.IsCancellationRequested) return false;
            _ = EnqueueCore(token => {
                if (!lifetime.IsCancellationRequested) action(token);
            }, priority);
            return true;
        }
    }

    public ICancellable Enqueue(Action<CancellationToken> action, TaskPriority priority) => EnqueueCore(action, priority);

    private ICancellable EnqueueCore(Action<CancellationToken> action, TaskPriority priority) => _inner.Enqueue(token => {
        if (token.IsCancellationRequested) return;
        lock (_sync) {
            if (_cancelled || token.IsCancellationRequested) return;
            using var entry = scriptEntry.Enter(context.Active);
            try { action(token); }
            finally {
                checkpoint();
                try {
                    Engine? engine = getEngine();
                    if (!_cancelled && engine != null) engine.Advanced.ProcessTasks();
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
    public void CancelAll() { lock (_sync) { _cancelled=true;_inner.CancelAll(); } }
    public void Dispose() { lock (_sync) { _cancelled=true;((IDisposable)_inner).Dispose(); } }
}
