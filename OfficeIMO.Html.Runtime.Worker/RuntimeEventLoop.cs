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
internal sealed class RuntimeEventLoop(IBrowsingContext context, Func<Engine?> getEngine, RuntimeScriptErrors errors) : IEventLoop, IMutationMicrotaskScheduler, IDisposable {
    private readonly IEventLoop _inner = new JsEventLoop(context);
    private JsValue? _enqueueMicrotask;
    private volatile bool _cancelled;

    internal void InitializeMicrotasks(Engine engine) => _enqueueMicrotask=engine.Evaluate("(() => {const resolved=Promise.resolve(),then=Promise.prototype.then,apply=Reflect.apply;Object.defineProperty(resolved,'constructor',{value:Object.freeze({[Symbol.species]:Promise})});return callback=>apply(then,resolved,[callback]);})()");

    void IMutationMicrotaskScheduler.EnqueueMutationMicrotask(Action notification) => EnqueueMicrotask(notification);

    internal void EnqueueMicrotask(Action notification) {
        var engine=getEngine();
        if(engine==null || _enqueueMicrotask==null) {
            _inner.Enqueue(_=>{if(!_cancelled)notification();},TaskPriority.Microtask);
            return;
        }
        lock(engine) {
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
    public void CancelAll() { _cancelled=true;_inner.CancelAll(); }
    public void Dispose() { _cancelled=true;((IDisposable)_inner).Dispose(); }
}
