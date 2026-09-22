using AngleSharp.Dom;
using Jint;
using Jint.Native;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

// The receiver owns the native timer; the registering realm owns its callback.
internal sealed class RuntimeTimerLifetime(Engine engine, IWindow caller) : IDisposable {
    private readonly HashSet<(IWindow Window, int Handle)> _timers = [];
    private bool _disposed;

    internal JsValue Alive => new ClrFunction(engine, "timerRealmAlive", (_, _) => !_disposed);
    internal JsValue Track => new ClrFunction(engine, "trackTimer", (_, args) => {
        var timer = Key(args);
        if (_disposed) timer.Window.ClearTimeout(timer.Handle);
        else _timers.Add(timer);
        return JsValue.Undefined;
    });
    internal JsValue Release => new ClrFunction(engine, "releaseTimer", (_, args) => {
        _timers.Remove(Key(args));
        return JsValue.Undefined;
    });

    private (IWindow Window, int Handle) Key(JsValue[] args) =>
        ((args[0].IsNull() || args[0].IsUndefined() || ReferenceEquals(args[0], engine.Global)) ? caller : (IWindow)args[0].ToObject()!, (int)args[1].AsNumber());

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        foreach (var timer in _timers) timer.Window.ClearTimeout(timer.Handle);
        _timers.Clear();
    }
}
