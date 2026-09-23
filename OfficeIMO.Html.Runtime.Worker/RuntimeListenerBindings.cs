using System.Runtime.CompilerServices;
using AngleSharp.Dom;
using AngleSharp.Dom.Events;
using Jint;
using Jint.Native;
using Jint.Native.Function;
using Jint.Runtime;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

// The provider creates a new CLR delegate each time a JS callback crosses its method
// boundary. Own registration identity so removal and duplicate suppression are reliable.
internal sealed class RuntimeListenerBindings : IDisposable {
    private readonly ConditionalWeakTable<IEventTarget, TargetListeners> _targets = new();
    private readonly Engine _engine;
    private readonly IEventTarget _window;
    private bool _disposed;

    internal RuntimeListenerBindings(Engine engine, IEventTarget window) { _engine = engine; _window = window; }

    internal JsValue Add => new ClrFunction(_engine, "addEventListener", (receiver, args) => {
        if (_disposed) return JsValue.Undefined;
        var target = Target(receiver);
        string type = TypeConverter.ToString(args.ElementAtOrDefault(0) ?? JsValue.Undefined);
        if (args.ElementAtOrDefault(1) is not Function callback) return JsValue.Undefined;
        var options = args.ElementAtOrDefault(2) ?? JsValue.Undefined;
        bool capture = Capture(options);
        bool once = false;
        bool passive = false;
        if (options.IsObject()) {
            var value = options.AsObject();
            once = TypeConverter.ToBoolean(value.Get("once"));
            passive = TypeConverter.ToBoolean(value.Get("passive"));
            var signal = value.Get("signal");
            if (!signal.IsNull() && !signal.IsUndefined())
                throw new HtmlScriptRuntimeException("Signal-controlled event listeners are not supported by this runtime profile.");
        }
        var registrations = _targets.GetValue(target, CreateListeners).Registrations;
        registrations.RemoveAll(item => item.ResetVersion != ResetVersion(target));
        if (registrations.Any(item => item.Type == type && item.Capture == capture && ReferenceEquals(item.Callback, callback))) return JsValue.Undefined;
        var registration = new Registration(type, capture, callback, ResetVersion(target));
        registration.Handler = (sender, ev) => {
            if (_disposed) return;
            if (once) Remove(target, registrations, registration);
            if (passive) {
                using var scope = ev.BeginPassiveListener();
                _engine.Invoke(callback, JsValue.FromObject(_engine, sender), new[] { JsValue.FromObject(_engine, ev) });
            } else {
                _engine.Invoke(callback, JsValue.FromObject(_engine, sender), new[] { JsValue.FromObject(_engine, ev) });
            }
        };
        registrations.Add(registration);
        target.AddEventListener(type, registration.Handler, capture);
        return JsValue.Undefined;
    });

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        foreach (var pair in _targets) {
            if (pair.Key is EventTarget native) native.OnReset -= pair.Value.OnReset;
            foreach (var registration in pair.Value.Registrations)
                pair.Key.RemoveEventListener(registration.Type, registration.Handler, registration.Capture);
        }
        _targets.Clear();
    }

    internal JsValue RemoveListener => new ClrFunction(_engine, "removeEventListener", (receiver, args) => {
        var target = Target(receiver);
        string type = TypeConverter.ToString(args.ElementAtOrDefault(0) ?? JsValue.Undefined);
        var callback = args.ElementAtOrDefault(1);
        bool capture = Capture(args.ElementAtOrDefault(2) ?? JsValue.Undefined);
        if (_targets.TryGetValue(target, out var listeners)) {
            listeners.Registrations.RemoveAll(item => item.ResetVersion != ResetVersion(target));
            var registration = listeners.Registrations.FirstOrDefault(item => item.Type == type && item.Capture == capture && ReferenceEquals(item.Callback, callback));
            if (registration != null) Remove(target, listeners.Registrations, registration);
        }
        return JsValue.Undefined;
    });

    private IEventTarget Target(JsValue receiver) => ReferenceEquals(receiver, _engine.Global) ? _window : receiver.ToObject() as IEventTarget
        ?? throw new HtmlScriptRuntimeException("The event listener receiver must be a DOM event target.");

    private static bool Capture(JsValue options) => TypeConverter.ToBoolean(options.IsObject() ? options.AsObject().Get("capture") : options);

    private static long ResetVersion(IEventTarget target) => (target as EventTarget)?.ListenerResetVersion ?? 0;

    private static void Remove(IEventTarget target, List<Registration> registrations, Registration registration) {
        target.RemoveEventListener(registration.Type, registration.Handler, registration.Capture);
        registrations.Remove(registration);
    }

    private static TargetListeners CreateListeners(IEventTarget target) {
        var state = new TargetListeners();
        if (target is EventTarget native) native.OnReset += state.OnReset;
        return state;
    }

    private sealed class TargetListeners {
        internal readonly List<Registration> Registrations = new();

        internal void OnReset(object? sender, EventArgs args) {
            if (sender is EventTarget target)
                Registrations.RemoveAll(item => item.ResetVersion != target.ListenerResetVersion);
        }
    }

    private sealed class Registration(string type, bool capture, Function callback, long resetVersion) {
        internal string Type { get; } = type;
        internal bool Capture { get; } = capture;
        internal Function Callback { get; } = callback;
        internal long ResetVersion { get; } = resetVersion;
        internal DomEventHandler Handler { get; set; } = null!;
    }
}
