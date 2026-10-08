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
    private readonly RuntimeEventTargets _targetsMap;
    private bool _disposed;

    internal RuntimeListenerBindings(Engine engine, RuntimeEventTargets targets) { _engine = engine; _targetsMap = targets; }

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
        if (registrations.Any(item => item.Type == type && item.Capture == capture && ReferenceEquals(item.Callback, callback))) return JsValue.Undefined;
        if (registrations.Count >= RuntimeEventTargets.ListenerLimit(target)) {
            var error = _engine.Intrinsics.Error.Construct("The event target listener limit was exceeded.");
            error.Set("name", "QuotaExceededError");
            throw new JavaScriptException(error);
        }
        var registration = new Registration(type, capture, callback);
        registration.Handler = (sender, ev) => {
            if (_disposed) return;
            if (once) Remove(target, registrations, registration);
            try {
                if (passive) {
                    using var scope = ev.BeginPassiveListener();
                    _engine.Invoke(callback, _targetsMap.Wrap(sender), new[] { JsValue.FromObject(_engine, ev) });
                } else {
                    _engine.Invoke(callback, _targetsMap.Wrap(sender), new[] { JsValue.FromObject(_engine, ev) });
                }
            } catch (JavaScriptException) when (RuntimeEventTargets.IsTransport(target)) {
                // The shared wrapper reported the failure. Transport dispatch must
                // still invoke later listeners and finish the request lifecycle.
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
            var registration = listeners.Registrations.FirstOrDefault(item => item.Type == type && item.Capture == capture && ReferenceEquals(item.Callback, callback));
            if (registration != null) Remove(target, listeners.Registrations, registration);
        }
        return JsValue.Undefined;
    });

    private IEventTarget Target(JsValue receiver) => _targetsMap.Resolve(receiver);

    private static bool Capture(JsValue options) => TypeConverter.ToBoolean(options.IsObject() ? options.AsObject().Get("capture") : options);

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
            Registrations.Clear();
        }
    }

    private sealed class Registration(string type, bool capture, Function callback) {
        internal string Type { get; } = type;
        internal bool Capture { get; } = capture;
        internal Function Callback { get; } = callback;
        internal DomEventHandler Handler { get; set; } = null!;
    }
}
