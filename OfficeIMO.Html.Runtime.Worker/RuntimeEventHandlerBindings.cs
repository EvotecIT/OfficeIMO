using System.Runtime.CompilerServices;
using AngleSharp.Dom;
using Jint;
using Jint.Native;
using Jint.Native.Function;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

// Handler properties share one native registration per target/event, regardless of the
// DOM wrapper or prototype through which script accesses them. Replacing a function
// preserves its listener position; clearing and assigning again creates a new position.
internal sealed class RuntimeEventHandlerBindings(Engine engine, IEventTarget window) {
    private readonly ConditionalWeakTable<IEventTarget, Dictionary<string, Registration>> _targets = new();
    private readonly Dictionary<string, (JsValue Get, JsValue Set)> _accessors = new(StringComparer.Ordinal);
    private static readonly HashSet<string> WindowBodyEvents = new(StringComparer.Ordinal) {
        "blur", "error", "focus", "load", "resize", "scroll", "afterprint", "beforeprint", "beforeunload",
        "hashchange", "languagechange", "message", "messageerror", "offline", "online", "pagehide", "pageshow",
        "pageswap", "pagereveal", "popstate", "rejectionhandled", "storage", "unhandledrejection", "unload"
    };

    internal (JsValue Get, JsValue Set) Accessors(string propertyName) {
        if (_accessors.TryGetValue(propertyName, out var existing)) return existing;
        string eventType = propertyName switch {
            "onwebkitanimationend" => "webkitAnimationEnd",
            "onwebkitanimationiteration" => "webkitAnimationIteration",
            "onwebkitanimationstart" => "webkitAnimationStart",
            "onwebkittransitionend" => "webkitTransitionEnd",
            _ => propertyName.Substring(2)
        };
        var getter = new ClrFunction(engine, "get " + propertyName, (receiver, _) => {
            var target = Target(receiver, eventType);
            return target != null && _targets.TryGetValue(target, out var handlers) && handlers.TryGetValue(eventType, out var registration)
                ? registration.Callback : JsValue.Null;
        });
        var setter = new ClrFunction(engine, "set " + propertyName, (receiver, args) => {
            var target = Target(receiver, eventType);
            if (target == null) return JsValue.Undefined;
            var handlers = _targets.GetValue(target, _ => new(StringComparer.Ordinal));
            handlers.TryGetValue(eventType, out var registration);
            if (args.ElementAtOrDefault(0) is Function callback) {
                if (registration != null) registration.Callback = callback;
                else {
                    registration = new Registration(callback);
                    registration.Handler = (sender, ev) => engine.Invoke(registration.Callback,
                        JsValue.FromObject(engine, sender), new[] { JsValue.FromObject(engine, ev) });
                    handlers.Add(eventType, registration);
                    target.AddEventListener(eventType, registration.Handler);
                }
            } else if (registration != null) {
                target.RemoveEventListener(eventType, registration.Handler);
                handlers.Remove(eventType);
            }
            return JsValue.Undefined;
        });
        var accessors = ((JsValue)getter, (JsValue)setter);
        _accessors.Add(propertyName, accessors);
        return accessors;
    }

    private IEventTarget? Target(JsValue receiver, string eventType) {
        var target = ReferenceEquals(receiver, engine.Global) ? window
            : receiver.ToObject() as IEventTarget ?? throw new ArgumentException("An event handler requires a DOM event target.");
        if (target is IElement element && element.NamespaceUri == "http://www.w3.org/1999/xhtml"
            && element.LocalName is "body" or "frameset" && WindowBodyEvents.Contains(eventType)) {
            var view = element.Owner?.DefaultView;
            return view != null && ReferenceEquals(view.Document, element.Owner) ? view : null;
        }
        return target;
    }

    private sealed class Registration(Function callback) {
        internal Function Callback { get; set; } = callback;
        internal DomEventHandler Handler { get; set; } = null!;
    }
}
