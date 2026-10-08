using System.Runtime.CompilerServices;
using AngleSharp.Dom;
using Jint;
using Jint.Native;
using Jint.Native.Object;

namespace OfficeIMO.Html.Runtime.Worker;

/// <summary>Maps script-owned transport objects to the same native event targets used by DOM bindings.</summary>
internal sealed class RuntimeEventTargets(Engine engine, IEventTarget window) {
    private readonly ConditionalWeakTable<ObjectInstance, TransportTarget> _transports = new();

    internal void Register(ObjectInstance wrapper, int listenerLimit) => _transports.Add(wrapper, new TransportTarget(wrapper, listenerLimit));

    internal static int ListenerLimit(IEventTarget target) => target is TransportTarget transport ? transport.ListenerLimit : int.MaxValue;
    internal static bool IsTransport(IEventTarget target) => target is TransportTarget;
    internal bool IsTransport(JsValue receiver) => receiver.IsObject() && _transports.TryGetValue(receiver.AsObject(), out _);

    internal IEventTarget Resolve(JsValue receiver) {
        if (ReferenceEquals(receiver, engine.Global)) return window;
        if (receiver.IsObject() && _transports.TryGetValue(receiver.AsObject(), out var target)) return target;
        return receiver.ToObject() as IEventTarget
            ?? throw new HtmlScriptRuntimeException("The event receiver must be an event target.");
    }

    internal JsValue Wrap(object? target) => target is TransportTarget transport
        ? transport.Wrapper : JsValue.FromObject(engine, target);

    private sealed class TransportTarget(ObjectInstance wrapper, int listenerLimit) : EventTarget {
        internal ObjectInstance Wrapper { get; } = wrapper;
        internal int ListenerLimit { get; } = listenerLimit;
    }
}
