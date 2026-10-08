using AngleSharp.Dom;
using AngleSharp.Dom.Events;
using Jint;
using Jint.Native;
using Jint.Native.Object;
using Jint.Runtime;
using Jint.Runtime.Descriptors;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

/// <summary>Connects transport EventTarget objects to native Event dispatch and the session's listener bindings.</summary>
internal sealed class RuntimeTransportEventBindings {
    internal ObjectInstance Exports { get; }
    internal JsValue DispatchUntrusted { get; }

    internal RuntimeTransportEventBindings(Engine engine, RuntimeEventTargets targets, ObjectInstance listeners,
        RuntimeEventHandlerBindings handlers, JsValue wrapperFactory, JsValue report, JsValue normalizeWindow, JsValue untrust) {
        using var stream = typeof(RuntimeTransportEventBindings).Assembly.GetManifestResourceStream("OfficeIMO.RuntimeTransportEvents.js")!;
        using var reader = new StreamReader(stream);
        JsValue factory = engine.Evaluate(reader.ReadToEnd());
        var register = new ClrFunction(engine, "registerEventTarget", (_, args) => {
            targets.Register(args[0].AsObject(), args[1].IsUndefined() ? int.MaxValue : (int)args[1].AsNumber());
            return JsValue.Undefined;
        });
        var dispatch = new ClrFunction(engine, "dispatchTransportEvent", (_, args) => {
            var target = targets.Resolve(args[0]);
            Event value = args[1].ToObject() as Event ?? throw new JavaScriptException(engine.Intrinsics.TypeError, "Expected an Event.");
            if (value.Phase != EventPhase.None) throw InvalidState("The event is already being dispatched.");
            ObjectInstance exposed = args[1].AsObject();
            exposed.FastSetProperty("target", new GetSetPropertyDescriptor(
                new ClrFunction(engine, "get target", (_, _) => targets.Wrap(value.OriginalTarget)), JsValue.Undefined, true, true));
            exposed.FastSetProperty("currentTarget", new GetSetPropertyDescriptor(
                new ClrFunction(engine, "get currentTarget", (_, _) => targets.Wrap(value.CurrentTarget)), JsValue.Undefined, true, true));
            RuntimeEventTrust.Set(value, args[2].AsBoolean());
            try { target.Dispatch(value); }
            catch (DomException error) { throw InvalidState(error.Message); }
            return !value.IsDefaultPrevented;
        });
        var handler = new ClrFunction(engine, "eventHandler", (_, args) => {
            var owned = handlers.Accessors(args[0].AsString());
            return engine.Invoke(wrapperFactory, new JsValue[] { owned.Get, owned.Set, report, "handler", normalizeWindow });
        });
        DispatchUntrusted = new ClrFunction(engine, "dispatchEvent", (receiver, args) =>
            engine.Invoke(dispatch, new JsValue[] { receiver, args.ElementAtOrDefault(0) ?? JsValue.Undefined, false }));
        JsValue userDispatch = engine.Invoke(wrapperFactory,
            new[] { DispatchUntrusted, JsValue.Undefined, report, "dispatch", normalizeWindow, untrust });
        Exports = engine.Invoke(factory, new[] { register, listeners.Get("add"), listeners.Get("remove"), dispatch, handler, userDispatch }).AsObject();

        JavaScriptException InvalidState(string message) {
            var error = engine.Intrinsics.Error.Construct(message);
            error.Set("name", "InvalidStateError");
            return new JavaScriptException(error);
        }
    }
}
