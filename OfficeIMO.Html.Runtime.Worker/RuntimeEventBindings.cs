using Jint;
using Jint.Native;
using Jint.Native.Function;
using Jint.Native.Object;
using Jint.Runtime.Descriptors;

namespace OfficeIMO.Html.Runtime.Worker;

internal static class RuntimeEventBindings {
    internal static void Install(Engine engine, AngleSharp.Dom.IEventTarget window, Action<string> report, JsValue normalizeWindow) {
        using var stream = typeof(RuntimeEventBindings).Assembly.GetManifestResourceStream("OfficeIMO.RuntimeBootstrap.js")!;
        using var reader = new StreamReader(stream);
        JsValue factory = engine.Evaluate(reader.ReadToEnd());
        JsValue reporter = JsValue.FromObject(engine, report);
        var listeners = new RuntimeListenerBindings(engine, window);
        var handlers = new RuntimeEventHandlerBindings(engine, window);
        var untrust = new Jint.Runtime.Interop.ClrFunction(engine,"untrust",(_,args)=>{
            if(args[0].ToObject() is AngleSharp.Dom.Events.Event value && value.Phase==0)RuntimeEventTrust.Set(value,false);
            return JsValue.Undefined;
        });
        var listenerMethods = engine.Invoke(factory, new[] { listeners.Add, listeners.RemoveListener, reporter, JsValue.Undefined, normalizeWindow }).AsObject();
        var visited = new HashSet<ObjectInstance>(ReferenceEqualityComparer.Instance) { engine.Global };
        foreach (var property in engine.Global.GetOwnProperties().ToArray()) {
            if (property.Value.Value is Function constructor && constructor.Get("prototype") is ObjectInstance prototype) visited.Add(prototype);
        }
        foreach (var prototype in visited) {
            var trusted = prototype.GetOwnProperty("isTrusted");
            if(trusted.Get is Function)
                prototype.FastSetProperty("isTrusted",new GetSetPropertyDescriptor(new Jint.Runtime.Interop.ClrFunction(engine,"get isTrusted",(receiver,_)=>
                    RuntimeEventTrust.Read(receiver.ToObject() as AngleSharp.Dom.Events.Event ?? throw new ArgumentException("Expected an event."))),JsValue.Undefined,trusted.Enumerable,trusted.Configurable));
            JsValue add = prototype.GetOwnProperty("addEventListener").Value;
            JsValue remove = prototype.GetOwnProperty("removeEventListener").Value;
            if (add is Function && remove is Function) {
                // Use the same native method pair across prototypes so the provider retains
                // the same callback registration when add and remove use different prototypes.
                prototype.FastSetProperty("addEventListener", new PropertyDescriptor(listenerMethods.Get("add"), true, false, true));
                prototype.FastSetProperty("removeEventListener", new PropertyDescriptor(listenerMethods.Get("remove"), true, false, true));
            }
            if (prototype.GetOwnProperty("dispatchEvent").Value is Function dispatch) {
                // Normalize the JS return contract against the event's cancellation state.
                var wrapper = engine.Invoke(factory, new JsValue[] { dispatch, JsValue.Undefined, reporter, "dispatch", normalizeWindow, untrust });
                prototype.FastSetProperty("dispatchEvent", new PropertyDescriptor(wrapper, true, false, true));
            }
            if (prototype.GetOwnProperty("composedPath").Value is Function) {
                var wrapper = engine.Invoke(factory, new JsValue[] { JsValue.Undefined, JsValue.Undefined, reporter, "path", normalizeWindow });
                prototype.FastSetProperty("composedPath", new PropertyDescriptor(wrapper, true, false, true));
            }
            foreach (var handler in prototype.GetOwnProperties().ToArray()) {
                if (!handler.Key.IsString() || !handler.Key.AsString().StartsWith("on", StringComparison.Ordinal) ||
                    handler.Value.Get is not Function || handler.Value.Set is not Function) continue;
                var owned = handlers.Accessors(handler.Key.AsString());
                var accessors = engine.Invoke(factory, new JsValue[] { owned.Get, owned.Set, reporter, "handler", normalizeWindow }).AsObject();
                prototype.FastSetProperty(handler.Key, new GetSetPropertyDescriptor(accessors.Get("get"), accessors.Get("set"), handler.Value.Enumerable, true));
            }
        }
    }
}
