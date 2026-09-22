using AngleSharp.Dom;
using Jint;
using Jint.Native;
using Jint.Native.Function;
using Jint.Native.Object;
using Jint.Runtime;
using Jint.Runtime.Descriptors;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

// A captured opener reference names the browsing context, not its old document.
// Forward native window accessors through the current root while keeping the
// provider's wrapper identity. Cross-origin navigation must revoke DOM access.
internal static class RuntimeWindowNavigationBindings {
    internal static void Install(Engine engine, IWindow caller, RuntimeFrameRealms realms) {
        var prototypes = new HashSet<ObjectInstance>(ReferenceEqualityComparer.Instance) { engine.Global };
        for (var current = JsValue.FromObject(engine, caller).AsObject(); current != null; current = current.Prototype)
            prototypes.Add(current);
        var windowPrototypes = new HashSet<ObjectInstance>(prototypes, ReferenceEqualityComparer.Instance);
        var windowMethods = new HashSet<string>(StringComparer.Ordinal) {
            "addEventListener", "removeEventListener", "dispatchEvent", "setTimeout", "setInterval",
            "clearTimeout", "clearInterval", "open", "close", "stop", "focus", "blur", "alert",
            "confirm", "print", "postMessage", "getComputedStyle", "matchMedia"
        };
        // Borrowed EventTarget methods and handler accessors must use the same
        // window check, even when reached through another DOM prototype.
        foreach (var property in engine.Global.GetOwnProperties().ToArray()) {
            if (property.Value.Value is Function constructor && constructor.Get("prototype") is ObjectInstance prototype)
                prototypes.Add(prototype);
        }
        foreach (var prototype in prototypes) {
            foreach (var property in prototype.GetOwnProperties().ToArray()) {
                if (!property.Key.IsString()) continue;
                string name = property.Key.AsString();
                if (name is "__proto__" or "constructor" or "prototype") continue;
                if (!windowPrototypes.Contains(prototype) && !windowMethods.Contains(name) && !name.StartsWith("on", StringComparison.Ordinal)) continue;
                if (property.Value.Value is Function method && windowMethods.Contains(name)) {
                    prototype.FastSetProperty(property.Key, new PropertyDescriptor(
                        new ClrFunction(engine, name, (receiver, args) => engine.Invoke(method, Receiver(receiver, name), args)),
                        property.Value.Writable, property.Value.Enumerable, property.Value.Configurable));
                }
                Function? getter = property.Value.Get as Function;
                Function? setter = property.Value.Set as Function;
                if (getter == null && setter == null) continue;
                JsValue read = getter == null ? JsValue.Undefined : new ClrFunction(engine, "get " + name, (receiver, args) => {
                    var target = Receiver(receiver, name);
                    var result = engine.Invoke(getter, target, args);
                    return !ReferenceEquals(result, engine.Global) && result.ToObject() is IWindow window ? realms.WrapWindow(engine, caller, window) : result;
                });
                JsValue write = setter == null ? JsValue.Undefined : new ClrFunction(engine, "set " + name, (receiver, args) =>
                    engine.Invoke(setter, Receiver(receiver, name, writing: true), args));
                prototype.FastSetProperty(property.Key, new GetSetPropertyDescriptor(read, write,
                    property.Value.Enumerable, property.Value.Configurable));
            }
        }

        JsValue Receiver(JsValue receiver, string property, bool writing = false) {
            if (ReferenceEquals(receiver, engine.Global)) return receiver;
            if (receiver.ToObject() is not IWindow window) return receiver;
            var current = realms.ResolveWindow(window);
            bool sameOrigin = string.Equals(RuntimeDocumentUrls.Origin(caller.Document), RuntimeDocumentUrls.Origin(current.Document), StringComparison.OrdinalIgnoreCase);
            if (!sameOrigin && (writing || property is not ("closed" or "length" or "window" or "self" or "frames" or "parent" or "top" or "opener" or "postMessage" or "close" or "focus" or "blur"))) {
                var error = engine.Intrinsics.Error.Construct("The target window is no longer same-origin.");
                error.Set("name", "SecurityError");
                throw new JavaScriptException(error);
            }
            return JsValue.FromObject(engine, current);
        }
    }
}
