using AngleSharp.Dom;
using Jint;
using Jint.Native;
using Jint.Native.Function;
using Jint.Native.Object;
using Jint.Runtime.Descriptors;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

// A browsing session has one global window. Normalize the provider's separate DOM
// window proxy at the public getter/callback boundaries while retaining its DOM APIs.
internal static class RuntimeWindowBindings {
    internal static JsValue Install(Engine engine, IWindow window) {
        var nativeWindow = JsValue.FromObject(engine, window).AsObject();
        var normalize = new ClrFunction(engine, "normalizeWindow", (_, args) => ReferenceEquals(args[0], nativeWindow) ? engine.Global : args[0]);
        var prototypes = new HashSet<ObjectInstance>(ReferenceEqualityComparer.Instance) { engine.Global, nativeWindow };
        foreach (var property in engine.Global.GetOwnProperties().ToArray()) {
            if (property.Value.Value is Function constructor && constructor.Get("prototype") is ObjectInstance prototype) prototypes.Add(prototype);
        }
        var names = new HashSet<string>(StringComparer.Ordinal) { "window", "self", "parent", "top", "frames", "defaultView", "view", "target", "currentTarget", "srcElement" };
        using var stream = typeof(RuntimeWindowBindings).Assembly.GetManifestResourceStream("OfficeIMO.RuntimeWindowBootstrap.js")!;
        using var reader = new StreamReader(stream);
        var methodFactory = engine.Evaluate(reader.ReadToEnd());
        var unsupportedWorker = engine.Evaluate("(function Worker(){const error=new Error('Dedicated workers are outside this session profile');error.name='NotSupportedError';throw error;})");
        foreach (var prototype in prototypes) {
            var worker = prototype.GetOwnProperty("Worker");
            if (worker.Value is Function)
                prototype.FastSetProperty("Worker", new PropertyDescriptor(unsupportedWorker, worker.Writable, worker.Enumerable, worker.Configurable));
            foreach (string name in new[] { "setTimeout", "setInterval", "clearTimeout", "clearInterval" }) {
                var descriptor = prototype.GetOwnProperty(name);
                if (descriptor.Value is not Function timer) continue;
                var wrapped = engine.Invoke(methodFactory, new JsValue[] { timer, normalize, name });
                prototype.FastSetProperty(name, new PropertyDescriptor(wrapped, descriptor.Writable, descriptor.Enumerable, descriptor.Configurable));
            }
            foreach (var property in prototype.GetOwnProperties().ToArray()) {
                if (!property.Key.IsString() || !names.Contains(property.Key.AsString()) || property.Value.Get is not Function getter) continue;
                var wrapped = new ClrFunction(engine, "get " + property.Key.AsString(), (receiver, args) => {
                    var result = engine.Invoke(getter, receiver, args);
                    return ReferenceEquals(result, nativeWindow) ? engine.Global : result;
                });
                prototype.FastSetProperty(property.Key, new GetSetPropertyDescriptor(wrapped, property.Value.Set, property.Value.Enumerable, property.Value.Configurable));
            }
        }
        foreach (string name in new[] { "window", "self", "parent", "top", "frames" }) {
            bool fixedAlias = name is "window" or "top";
            engine.Global.FastSetProperty(name, new PropertyDescriptor(engine.Global, !fixedAlias, true, !fixedAlias));
        }
        return normalize;
    }
}
