using AngleSharp.Dom;
using Jint;
using Jint.Native;
using Jint.Native.Function;
using Jint.Native.Object;
using Jint.Runtime;
using Jint.Runtime.Descriptors;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

// Reject unsupported native navigation before the provider mutates its URL or
// starts loading. Qualified root navigation is installed by RuntimeHistoryBindings.
internal static class RuntimeNativeNavigationBindings {
    internal static void Install(Engine engine, IWindow window, Func<object?, bool> blocked) {
        void Check(JsValue receiver) {
            object? target = receiver.IsNull() || receiver.IsUndefined() || ReferenceEquals(receiver, engine.Global)
                ? window : receiver.ToObject();
            if (!blocked(target)) return;
            var error = engine.Intrinsics.Error.Construct("Navigation of auxiliary windows is not supported.");
            error.Set("name", "NotSupportedError");
            throw new JavaScriptException(error);
        }

        var targets = new HashSet<ObjectInstance>(ReferenceEqualityComparer.Instance) { engine.Global };
        foreach (var native in new object[] { window, window.Document, window.Document.Location }) {
            for (var current = JsValue.FromObject(engine, native).AsObject(); current != null; current = current.Prototype)
                targets.Add(current);
        }
        foreach (var target in targets) {
            foreach (string name in new[] { "location", "href", "hash", "host", "hostname", "pathname", "port", "protocol", "search", "username", "password" }) {
                var property = target.GetOwnProperty(name);
                if (property.Set is not Function setter) continue;
                target.FastSetProperty(name, new GetSetPropertyDescriptor(property.Get,
                    new ClrFunction(engine, "set " + name, (receiver, args) => {
                        Check(receiver);
                        return engine.Invoke(setter, receiver, args);
                    }), property.Enumerable, property.Configurable));
            }
            foreach (string name in new[] { "assign", "replace", "reload" }) {
                var property = target.GetOwnProperty(name);
                if (property.Value is not Function method) continue;
                target.FastSetProperty(name, new PropertyDescriptor(new ClrFunction(engine, name, (receiver, args) => {
                    Check(receiver);
                    return engine.Invoke(method, receiver, args);
                }), property.Writable, property.Enumerable, property.Configurable));
            }
        }
    }
}
