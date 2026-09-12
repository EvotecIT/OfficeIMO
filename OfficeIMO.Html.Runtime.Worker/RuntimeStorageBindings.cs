using Jint;
using Jint.Native;
using Jint.Runtime.Descriptors;

namespace OfficeIMO.Html.Runtime.Worker;

// Both areas belong to this one-document session. No host browser profile or disk
// storage is read, and opening another session starts with independent empty areas.
internal static class RuntimeStorageBindings {
    internal static void Install(Engine engine, int maxCharacters) {
        using var stream = typeof(RuntimeStorageBindings).Assembly.GetManifestResourceStream("OfficeIMO.RuntimeStorageBootstrap.js")!;
        using var reader = new StreamReader(stream);
        var exports = engine.Invoke(engine.Evaluate(reader.ReadToEnd()), new JsValue[] { maxCharacters }).AsObject();
        foreach (var property in exports.GetOwnProperties()) {
            if (property.Key == "Storage") engine.Global.FastSetProperty(property.Key, new PropertyDescriptor(property.Value.Value, true, false, true));
            else {
                var value = property.Value.Value;
                var getter = new Jint.Runtime.Interop.ClrFunction(engine, "get " + property.Key, (_, _) => value);
                engine.Global.FastSetProperty(property.Key, new GetSetPropertyDescriptor(getter, JsValue.Undefined, true, true));
            }
        }
    }
}
