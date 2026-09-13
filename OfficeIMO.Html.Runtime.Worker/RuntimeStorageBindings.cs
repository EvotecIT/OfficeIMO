using Jint;
using Jint.Native;
using Jint.Runtime.Descriptors;
using Jint.Runtime.Interop;
using Jint.Runtime;
using System.Text.Json;

namespace OfficeIMO.Html.Runtime.Worker;

// Both areas belong to one browsing session and are partitioned by document origin.
// No host browser profile or disk storage is read.
internal static class RuntimeStorageBindings {
    internal static void Install(Engine engine, int maxCharacters, RuntimeBrowsingStorage? storage = null, string origin = "") {
        storage ??= new RuntimeBrowsingStorage(maxCharacters);
        var read = new ClrFunction(engine, "readStorage", (_, args) => JsonSerializer.Serialize(storage.Read(origin, args[0].AsBoolean())));
        var write = new ClrFunction(engine, "writeStorage", (_, args) => {
            try { storage.Write(origin, args[0].AsBoolean(), args[1].AsString(), args[2].AsString(), args[3].AsString()); }
            catch (HtmlScriptRuntimeException failure) {
                var error = engine.Intrinsics.Error.Construct(failure.Message);
                error.Set("name", "QuotaExceededError");
                throw new JavaScriptException(error);
            }
            return JsValue.Undefined;
        });
        using var stream = typeof(RuntimeStorageBindings).Assembly.GetManifestResourceStream("OfficeIMO.RuntimeStorageBootstrap.js")!;
        using var reader = new StreamReader(stream);
        var exports = engine.Invoke(engine.Evaluate(reader.ReadToEnd()), new JsValue[] { maxCharacters, read, write }).AsObject();
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
