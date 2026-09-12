using Jint;
using Jint.Native;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

// Interpreter objects never cross the process boundary. This clone service owns
// independent state graphs within one realm, using native brands rather than tags.
internal static class RuntimeStructuredClone {
    internal static JsValue Create(Engine engine, int maximumBytes) {
        var brand = new ClrFunction(engine, "stateBrand", (_, args) => args[0] switch {
            JsObject => "object",
            JsArray => "array",
            JsDate => "date",
            JsRegExp => "regexp",
            JsMap => "map",
            JsSet => "set",
            JsError => "error",
            JsArrayBuffer buffer when buffer.GetType() == typeof(JsArrayBuffer) => "buffer",
            _ => ReferenceEquals(args[0], engine.Intrinsics.Object.PrototypeObject) ? "object" : "exotic"
        });
        using var stream = typeof(RuntimeStructuredClone).Assembly.GetManifestResourceStream("OfficeIMO.RuntimeStructuredClone.js")!;
        using var reader = new StreamReader(stream);
        return engine.Invoke(engine.Evaluate(reader.ReadToEnd()), new JsValue[] { brand, maximumBytes });
    }
}
