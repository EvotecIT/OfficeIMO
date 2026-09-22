using Jint;
using Jint.Native;
using Jint.Native.Object;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

// Interpreter objects never cross the process boundary. This clone service owns
// independent state graphs within one realm, using native brands rather than tags.
internal static class RuntimeStructuredClone {
    internal static JsValue Create(Engine engine, int maximumBytes) {
        using var stream = typeof(RuntimeStructuredClone).Assembly.GetManifestResourceStream("OfficeIMO.RuntimeStructuredClone.js")!;
        using var reader = new StreamReader(stream);
        return engine.Invoke(engine.Evaluate(reader.ReadToEnd()), new JsValue[] { Brand(engine), maximumBytes, Errors(engine), Buffers(engine) });
    }

    internal static ObjectInstance CreateTransport(Engine engine, int maximumCharacters) {
        using var stream = typeof(RuntimeStructuredClone).Assembly.GetManifestResourceStream("OfficeIMO.RuntimeStructuredCloneTransport.js")!;
        using var reader = new StreamReader(stream);
        return engine.Invoke(engine.Evaluate(reader.ReadToEnd()), new JsValue[] { Brand(engine), maximumCharacters, Errors(engine), Buffers(engine) }).AsObject();
    }

    private static JsValue Buffers(Engine engine) {
        using var stream = typeof(RuntimeStructuredClone).Assembly.GetManifestResourceStream("OfficeIMO.RuntimeStructuredCloneBuffers.js")!;
        using var reader = new StreamReader(stream);
        var tracking = new ClrFunction(engine, "isLengthTrackingArrayBufferView", (_, args) => args[0].IsLengthTrackingArrayBufferView());
        return engine.Invoke(engine.Evaluate(reader.ReadToEnd()), new JsValue[] { tracking });
    }

    private static JsValue Errors(Engine engine) {
        using var stream = typeof(RuntimeStructuredClone).Assembly.GetManifestResourceStream("OfficeIMO.RuntimeStructuredCloneErrors.js")!;
        using var reader = new StreamReader(stream);
        return engine.Evaluate(reader.ReadToEnd());
    }

    private static ClrFunction Brand(Engine engine) => new(engine, "structuredCloneBrand", (_, args) => args[0] switch {
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
}
