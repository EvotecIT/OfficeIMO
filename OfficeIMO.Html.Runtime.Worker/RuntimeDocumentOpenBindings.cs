using AngleSharp.Dom;
using Jint;
using Jint.Native;
using Jint.Native.Function;
using Jint.Runtime;
using Jint.Runtime.Descriptors;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

internal static class RuntimeDocumentOpenBindings {
    internal static void Install(Engine engine, IDocument document, RuntimeScriptEntry entry, Action<IDocument> opened,
        Func<IWindow, JsValue[], JsValue> openWindow) {
        var open = new ClrFunction(engine, "open", (receiver, args) => {
            if (receiver.ToObject() is not Document target) throw Error(engine, "TypeError", "Document.open requires a Document receiver.");
            if (args.Length >= 3) {
                if (!target.IsFullyActive || target.DefaultView == null) throw Error(engine, "InvalidAccessError", "The document is not fully active.");
                return openWindow(target.DefaultView, args);
            }
            foreach (var argument in args) TypeConverter.ToString(argument);
            try {
                target.OpenFrom(entry.Document ?? document);
                opened(target);
                return JsValue.FromObject(engine, target);
            } catch (DomException error) {
                throw Error(engine, error.Name + "Error", error.Message);
            }
        });
        ClrFunction Write(string name, bool lineFeed) => new(engine, name, (receiver, args) => {
            if (receiver.ToObject() is not Document target) throw Error(engine, "TypeError", $"Document.{name} requires a Document receiver.");
            var content = string.Concat(args.Select(TypeConverter.ToString));
            var wasReady = target.IsReady;
            try {
                if (lineFeed) target.WriteLineFrom(entry.Document ?? document, content);
                else target.WriteFrom(entry.Document ?? document, content);
                if (wasReady && !target.IsReady) opened(target);
                return JsValue.Undefined;
            } catch (DomException error) {
                throw Error(engine, error.Name + "Error", error.Message);
            }
        });
        var write = Write("write", false);
        var writeln = Write("writeln", true);
        for (var prototype = JsValue.FromObject(engine, document).AsObject().Prototype; prototype != null; prototype = prototype.Prototype) {
            if (prototype.GetOwnProperty("open").Value is Function) {
                prototype.FastSetProperty("open", new PropertyDescriptor(open, true, false, true));
            }
            if (prototype.GetOwnProperty("write").Value is Function) {
                prototype.FastSetProperty("write", new PropertyDescriptor(write, true, false, true));
            }
            if (prototype.GetOwnProperty("writeln").Value is Function) {
                prototype.FastSetProperty("writeln", new PropertyDescriptor(writeln, true, false, true));
            }
        }
    }

    private static JavaScriptException Error(Engine engine, string name, string message) {
        var error = engine.Intrinsics.Error.Construct(message);
        error.Set("name", name);
        return new JavaScriptException(error);
    }
}
