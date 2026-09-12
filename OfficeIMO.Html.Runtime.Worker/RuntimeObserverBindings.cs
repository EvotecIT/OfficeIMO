using Jint;
using Jint.Native;
using Jint.Runtime.Descriptors;
using Jint.Runtime.Interop;
using AngleSharp.Dom;

namespace OfficeIMO.Html.Runtime.Worker;

internal static class RuntimeObserverBindings {
    internal static void Install(Engine engine, IDocument document, Action<string> report) {
        using var stream = typeof(RuntimeObserverBindings).Assembly.GetManifestResourceStream("OfficeIMO.RuntimeObserverBootstrap.js")!;
        using var reader = new StreamReader(stream);
        var factory = engine.Evaluate(reader.ReadToEnd());
        var create = new ClrFunction(engine, "createObserver", (_, args) => {
            var callback = args[0];
            var observer = new MutationObserver((records, _) => {
                lock (engine) {
                    try {
                        var values = new JsArray(engine, records.Select(record => JsValue.FromObject(engine, record)).ToArray());
                        engine.Invoke(callback, new JsValue[] { values });
                    } catch (Exception error) { report(error.Message); }
                }
            });
            return JsValue.FromObject(engine, observer);
        });
        var take = new ClrFunction(engine, "takeRecords", (_, args) => {
            var observer = (MutationObserver)args[0].ToObject()!;
            return new JsArray(engine, observer.Flush().Select(record => JsValue.FromObject(engine, record)).ToArray());
        });
        var observe = new ClrFunction(engine, "observe", (_, args) => {
            var observer = (MutationObserver)args[0].ToObject()!;
            var target = args[1].ToObject() as INode ?? throw new ArgumentException("The observation target must be a Node.");
            var options = args[2].AsObject();
            bool? Optional(string name) => options.Get(name).IsUndefined() ? null : options.Get(name).AsBoolean();
            var filter = options.Get("attributeFilter");
            observer.Connect(target, options.Get("childList").AsBoolean(), options.Get("subtree").AsBoolean(),
                Optional("attributes"), Optional("characterData"), Optional("attributeOldValue"), Optional("characterDataOldValue"),
                filter.IsUndefined() ? null : filter.AsArray().Select(value => value.AsString()).ToArray());
            return JsValue.Undefined;
        });
        // The DOM provider uses null for an inapplicable node list. Web callers
        // require a NodeList even for attribute and character-data records.
        var emptyNodes = JsValue.FromObject(engine, document.CreateDocumentFragment().ChildNodes);
        var recordPrototype = engine.Global.Get("MutationRecord").AsObject().Get("prototype").AsObject();
        foreach (string name in new[] { "addedNodes", "removedNodes" }) {
            var descriptor = recordPrototype.GetOwnProperty(name);
            var getter = descriptor.Get!;
            var wrapped = new ClrFunction(engine, "get " + name, (receiver, args) => {
                var result = engine.Invoke(getter, receiver, args);
                return result.IsNull() ? emptyNodes : result;
            });
            recordPrototype.FastSetProperty(name, new GetSetPropertyDescriptor(wrapped, descriptor.Set, descriptor.Enumerable, descriptor.Configurable));
        }
        var exports = engine.Invoke(factory, new[] { create, take, observe, JsValue.FromObject(engine, report) }).AsObject();
        foreach (var property in exports.GetOwnProperties())
            engine.Global.FastSetProperty(property.Key, new PropertyDescriptor(property.Value.Value, true, false, true));
    }
}
