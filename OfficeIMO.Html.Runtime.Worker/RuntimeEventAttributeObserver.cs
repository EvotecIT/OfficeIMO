using AngleSharp.Dom;
using AngleSharp.Scripting;
using Jint;
using Jint.Native;

namespace OfficeIMO.Html.Runtime.Worker;

// Attribute handlers use the same property setter as authored JS, including replacement and removal.
internal sealed class RuntimeEventAttributeObserver : IAttributeObserver {
    private readonly Func<IElement, Engine> _getEngine;
    internal RuntimeEventAttributeObserver(Func<IElement, Engine> getEngine) => _getEngine = getEngine;

    public void NotifyChange(IElement host, string name, string? value) {
        if (!name.StartsWith("on", StringComparison.Ordinal)) return;
        Engine engine = _getEngine(host);
        var element = JsValue.FromObject(engine, host).AsObject();
        if (!element.HasProperty(name)) return;
        JsValue callback = value == null ? JsValue.Null : engine.Intrinsics.Function.Construct(new JsValue[] { "event", value }, JsValue.Undefined);
        element.Set(name, callback);
    }
}
