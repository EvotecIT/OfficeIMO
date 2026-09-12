using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using Jint;
using Jint.Native;
using Jint.Native.Function;
using Jint.Native.Object;
using Jint.Runtime.Descriptors;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

internal static class RuntimeInteractionBindings {
    internal static void Install(Engine engine, IDocument document, RuntimeFocusController focus, RuntimeAutomation automation) {
        var htmlPrototype = engine.Global.Get("HTMLElement").AsObject().Get("prototype").AsObject();
        var focusMethod = new ClrFunction(engine, "focus", (receiver, _) => {
            focus.Focus(receiver.ToObject() as IHtmlElement ?? throw new ArgumentException("Focus requires an HTML element."));
            return JsValue.Undefined;
        });
        var blurMethod = new ClrFunction(engine, "blur", (receiver, _) => {
            focus.Blur(receiver.ToObject() as IHtmlElement ?? throw new ArgumentException("Blur requires an HTML element."));
            return JsValue.Undefined;
        });
        var clicking = new HashSet<IHtmlElement>(ReferenceEqualityComparer.Instance);
        var clickMethod = new ClrFunction(engine,"click",(receiver,_)=>{
            var element = receiver.ToObject() as IHtmlElement ?? throw new ArgumentException("Click requires an HTML element.");
            if (RuntimeFocusController.Disabled(element) || !clicking.Add(element)) return JsValue.Undefined;
            try {
                var result = automation.Activate(element,focusTarget:false);
                if (result.Status != HtmlAutomationStatus.Success) {
                    var error = engine.Intrinsics.Error.Construct(result.Message);
                    error.Set("name","NotSupportedError");
                    throw new Jint.Runtime.JavaScriptException(error);
                }
                return JsValue.Undefined;
            } finally { clicking.Remove(element); }
        });
        // The JS adapter repeats inherited native members on concrete DOM prototypes.
        // Replace each declaration so an HTMLInputElement cannot retain the native stub.
        foreach (var property in engine.Global.GetOwnProperties().ToArray()) {
            if (property.Value.Value is not Function constructor || constructor.Get("prototype") is not ObjectInstance prototype) continue;
            var ancestor = prototype;
            while (ancestor != null && !ReferenceEquals(ancestor, htmlPrototype)) ancestor = ancestor.Prototype;
            if (ancestor == null) continue;
            if (prototype.GetOwnProperty("focus").Value is Function)
                prototype.FastSetProperty("focus", new PropertyDescriptor(focusMethod, true, false, true));
            if (prototype.GetOwnProperty("blur").Value is Function)
                prototype.FastSetProperty("blur", new PropertyDescriptor(blurMethod, true, false, true));
            if (prototype.GetOwnProperty("click").Value is Function)
                prototype.FastSetProperty("click", new PropertyDescriptor(clickMethod, true, false, true));
        }
        // Script click and typed actions share activation, including cancellation and navigation.
        for (var prototype = JsValue.FromObject(engine, document).AsObject().Prototype; prototype != null; prototype = prototype.Prototype) {
            var descriptor = prototype.GetOwnProperty("activeElement");
            if (descriptor.Get == null) continue;
            var getter = new ClrFunction(engine, "get activeElement", (receiver, _) => {
                if (!ReferenceEquals(receiver.ToObject(), document)) throw new ArgumentException("The getter requires this session's document.");
                return JsValue.FromObject(engine, focus.Focused ?? document.Body ?? document.DocumentElement);
            });
            prototype.FastSetProperty("activeElement", new GetSetPropertyDescriptor(getter, null, descriptor.Enumerable, descriptor.Configurable));
            break;
        }
    }
}
