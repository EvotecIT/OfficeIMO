using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using Jint;
using Jint.Native;
using Jint.Native.Function;
using Jint.Native.Object;
using Jint.Runtime;
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
        InstallFormMethods(engine, automation);
        InstallViewportMethods(engine, document, automation.Viewport);
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

    private static void InstallFormMethods(Engine engine, RuntimeAutomation automation) {
        if (engine.Global.Get("HTMLFormElement") is not Function constructor
            || constructor.Get("prototype") is not ObjectInstance prototype) return;
        prototype.FastSetProperty("reset", new PropertyDescriptor(new ClrFunction(engine, "reset", (receiver, _) => {
            IHtmlFormElement form = Form(receiver);
            ThrowIfFailed(engine, automation.ResetForm(form));
            return JsValue.Undefined;
        }), true, false, true));
        prototype.FastSetProperty("submit", new PropertyDescriptor(new ClrFunction(engine, "submit", (receiver, _) => {
            IHtmlFormElement form = Form(receiver);
            ThrowIfFailed(engine, automation.SubmitForm(form, null, dispatchSubmitEvent: false, validate: false));
            return JsValue.Undefined;
        }), true, false, true));
        prototype.FastSetProperty("requestSubmit", new PropertyDescriptor(new ClrFunction(engine, "requestSubmit", (receiver, args) => {
            IHtmlFormElement form = Form(receiver);
            IHtmlElement? submitter = args.Length == 0 || args[0].IsUndefined()
                ? null
                : args[0].ToObject() as IHtmlElement ?? throw new ArgumentException("The submitter must be an HTML element.");
            ThrowIfFailed(engine, automation.SubmitForm(form, submitter, dispatchSubmitEvent: true, validate: true));
            return JsValue.Undefined;
        }), true, false, true));
    }

    private static void InstallViewportMethods(Engine engine, IDocument document, RuntimeViewport viewport) {
        if (!viewport.Enabled) return;
        var nativeWindow = JsValue.FromObject(engine, document.DefaultView).AsObject();
        foreach (ObjectInstance target in new[] { engine.Global, nativeWindow }) {
            target.FastSetProperty("innerWidth", Getter(engine, "innerWidth", () => viewport.Width));
            target.FastSetProperty("innerHeight", Getter(engine, "innerHeight", () => viewport.Height));
            target.FastSetProperty("scrollX", Getter(engine, "scrollX", () => viewport.ScrollX));
            target.FastSetProperty("pageXOffset", Getter(engine, "pageXOffset", () => viewport.ScrollX));
            target.FastSetProperty("scrollY", Getter(engine, "scrollY", () => viewport.ScrollY));
            target.FastSetProperty("pageYOffset", Getter(engine, "pageYOffset", () => viewport.ScrollY));
            target.FastSetProperty("scrollTo", new PropertyDescriptor(new ClrFunction(engine, "scrollTo", (_, args) => {
                viewport.ScrollTo(Number(args, 0), Number(args, 1));
                return JsValue.Undefined;
            }), true, false, true));
            target.FastSetProperty("scroll", target.GetOwnProperty("scrollTo"));
            target.FastSetProperty("scrollBy", new PropertyDescriptor(new ClrFunction(engine, "scrollBy", (_, args) => {
                viewport.ScrollBy(Number(args, 0), Number(args, 1));
                return JsValue.Undefined;
            }), true, false, true));
        }

        if (engine.Global.Get("HTMLElement") is not Function constructor
            || constructor.Get("prototype") is not ObjectInstance prototype) return;
        var scrollIntoView = new ClrFunction(engine, "scrollIntoView", (receiver, _) => {
            IElement element = receiver.ToObject() as IElement ?? throw new ArgumentException("scrollIntoView requires an element.");
            viewport.ScrollIntoView(element, CancellationToken.None);
            return JsValue.Undefined;
        });
        var getBoundingClientRect = new ClrFunction(engine, "getBoundingClientRect", (receiver, _) => {
            IElement element = receiver.ToObject() as IElement ?? throw new ArgumentException("getBoundingClientRect requires an element.");
            HtmlRuntimeRect? box = viewport.Measure(element, CancellationToken.None).Box;
            double x = box?.X ?? 0D, y = box?.Y ?? 0D, width = box?.Width ?? 0D, height = box?.Height ?? 0D;
            ObjectInstance rect = engine.Evaluate("({})").AsObject();
            foreach (var pair in new Dictionary<string, double> {
                ["x"] = x, ["y"] = y, ["left"] = x, ["top"] = y,
                ["right"] = x + width, ["bottom"] = y + height, ["width"] = width, ["height"] = height
            }) rect.FastSetProperty(pair.Key, new PropertyDescriptor(pair.Value, false, true, true));
            return rect;
        });
        foreach (var property in engine.Global.GetOwnProperties().ToArray()) {
            if (property.Value.Value is not Function elementConstructor
                || elementConstructor.Get("prototype") is not ObjectInstance candidate) continue;
            ObjectInstance? ancestor = candidate;
            while (ancestor != null && !ReferenceEquals(ancestor, prototype)) ancestor = ancestor.Prototype;
            if (ancestor == null) continue;
            candidate.FastSetProperty("scrollIntoView", new PropertyDescriptor(scrollIntoView, true, false, true));
            candidate.FastSetProperty("getBoundingClientRect", new PropertyDescriptor(getBoundingClientRect, true, false, true));
        }
    }

    private static PropertyDescriptor Getter(Engine engine, string name, Func<double> value) =>
        new GetSetPropertyDescriptor(new ClrFunction(engine, "get " + name, (_, _) => value()), null, true, true);

    private static double Number(JsValue[] args, int index) =>
        index < args.Length ? TypeConverter.ToNumber(args[index]) : 0D;

    private static IHtmlFormElement Form(JsValue receiver) =>
        receiver.ToObject() as IHtmlFormElement ?? throw new ArgumentException("The method requires an HTML form element.");

    private static void ThrowIfFailed(Engine engine, RuntimeFormActionOutcome outcome) {
        if (outcome.Status == HtmlAutomationStatus.Success) return;
        var error = engine.Intrinsics.Error.Construct(outcome.Message);
        error.Set("name", outcome.Status == HtmlAutomationStatus.InvalidValue ? "TypeError" : "NotSupportedError");
        throw new Jint.Runtime.JavaScriptException(error);
    }
}
