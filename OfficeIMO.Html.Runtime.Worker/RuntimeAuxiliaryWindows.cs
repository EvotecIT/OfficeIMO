using AngleSharp;
using AngleSharp.Browser;
using AngleSharp.Dom;
using Jint;
using Jint.Native;
using Jint.Native.Object;
using Jint.Runtime;
using Jint.Runtime.Descriptors;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

// Window policy lives in the host; the DOM provider owns initial-document creation
// and the existing realm owner owns scripts, messages, resources and retirement.
internal sealed class RuntimeAuxiliaryWindows(HtmlScriptRequest options, RuntimeFrameBudget budget,
    RuntimeFrameRealms realms, RuntimeScriptEntry entry, Func<IDocument, Engine?> ensureEngine) {
    private readonly List<IWindow> _windows = [];

    internal void Install(Engine engine, IWindow window) {
        JsValue Wrap(IWindow? value) => realms.WrapWindow(engine, window, value);
        IWindow Target(JsValue receiver) => realms.ResolveWindow(receiver.IsNull() || receiver.IsUndefined() || ReferenceEquals(receiver, engine.Global) ? window : receiver.ToObject() as IWindow
            ?? throw Error(engine, "TypeError", "The receiver must be a Window."));
        var open = new ClrFunction(engine, "open", (receiver, args) => {
            Target(receiver);
            return Open(engine, window, args);
        });
        var close = new ClrFunction(engine, "close", (receiver, _) => {
            var target = Target(receiver);
            if (!realms.IsAuxiliary(target)) throw Error(engine, "NotSupportedError", "Only script-created auxiliary windows can be closed by this profile.");
            realms.CloseAuxiliary(target);
            return JsValue.Undefined;
        });
        var targets = new HashSet<ObjectInstance>(ReferenceEqualityComparer.Instance) { engine.Global };
        for (var current = JsValue.FromObject(engine, window).AsObject(); current != null; current = current.Prototype) targets.Add(current);
        foreach (var target in targets) {
            if (ReferenceEquals(target, engine.Global) || target.GetOwnProperty("open").Value is Jint.Native.Function.Function)
                target.FastSetProperty("open", new PropertyDescriptor(open, true, false, true));
            if (ReferenceEquals(target, engine.Global) || target.GetOwnProperty("close").Value is Jint.Native.Function.Function)
                target.FastSetProperty("close", new PropertyDescriptor(close, true, false, true));
            if (ReferenceEquals(target, engine.Global) || target.GetOwnProperty("parent").Get != null) {
                target.FastSetProperty("parent", new GetSetPropertyDescriptor(new ClrFunction(engine, "get parent", (receiver, _) => {
                    var self = Target(receiver);
                    return Wrap(RuntimeFrameRealms.IsFrameContext(self.Document.Context) ? self.Document.Context.Parent?.Current : self);
                }), JsValue.Undefined, true, true));
                target.FastSetProperty("top", new GetSetPropertyDescriptor(new ClrFunction(engine, "get top", (receiver, _) => {
                    var top = Target(receiver);
                    while (RuntimeFrameRealms.IsFrameContext(top.Document.Context) && top.Document.Context.Parent?.Current is { } parent) top = parent;
                    return Wrap(top);
                }), JsValue.Undefined, true, true));
                target.FastSetProperty("opener", new GetSetPropertyDescriptor(new ClrFunction(engine, "get opener", (receiver, _) => Wrap(realms.OpenerFor(Target(receiver)))), JsValue.Undefined, true, true));
            }
        }
    }

    internal JsValue Open(Engine engine, IWindow realmWindow, JsValue[] args) {
        if (options.Profile != HtmlRuntimeProfile.WebApplicationV1) throw Error(engine, "NotSupportedError", "Auxiliary windows require WebApplicationV1.");
        string url = args.Length > 0 ? TypeConverter.ToString(args[0]) : "";
        string name = args.Length > 1 ? TypeConverter.ToString(args[1]) : "_blank";
        string features = args.Length > 2 ? TypeConverter.ToString(args[2]) : "";
        if (features.Length != 0) throw Error(engine, "NotSupportedError", "Window features are not supported by this runtime profile.");
        if (name.Length > 256) throw Error(engine, "QuotaExceededError", "The auxiliary window name exceeds its character budget.");
        if (name.StartsWith('_') && !name.Equals("_blank", StringComparison.OrdinalIgnoreCase))
            throw Error(engine, "NotSupportedError", "This window-opening operation requires a new auxiliary target.");
        var source = entry.Document ?? realmWindow.Document;
        var parsed = new Url(url.Length == 0 ? "about:blank" : url);
        if (parsed.IsInvalid || parsed.Scheme != "about" || parsed.Data != "blank" || !string.IsNullOrEmpty(parsed.HostName) ||
            !string.IsNullOrEmpty(parsed.UserName) || !string.IsNullOrEmpty(parsed.Password))
            throw Error(engine, "NotSupportedError", "This profile opens initially blank auxiliary windows only.");
        if ((source.Context.Security & Sandboxes.AuxiliaryNavigation) != 0) return JsValue.Null;
        bool named = name.Length != 0 && !name.Equals("_blank", StringComparison.OrdinalIgnoreCase);
        var existing = named ? _windows.FirstOrDefault(window => !window.IsClosed && window.Name == name) : null;
        if (existing != null) {
            if (url.Length != 0) throw Error(engine, "NotSupportedError", "Navigation of an existing auxiliary window is not supported.");
            realms.SetAuxiliaryOpener(existing, source.DefaultView);
            return realms.WrapWindow(engine, realmWindow, existing);
        }
        if (!budget.TryReserveAuxiliaryWindow()) return JsValue.Null;
        var context = source.Context.CreateChild(named ? name : null, source.Context.Security);
        var document = (Document)context.OpenInitialDocument();
        document.DocumentUrl.Href = parsed.Href;
        var window = document.DefaultView!;
        window.Name = named ? name : "";
        realms.RegisterAuxiliary(window, source.DefaultView);
        if (ensureEngine(document) == null) {
            realms.CloseAuxiliary(window);
            return JsValue.Null;
        }
        _windows.Add(window);
        return realms.WrapWindow(engine, realmWindow, window);
    }

    private static JavaScriptException Error(Engine engine, string name, string message) {
        var error = engine.Intrinsics.Error.Construct(message);
        error.Set("name", name);
        return new JavaScriptException(error);
    }
}
