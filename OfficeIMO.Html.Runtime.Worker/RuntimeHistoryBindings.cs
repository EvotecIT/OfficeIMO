using AngleSharp.Browser;
using AngleSharp.Dom;
using Jint;
using Jint.Native;
using Jint.Native.Function;
using Jint.Native.Object;
using Jint.Runtime;
using Jint.Runtime.Descriptors;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

// Own history while retaining the DOM's public URL parser and document URL record.
internal sealed class RuntimeHistoryBindings {
    private readonly Engine _engine;
    private readonly JsValue _navigate;

    internal RuntimeHistoryBindings(Engine engine, IDocument document, IEventLoop loop, HtmlScriptRequest options) {
        _engine = engine;
        var nativeDocument = document as Document ?? throw new HtmlScriptRuntimeException("The DOM provider does not expose its document URL record.");
        var current = new ClrFunction(engine, "currentUrl", (_, _) => document.Url);
        var parse = new ClrFunction(engine, "parseRoute", (_, args) => {
            string input = TypeConverter.ToString(args[0]);
            var url = new Url(new Url(RuntimeDocumentUrls.Base(document)), input);
            if (url.IsInvalid) throw Error(engine, "SecurityError", "The route URL is invalid.");
            try { HtmlRuntimeResourcePolicy.ValidateUrl(new Uri(url.Href)); }
            catch (Exception error) when (error is ArgumentException or UriFormatException) { throw Error(engine, "SecurityError", "Route URLs must use HTTP(S) without credentials."); }
            if (url.Origin != new Url(document.Url).Origin) throw Error(engine, "SecurityError", "History cannot change the document origin.");
            return url.Href;
        });
        var update = new ClrFunction(engine, "updateRoute", (_, args) => {
            // Mutate the retained URL record through its public API. The Location
            // setter would also start the provider's independent navigation flow.
            RuntimeDocumentUrls.Rewrite(nativeDocument,args[0].AsString());
            return JsValue.Undefined;
        });
        var enqueue = new ClrFunction(engine, "queueTraversal", (_, args) => {
            var callback = args[0];
            loop.Enqueue(_ => engine.Invoke(callback), TaskPriority.Normal);
            return JsValue.Undefined;
        });
        var dispatch = new ClrFunction(engine,"dispatchHistoryEvent",(_,args)=>{
            var historyEvent=(AngleSharp.Dom.Events.Event)args[0].ToObject()!;
            RuntimeEventTrust.Set(historyEvent,true);
            document.DefaultView!.Dispatch(historyEvent);
            return JsValue.Undefined;
        });
        using var stream = typeof(RuntimeHistoryBindings).Assembly.GetManifestResourceStream("OfficeIMO.RuntimeHistoryBootstrap.js")!;
        using var reader = new StreamReader(stream);
        var exports = engine.Invoke(engine.Evaluate(reader.ReadToEnd()), new JsValue[] {
            current, parse, update, enqueue, RuntimeStructuredClone.Create(engine, options.MaxHistoryStateBytes), options.MaxHistoryEntries,
            options.MaxHistoryTotalStateBytes, options.MaxPendingHistoryTasks, dispatch
        }).AsObject();
        _navigate = exports.Get("navigate");
        var history = exports.Get("history");
        var location = exports.Get("location");
        var nativeWindow = JsValue.FromObject(engine, document.DefaultView).AsObject();
        foreach (var property in engine.Global.GetOwnProperties().ToArray()) {
            if (property.Value.Value is not Function constructor || constructor.Get("prototype") is not ObjectInstance prototype) continue;
            foreach(string name in new[]{"baseURI","href","src","action","formAction"}) {
                var descriptor=prototype.GetOwnProperty(name);
                if (descriptor.Get is not Function getter) continue;
                prototype.FastSetProperty(name,new GetSetPropertyDescriptor(new ClrFunction(engine,"get "+name,(receiver,args)=>{
                    RuntimeDocumentUrls.Base(document);
                    if(name=="href" && receiver.ToObject() is AngleSharp.Html.Dom.IHtmlBaseElement element)
                        return new Url(new Url(element.Owner?.Url ?? document.Url),element.GetAttribute("href") ?? "").Href;
                    return engine.Invoke(getter,receiver,args);
                }),descriptor.Set,descriptor.Enumerable,descriptor.Configurable));
            }
        }
        foreach (var target in new[] { engine.Global, nativeWindow }) {
            foreach(string name in new[]{"History","Location","PopStateEvent","HashChangeEvent"})
                target.FastSetProperty(name,new PropertyDescriptor(exports.Get(name),true,false,true));
            target.FastSetProperty("history", new GetSetPropertyDescriptor(new ClrFunction(engine,"get history",(_,_)=>history), JsValue.Undefined, true, true));
            target.FastSetProperty("location", new GetSetPropertyDescriptor(new ClrFunction(engine,"get location",(_,_)=>location),
                new ClrFunction(engine,"set location",(_,args)=>engine.Invoke(_navigate,new JsValue[]{TypeConverter.ToString(args[0]),false})),true,false));
        }
        for (var prototype = JsValue.FromObject(engine,document).AsObject().Prototype; prototype != null; prototype = prototype.Prototype) {
            if (prototype.GetOwnProperty("location").Get is not Function) continue;
            prototype.FastSetProperty("location", new GetSetPropertyDescriptor(new ClrFunction(engine,"get location",(receiver,_)=>{
                if (!ReferenceEquals(receiver.ToObject(),document)) throw Error(engine,"NotSupportedError","Location belongs to this session's document.");
                return location;
            }), new ClrFunction(engine,"set location",(receiver,args)=>{
                if (!ReferenceEquals(receiver.ToObject(),document)) throw Error(engine,"NotSupportedError","Location belongs to this session's document.");
                return engine.Invoke(_navigate,new JsValue[]{TypeConverter.ToString(args[0]),false});
            }),true,false));
        }
    }

    internal void NavigateFragment(string target, bool replace = false) => _engine.Invoke(_navigate, new JsValue[] { target, replace });

    private static JavaScriptException Error(Engine engine, string name, string message) {
        var error = engine.Intrinsics.Error.Construct(message);
        error.Set("name", name);
        return new JavaScriptException(error);
    }
}
