using System.Net;
using AngleSharp.Dom;
using Jint;
using Jint.Native;
using Jint.Runtime.Descriptors;

namespace OfficeIMO.Html.Runtime.Worker;

// Retain the provider URL parser; replace its incomplete query-list/stringifier binding.
internal static class RuntimeUrlBindings {
    internal static void Install(Engine engine, IWindow window) {
        using var stream = typeof(RuntimeUrlBindings).Assembly.GetManifestResourceStream("OfficeIMO.RuntimeUrlBootstrap.js")!;
        using var reader = new StreamReader(stream);
        var factory = engine.Evaluate(reader.ReadToEnd());
        var exports = engine.Invoke(factory, new[] { JsValue.FromObject(engine, (Func<string, string>)WebUtility.UrlDecode) }).AsObject();
        var constructor = exports.Get("URLSearchParams");
        engine.Global.FastSetProperty("URLSearchParams", new PropertyDescriptor(constructor, true, false, true));
        JsValue.FromObject(engine, window).AsObject().FastSetProperty("URLSearchParams", new PropertyDescriptor(constructor, true, false, true));
        engine.Global.Get("URL").AsObject().Get("prototype").AsObject().FastSetProperty("searchParams",
            new GetSetPropertyDescriptor(exports.Get("getSearchParams"), JsValue.Undefined, true, true));
    }
}
