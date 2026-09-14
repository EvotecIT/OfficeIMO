using Jint;
using Jint.Native;
using Jint.Runtime.Interop;
using Jint.Runtime.Descriptors;

namespace OfficeIMO.Html.Runtime.Worker;

internal static class RuntimeConsoleBindings {
    internal static void Install(Engine engine, RuntimeDiagnostics diagnostics) {
        JsValue factory = engine.Evaluate("(write)=>Object.freeze({log:(...v)=>write('log',v.map(String).join(' ')),info:(...v)=>write('info',v.map(String).join(' ')),warn:(...v)=>write('warn',v.map(String).join(' ')),error:(...v)=>write('error',v.map(String).join(' ')),debug:(...v)=>write('debug',v.map(String).join(' '))})");
        var write = new ClrFunction(engine, "write", (_, args) => {
            string level = args[0].AsString();
            string message = args[1].AsString();
            diagnostics.Record(HtmlRuntimeEventKind.Console, level, "reported", DateTimeOffset.UtcNow,
                detail: diagnostics.IncludeConsoleMessages ? message : null);
            return JsValue.Undefined;
        });
        JsValue console = engine.Invoke(factory, new JsValue[] { write });
        engine.Global.FastSetProperty("console", new PropertyDescriptor(console, writable: false, enumerable: true, configurable: false));
    }
}
