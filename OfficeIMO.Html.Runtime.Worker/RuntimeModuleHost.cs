using Jint;
using Jint.Native;
using Jint.Runtime;
using Jint.Runtime.Interop;
using Jint.Runtime.Modules;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed class RuntimeModuleHost(Engine engine, Func<RuntimeModuleLoader> modules) : Host {
    public override List<KeyValuePair<JsValue, JsValue>> GetImportMetaProperties(Module moduleRecord) {
        var resolve = new ClrFunction(engine, "resolve", (_, arguments) => {
            string specifier = TypeConverter.ToString(arguments.ElementAtOrDefault(0) ?? JsValue.Undefined);
            return modules().ResolveForImportMeta(specifier, moduleRecord.Location!);
        }, 1);
        return new() { new("url", moduleRecord.Location!), new("resolve", resolve) };
    }
}
