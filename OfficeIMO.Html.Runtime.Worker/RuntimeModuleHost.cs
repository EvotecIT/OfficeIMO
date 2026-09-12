using Jint.Native;
using Jint.Runtime;
using Jint.Runtime.Modules;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed class RuntimeModuleHost : Host {
    public override List<KeyValuePair<JsValue, JsValue>> GetImportMetaProperties(Module moduleRecord) =>
        new() { new("url", moduleRecord.Location!) };
}
