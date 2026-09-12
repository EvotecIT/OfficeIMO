using System.Text;
using AngleSharp.Browser;
using AngleSharp.Dom;
using AngleSharp.Io;
using AngleSharp.Scripting;
using Jint.Runtime.Modules;

namespace OfficeIMO.Html.Runtime.Worker;

// Module evaluation yields the native event loop while top-level awaits are pending.
// Classic execution and the retained DOM wrappers continue using the same Jint engine.
internal sealed class RuntimeScriptingService(JsScriptingService scripting, Func<RuntimeModuleLoader> modules, HtmlScriptRequest request) : IScriptingService {
    public string Type => scripting.Type;
    public bool SupportsType(string type) => ((IScriptingService)scripting).SupportsType(type);

    public object EvaluateScript(IDocument document, string source, string type, string sourceUrl) {
        if (type?.Equals("importmap", StringComparison.OrdinalIgnoreCase) == true) {
            modules().ImportMap.Load(source, new Uri(RuntimeDocumentUrls.Base(document)));
            return null!;
        }
        if (type?.Equals("module", StringComparison.OrdinalIgnoreCase) == true)
            throw new HtmlScriptRuntimeException("Module scripts require asynchronous document execution.");
        var engine = scripting.GetOrCreateJint(document);
        string location = string.IsNullOrEmpty(sourceUrl) ? RuntimeDocumentUrls.Base(document) : new Uri(new Uri(RuntimeDocumentUrls.Base(document)), sourceUrl).AbsoluteUri;
        lock (engine) return engine.Evaluate(source, location);
    }

    public async Task EvaluateScriptAsync(IResponse response, ScriptOptions options, CancellationToken cancel) {
        bool isModule = string.Equals(options.Element?.Type, "module", StringComparison.OrdinalIgnoreCase);
        using var reader = new StreamReader(response.Content, isModule ? Encoding.UTF8 : options.Encoding ?? Encoding.UTF8, !isModule);
        string source = await reader.ReadToEndAsync(cancel).ConfigureAwait(false);
        if (!isModule) {
            string sourceUrl = options.Element?.HasAttribute("src") == true ? response.Address.Href : RuntimeDocumentUrls.Base(options.Document);
            await options.EventLoop.EnqueueAsync(_ => EvaluateScript(options.Document, source, options.Element?.Type!, sourceUrl), TaskPriority.Critical).WaitAsync(cancel);
            return;
        }
        var engine = scripting.GetOrCreateJint(options.Document);
        ModuleImportOperation operation = await options.EventLoop.EnqueueAsync(_ => {
            bool external = options.Element!.HasAttribute("src");
            var baseUrl = new Uri(external ? response.Address.Href : RuntimeDocumentUrls.Base(options.Document));
            if (external) RuntimeModuleLoader.Validate((int)response.StatusCode, response.Headers.TryGetValue("Content-Type", out var mime) ? mime : "");
            string identity = modules().Register(source, baseUrl, external ? new Uri(new Uri(RuntimeDocumentUrls.Base(options.Document)), options.Element.Source!).AbsoluteUri : null);
            return engine.Modules.StartImport(identity);
        }, TaskPriority.Critical).WaitAsync(cancel);
        while (true) {
            bool done = await options.EventLoop.EnqueueAsync(_ => {
                engine.Advanced.ProcessTasks();
                if (!operation.IsCompleted) return false;
                operation.GetResult();
                return true;
            }).WaitAsync(cancel);
            if (done) return;
            await Task.Delay(request.PollInterval, cancel);
        }
    }
}
