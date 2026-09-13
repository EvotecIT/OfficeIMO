using System.Text;
using AngleSharp.Browser;
using AngleSharp.Dom;
using AngleSharp.Io;
using AngleSharp.Scripting;
using Jint.Runtime.Modules;
using Jint;
using Jint.Native;
using Jint.Runtime.Interop;

namespace OfficeIMO.Html.Runtime.Worker;

// Module evaluation yields the native event loop while top-level awaits are pending.
// Classic execution and the retained DOM wrappers continue using the same Jint engine.
internal sealed class RuntimeScriptingService(JsScriptingService scripting, Func<RuntimeModuleLoader> modules, HtmlScriptRequest request, Action<string> report) : IScriptingService, IDisposable {
    private JsValue _observeLoad = null!;
    private readonly CancellationTokenSource _lifetime = new();
    private readonly List<Task> _evaluations = [];
    private int _disposed;

    internal void Initialize(Engine engine) => _observeLoad = engine.Evaluate("(() => {const then=Promise.prototype.then,apply=Reflect.apply,define=Object.defineProperty,constructor=Object.freeze({[Symbol.species]:Promise});return (promise,fulfilled,rejected)=>{define(promise,'constructor',{value:constructor});apply(then,promise,[fulfilled,rejected]);};})()");

    private Task ObserveLoad(Engine engine, JsValue promise) {
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var fulfilled = new ClrFunction(engine, "moduleGraphLoaded", (_, _) => { completion.TrySetResult(); return JsValue.Undefined; });
        var rejected = new ClrFunction(engine, "moduleGraphFailed", (_, args) => { completion.TrySetException(new HtmlScriptRuntimeException(args[0].ToString())); return JsValue.Undefined; });
        engine.Invoke(_observeLoad, new JsValue[] { promise, fulfilled, rejected });
        return completion.Task;
    }

    internal async Task WaitForModuleEvaluationsAsync(CancellationToken token) {
        while (true) {
            Task[] pending;
            lock (_evaluations) { _evaluations.RemoveAll(task => task.IsCompleted); pending = _evaluations.ToArray(); }
            if (pending.Length == 0) return;
            await Task.WhenAll(pending).WaitAsync(token);
        }
    }

    public void Dispose() {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        _lifetime.Cancel();
        _lifetime.Dispose();
    }
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
        bool isModule = string.Equals(options.PreparedType, "module", StringComparison.OrdinalIgnoreCase);
        using var reader = new StreamReader(response.Content, isModule ? Encoding.UTF8 : options.Encoding ?? Encoding.UTF8, !isModule);
        string source = await reader.ReadToEndAsync(cancel).ConfigureAwait(false);
        if (!isModule) {
            string sourceUrl = options.IsExternal ? response.Address.Href : RuntimeDocumentUrls.Base(options.Document);
            await options.EventLoop.EnqueueAsync(_ => EvaluateScript(options.Document, source, options.PreparedType!, sourceUrl), TaskPriority.Critical).WaitAsync(cancel);
            return;
        }
        var engine = scripting.GetOrCreateJint(options.Document);
        var prepared = await options.EventLoop.EnqueueAsync(_ => {
            bool external = options.IsExternal;
            var baseUrl = new Uri(response.Address.Href);
            if (external) RuntimeModuleLoader.Validate((int)response.StatusCode, response.Headers.TryGetValue("Content-Type", out var mime) ? mime : "");
            string identity = modules().Register(source, baseUrl, external ? options.PreparedSourceUrl : null);
            return (Identity: identity, Source: modules().GetSource(identity));
        }, TaskPriority.Critical).WaitAsync(cancel);
        var moduleSource = await prepared.Source.WaitAsync(_lifetime.Token).WaitAsync(cancel);
        var loaded = await options.EventLoop.EnqueueAsync(_ => {
            var record = modules().Prepare(engine, prepared.Identity, moduleSource);
            return ObserveLoad(engine, record.LoadRequestedModules());
        }, TaskPriority.Critical).WaitAsync(cancel);
        await loaded.WaitAsync(_lifetime.Token).WaitAsync(cancel);
        ModuleImportOperation operation = await options.EventLoop.EnqueueAsync(_ => engine.Modules.StartImport(prepared.Identity), TaskPriority.Critical).WaitAsync(cancel);
        var evaluation = CompleteEvaluationAsync(operation, engine, options.EventLoop);
        lock (_evaluations) { _evaluations.RemoveAll(task => task.IsCompleted); _evaluations.Add(evaluation); }
    }

    private async Task CompleteEvaluationAsync(ModuleImportOperation operation, Engine engine, IEventLoop loop) {
        try {
            while (true) {
                bool done = await loop.EnqueueAsync(_ => {
                    engine.Advanced.ProcessTasks();
                    if (!operation.IsCompleted) return false;
                    operation.GetResult();
                    return true;
                }).WaitAsync(_lifetime.Token);
                if (done) return;
                await Task.Delay(request.PollInterval, _lifetime.Token);
            }
        } catch (OperationCanceledException) when (_lifetime.IsCancellationRequested) { }
        catch (Exception error) { report(error.Message); }
    }
}
