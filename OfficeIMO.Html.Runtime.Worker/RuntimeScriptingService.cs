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
internal sealed class RuntimeScriptingService(JsScriptingService scripting, Func<IDocument, RuntimeModuleLoader?> modules,
    Func<IDocument, CancellationToken> realmLifetime, Func<IDocument, Engine?> ensureEngine, object sessionSync,
    HtmlScriptRequest request, Action<string> report, RuntimeScriptEntry scriptEntry) : IScriptingService, ISynchronousScriptingService, IDisposable {
    private readonly Dictionary<Engine, JsValue> _observeLoad = new(ReferenceEqualityComparer.Instance);
    private readonly CancellationTokenSource _lifetime = new();
    private readonly List<(IDocument Document, Task Task)> _evaluations = [];
    private int _disposed;

    internal void Initialize(Engine engine) => _observeLoad.Add(engine,
        engine.Evaluate("(() => {const then=Promise.prototype.then,apply=Reflect.apply,define=Object.defineProperty,constructor=Object.freeze({[Symbol.species]:Promise});return (promise,fulfilled,rejected)=>{define(promise,'constructor',{value:constructor});apply(then,promise,[fulfilled,rejected]);};})()"));

    private Task ObserveLoad(Engine engine, JsValue promise) {
        var completion = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        var fulfilled = new ClrFunction(engine, "moduleGraphLoaded", (_, _) => { completion.TrySetResult(); return JsValue.Undefined; });
        var rejected = new ClrFunction(engine, "moduleGraphFailed", (_, args) => { completion.TrySetException(new HtmlScriptRuntimeException(args[0].ToString())); return JsValue.Undefined; });
        engine.Invoke(_observeLoad[engine], new JsValue[] { promise, fulfilled, rejected });
        return completion.Task;
    }

    internal async Task WaitForModuleEvaluationsAsync(IDocument root, CancellationToken token) {
        while (true) {
            Task[] pending;
            lock (_evaluations) { _evaluations.RemoveAll(item => item.Task.IsCompleted); pending = _evaluations.Where(item => BelongsToRoot(item.Document, root)).Select(item => item.Task).ToArray(); }
            if (pending.Length == 0) return;
            await Task.WhenAll(pending).WaitAsync(token);
        }
    }

    private static bool BelongsToRoot(IDocument document, IDocument root) {
        for (var context = document.Context; context != null; context = context.Parent) {
            if (ReferenceEquals(context, root.Context)) return true;
            if (!RuntimeFrameRealms.IsFrameContext(context)) return false;
        }
        return false;
    }

    internal void Retire(Engine engine) => _observeLoad.Remove(engine);

    public void Dispose() {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        _lifetime.Cancel();
        _observeLoad.Clear();
        _lifetime.Dispose();
    }
    public string Type => scripting.Type;
    public bool SupportsType(string type) => ((IScriptingService)scripting).SupportsType(type);

    public object EvaluateScript(IDocument document, string source, string type, string sourceUrl) {
        var engine = ensureEngine(document);
        if (engine == null) return null!;
        RuntimeModuleLoader loader = modules(document)
            ?? throw new HtmlScriptRuntimeException("The document module realm is no longer active.");
        if (type?.Equals("importmap", StringComparison.OrdinalIgnoreCase) == true) {
            loader.ImportMap.Load(source, new Uri(RuntimeDocumentUrls.Base(document)));
            return null!;
        }
        if (type?.Equals("module", StringComparison.OrdinalIgnoreCase) == true)
            throw new HtmlScriptRuntimeException("Module scripts require asynchronous document execution.");
        string location = string.IsNullOrEmpty(sourceUrl) ? RuntimeDocumentUrls.Base(document) : new Uri(new Uri(RuntimeDocumentUrls.Base(document)), sourceUrl).AbsoluteUri;
        lock (sessionSync) {
            using var entry = scriptEntry.Enter(document);
            return engine.Evaluate(source, location);
        }
    }

    public async Task EvaluateScriptAsync(IResponse response, ScriptOptions options, CancellationToken cancel) {
        bool isModule = string.Equals(options.PreparedType, "module", StringComparison.OrdinalIgnoreCase);
        Engine? engine = ensureEngine(options.Document);
        if (engine == null) return;
        RuntimeModuleLoader loader = modules(options.Document)
            ?? throw new HtmlScriptRuntimeException("The document module realm is no longer active.");
        if (!isModule) {
            using var reader = new StreamReader(response.Content, options.Encoding ?? Encoding.UTF8, true);
            string classicSource = await reader.ReadToEndAsync(cancel).ConfigureAwait(false);
            string sourceUrl = options.IsExternal ? response.Address.Href : RuntimeDocumentUrls.Base(options.Document);
            await options.EventLoop.EnqueueAsync(_ => {
                cancel.ThrowIfCancellationRequested();
                return EvaluateScript(options.Document, classicSource, options.PreparedType!, sourceUrl);
            }, TaskPriority.Critical).WaitAsync(cancel);
            return;
        }
        using var content = new MemoryStream();
        await response.Content.CopyToAsync(content, 81920, cancel).ConfigureAwait(false);
        byte[] buffer = content.ToArray();
        string source = Encoding.UTF8.GetString(buffer);
        if (source.StartsWith('\uFEFF')) source = source[1..];
        var prepared = await options.EventLoop.EnqueueAsync(_ => {
            bool external = options.IsExternal;
            var baseUrl = new Uri(response.Address.Href);
            string contentType = response.Headers.TryGetValue("Content-Type", out var mime) ? mime : "text/javascript";
            if (external) RuntimeModuleLoader.ValidateJavaScript((int)response.StatusCode, contentType);
            string? integrityMetadata = null;
            if (external) {
                if (options.PreparedIntegritySnapshot is not { IsResolved: true } snapshot)
                    throw new HtmlScriptRuntimeException("The module integrity selection was not completed during resource loading.");
                integrityMetadata = snapshot.Value;
            }
            return loader.Register(source, buffer, baseUrl, external ? options.PreparedSourceUrl : null,
                integrityMetadata, contentType, (int)response.StatusCode, response.Headers,
                external);
        }, TaskPriority.Critical).WaitAsync(cancel);
        using var loading = CancellationTokenSource.CreateLinkedTokenSource(_lifetime.Token,
            realmLifetime(options.Document), cancel);
        var moduleSource = await prepared.Source.WaitAsync(loading.Token);
        var loaded = await options.EventLoop.EnqueueAsync(_ => {
            var record = loader.Prepare(engine, prepared.Identity, moduleSource);
            return ObserveLoad(engine, record.LoadRequestedModules());
        }, TaskPriority.Critical).WaitAsync(loading.Token);
        await loaded.WaitAsync(loading.Token);
        ModuleImportOperation operation = await options.EventLoop.EnqueueAsync(
            _ => engine.Modules.StartImport(prepared.Identity), TaskPriority.Critical).WaitAsync(loading.Token);
        var evaluation = CompleteEvaluationAsync(operation, engine, options.EventLoop, realmLifetime(options.Document));
        lock (_evaluations) { _evaluations.RemoveAll(item => item.Task.IsCompleted); _evaluations.Add((options.Document, evaluation)); }
    }

    private async Task CompleteEvaluationAsync(ModuleImportOperation operation, Engine engine, IEventLoop loop,
        CancellationToken realmToken) {
        using var lifetime = CancellationTokenSource.CreateLinkedTokenSource(_lifetime.Token, realmToken);
        try {
            while (true) {
                bool done = await loop.EnqueueAsync(_ => {
                    engine.Advanced.ProcessTasks();
                    if (!operation.IsCompleted) return false;
                    operation.GetResult();
                    return true;
                }).WaitAsync(lifetime.Token);
                if (done) return;
                await Task.Delay(request.PollInterval, lifetime.Token);
            }
        } catch (OperationCanceledException) when (lifetime.IsCancellationRequested) { }
        catch (Exception error) { report(error.Message); }
    }
}
