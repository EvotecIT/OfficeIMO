using AngleSharp.Browser;
using AngleSharp.Dom;
using Jint;
using Jint.Runtime;
using Jint.Runtime.Modules;

namespace OfficeIMO.Html.Runtime.Worker;

// Module identity belongs to the interpreter; source loading uses the same bounded
// transport as fetch. Async completions return to the owning native event loop.
internal sealed class RuntimeModuleLoader(IDocument document, RuntimeModuleSourceCache sources, Func<IEventLoop> loop, Func<Engine> getEngine) : IAsyncModuleLoader {
    private readonly Dictionary<string, Jint.Runtime.Modules.Module> _records = new(StringComparer.Ordinal);
    private readonly HashSet<string> _preparedRootIntegrity = new(StringComparer.Ordinal);
    internal RuntimeImportMap ImportMap { get; } = new();

    internal Jint.Runtime.Modules.Module Prepare(Engine engine, string identity, RuntimeModuleSource source) =>
        Record(engine, identity, source);

    private Jint.Runtime.Modules.Module Record(Engine engine, string identity, RuntimeModuleSource source) {
        if (!_records.TryGetValue(identity, out var record)) {
            var resolved = new ResolvedSpecifier(new ModuleRequest(identity, []), source.Location, new Uri(source.Location), SpecifierType.RelativeOrAbsolute);
            record = ModuleFactory.BuildSourceTextModule(engine, resolved, source.Text);
            _records.Add(identity, record);
        }
        return record;
    }

    internal (string Identity, Task<RuntimeModuleSource> Source) Register(string source, byte[] buffer, Uri baseUrl,
        string? externalIdentity, string? integrityMetadata, string contentType, int statusCode,
        IEnumerable<KeyValuePair<string, string>> headers, bool hasPreparedIntegritySelection) {
        string identity = externalIdentity ?? new UriBuilder(baseUrl) { Fragment = "officeimo-inline-" + Guid.NewGuid().ToString("N") }.Uri.AbsoluteUri;
        if (hasPreparedIntegritySelection) _preparedRootIntegrity.Add(identity);
        return (identity, sources.Register(identity, new(source, baseUrl.AbsoluteUri, buffer, contentType, statusCode,
            new Dictionary<string, string>(headers, StringComparer.OrdinalIgnoreCase)), integrityMetadata));
    }

    internal string ResolveForImportMeta(string specifier, string referrer) {
        try {
            Uri url = ImportMap.Resolve(specifier, new Uri(referrer));
            HtmlRuntimeResourcePolicy.ValidateUrl(url);
            return url.AbsoluteUri;
        } catch (Exception error) when (error is ArgumentException or UriFormatException or HtmlScriptRuntimeException) {
            throw new JavaScriptException(getEngine().Intrinsics.TypeError, error.Message);
        }
    }

    public ResolvedSpecifier Resolve(string? referencingModuleLocation, ModuleRequest moduleRequest) {
        try {
            if (moduleRequest.Attributes.Length != 0) throw new HtmlScriptRuntimeException("Only JavaScript modules without import attributes are supported.");
            Uri url = referencingModuleLocation == null && Uri.TryCreate(moduleRequest.Specifier, UriKind.Absolute, out var root) ? root
                : ImportMap.Resolve(moduleRequest.Specifier, Uri.TryCreate(referencingModuleLocation, UriKind.Absolute, out var referrer) ? referrer : new Uri(RuntimeDocumentUrls.Base(document)));
            HtmlRuntimeResourcePolicy.ValidateUrl(url);
            return new ResolvedSpecifier(moduleRequest, url.AbsoluteUri, url, SpecifierType.RelativeOrAbsolute);
        } catch (Exception error) when (error is ArgumentException or UriFormatException or HtmlScriptRuntimeException) {
            throw new JavaScriptException(getEngine().Intrinsics.TypeError, error.Message);
        }
    }

    public Jint.Runtime.Modules.Module LoadModule(Engine engine, ResolvedSpecifier resolved) =>
        throw new HtmlScriptRuntimeException("Module loading requires the asynchronous runtime path.");

    public void LoadModuleAsync(Engine engine, ResolvedSpecifier resolved, ModuleLoadCompletion completion) {
        string? integrityMetadata = _preparedRootIntegrity.Contains(resolved.Key) ? null : ImportMap.IntegrityFor(resolved.Uri!);
        var pending = sources.GetOrLoad(resolved.Key, resolved.Uri!, integrityMetadata);
        if (pending.IsCompleted) Settle(pending, engine, resolved, completion);
        else _ = CompleteAsync(pending, engine, resolved, completion);
    }

    private async Task CompleteAsync(Task<RuntimeModuleSource> pending, Engine engine, ResolvedSpecifier resolved, ModuleLoadCompletion completion) {
        try { await pending.ConfigureAwait(false); } catch { /* Settle transports the original failure on the engine thread. */ }
        loop().Enqueue(_ => Settle(pending, engine, resolved, completion), TaskPriority.Normal);
    }

    private void Settle(Task<RuntimeModuleSource> pending, Engine engine, ResolvedSpecifier resolved, ModuleLoadCompletion completion) {
        try {
            RuntimeModuleSource source = pending.GetAwaiter().GetResult();
            completion.SetModule(Record(engine, resolved.Key, source));
        } catch (Exception error) { completion.SetError(error); }
    }

    internal static void Validate(int status, string contentType) {
        if (status < 200 || status >= 300) throw new HtmlScriptRuntimeException("Module loading returned HTTP " + status + ".");
        string mime = contentType.Split(';', 2)[0].Trim().ToLowerInvariant();
        if (mime is not ("text/javascript" or "application/javascript" or "application/ecmascript" or "text/ecmascript" or
            "application/x-ecmascript" or "application/x-javascript" or "text/javascript1.0" or "text/javascript1.1" or
            "text/javascript1.2" or "text/javascript1.3" or "text/javascript1.4" or "text/javascript1.5" or "text/jscript" or
            "text/livescript" or "text/x-ecmascript" or "text/x-javascript"))
            throw new HtmlScriptRuntimeException("The module response does not have a JavaScript MIME type.");
    }
}
