using System.Net.Http.Headers;
using AngleSharp.Browser;
using AngleSharp.Dom;
using Jint;
using Jint.Runtime;
using Jint.Runtime.Modules;

namespace OfficeIMO.Html.Runtime.Worker;

// Module identity belongs to the interpreter; source loading uses the same bounded
// transport as fetch. Async completions return to the owning native event loop.
internal sealed class RuntimeModuleLoader(IDocument document, RuntimeModuleSourceCache sources, Func<RuntimeEventLoop> loop,
    Func<Engine> getEngine, Func<CancellationToken> realmLifetime) : IAsyncModuleLoader {
    private readonly Dictionary<ModuleIdentity, Jint.Runtime.Modules.Module> _records = [];
    private readonly HashSet<string> _preparedRootIntegrity = new(StringComparer.Ordinal);
    internal RuntimeImportMap ImportMap { get; } = new();

    internal Jint.Runtime.Modules.Module Prepare(Engine engine, string identity, RuntimeModuleSource source) =>
        Record(engine, new ResolvedSpecifier(new ModuleRequest(identity, []), identity,
            new Uri(identity), SpecifierType.RelativeOrAbsolute), source);

    private Jint.Runtime.Modules.Module Record(Engine engine, ResolvedSpecifier resolved, RuntimeModuleSource source) {
        ModuleKind kind = Kind(resolved.ModuleRequest);
        var identity = new ModuleIdentity(resolved.Key, kind);
        if (!_records.TryGetValue(identity, out var record)) {
            Validate(source.StatusCode, source.ContentType, kind);
            var loaded = new ResolvedSpecifier(resolved.ModuleRequest, source.Location,
                new Uri(source.Location), SpecifierType.RelativeOrAbsolute);
            record = kind == ModuleKind.Json
                ? ModuleFactory.BuildJsonModule(engine, loaded, source.Text)
                : ModuleFactory.BuildSourceTextModule(engine, loaded, source.Text);
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
            _ = Kind(moduleRequest);
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
        CancellationToken lifetime = realmLifetime();
        var pending = sources.GetOrLoad(resolved.Key, resolved.Uri!, integrityMetadata, lifetime);
        if (pending.IsCompleted) Settle(pending, engine, resolved, completion);
        else _ = CompleteAsync(pending, engine, resolved, completion, loop(), lifetime);
    }

    private async Task CompleteAsync(Task<RuntimeModuleSource> pending, Engine engine,
        ResolvedSpecifier resolved, ModuleLoadCompletion completion, RuntimeEventLoop ownerLoop, CancellationToken lifetime) {
        try { await pending.ConfigureAwait(false); } catch { /* Settle transports the original failure on the engine thread. */ }
        ownerLoop.TryEnqueue(_ => Settle(pending, engine, resolved, completion), TaskPriority.Normal, lifetime);
    }

    private void Settle(Task<RuntimeModuleSource> pending, Engine engine, ResolvedSpecifier resolved, ModuleLoadCompletion completion) {
        try {
            RuntimeModuleSource source = pending.GetAwaiter().GetResult();
            completion.SetModule(Record(engine, resolved, source));
        } catch (Exception error) {
            completion.SetError(error is JavaScriptException
                ? error
                : new JavaScriptException(engine.Intrinsics.TypeError, error.Message));
        }
    }

    internal static void ValidateStatus(int status) {
        if (status < 200 || status >= 300) throw new HtmlScriptRuntimeException("Module loading returned HTTP " + status + ".");
    }

    internal static void ValidateJavaScript(int status, string contentType) =>
        Validate(status, contentType, ModuleKind.JavaScript);

    private static void Validate(int status, string contentType, ModuleKind kind) {
        ValidateStatus(status);
        if (!MediaTypeHeaderValue.TryParse(contentType, out MediaTypeHeaderValue? parsed)
            || string.IsNullOrWhiteSpace(parsed.MediaType))
            throw new HtmlScriptRuntimeException("The module response does not have a valid MIME type.");
        string mime = parsed.MediaType.ToLowerInvariant();
        if (kind == ModuleKind.Json) {
            if (mime is "application/json" or "text/json" || mime.EndsWith("+json", StringComparison.Ordinal)) return;
            throw new HtmlScriptRuntimeException("The JSON module response does not have a JSON MIME type.");
        }
        if (mime is not ("text/javascript" or "application/javascript" or "application/ecmascript" or "text/ecmascript" or
            "application/x-ecmascript" or "application/x-javascript" or "text/javascript1.0" or "text/javascript1.1" or
            "text/javascript1.2" or "text/javascript1.3" or "text/javascript1.4" or "text/javascript1.5" or "text/jscript" or
            "text/livescript" or "text/x-ecmascript" or "text/x-javascript"))
            throw new HtmlScriptRuntimeException("The module response does not have a JavaScript MIME type.");
    }

    private static ModuleKind Kind(ModuleRequest request) {
        if (request.Attributes.Length == 0) return ModuleKind.JavaScript;
        if (request.Attributes.Length == 1 && request.Attributes[0] is { Key: "type", Value: "json" })
            return ModuleKind.Json;
        throw new HtmlScriptRuntimeException("This runtime supports the 'type: json' import attribute only.");
    }

    private readonly record struct ModuleIdentity(string Url, ModuleKind Kind);
    private enum ModuleKind { JavaScript, Json }
}
