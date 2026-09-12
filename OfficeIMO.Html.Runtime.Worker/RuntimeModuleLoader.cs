using System.Text;
using AngleSharp.Browser;
using AngleSharp.Dom;
using Jint;
using Jint.Runtime;
using Jint.Runtime.Modules;

namespace OfficeIMO.Html.Runtime.Worker;

// Module identity belongs to the interpreter; source loading uses the same bounded
// transport as fetch. Async completions return to the owning native event loop.
internal sealed class RuntimeModuleLoader(IDocument document, RuntimeResourceLoader resources, Func<IEventLoop> loop, Func<Engine> getEngine, int maximum) : IAsyncModuleLoader {
    private readonly Dictionary<string, Task<Source>> _sources = new(StringComparer.Ordinal);
    internal RuntimeImportMap ImportMap { get; } = new();

    internal string Register(string source, Uri baseUrl, string? externalIdentity) {
        ImportMap.Seal();
        string identity = externalIdentity ?? new UriBuilder(baseUrl) { Fragment = "officeimo-inline-" + Guid.NewGuid().ToString("N") }.Uri.AbsoluteUri;
        if (!_sources.ContainsKey(identity)) { Reserve(); _sources.Add(identity, Task.FromResult(new Source(source, baseUrl.AbsoluteUri))); }
        return identity;
    }

    public ResolvedSpecifier Resolve(string? referencingModuleLocation, ModuleRequest moduleRequest) {
        try {
            if (moduleRequest.Attributes.Length != 0) throw new HtmlScriptRuntimeException("Only JavaScript modules without import attributes are supported.");
            Uri url = referencingModuleLocation == null && _sources.ContainsKey(moduleRequest.Specifier) ? new Uri(moduleRequest.Specifier)
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
        if (!_sources.TryGetValue(resolved.Key, out var pending)) {
            Reserve();
            pending = LoadAsync(resolved.Uri!);
            _sources.Add(resolved.Key, pending);
        }
        if (pending.IsCompleted) Settle(pending, engine, resolved, completion);
        else _ = CompleteAsync(pending, engine, resolved, completion);
    }

    private void Reserve() {
        if (_sources.Count >= maximum) throw new HtmlScriptRuntimeException("The module source count budget was exceeded.");
    }

    private async Task<Source> LoadAsync(Uri url) {
        var resource = await resources.FetchAsync(url, new RuntimeFetchRequest(), CancellationToken.None).ConfigureAwait(false);
        Validate(resource.StatusCode, resource.ContentType);
        string source = Encoding.UTF8.GetString(resource.Buffer);
        if (source.StartsWith('\uFEFF')) source = source[1..];
        // The response URL is the base for imports and import.meta.url; the module
        // cache remains keyed by the requested URL, including its fragment.
        return new Source(source, resource.FinalUrl.AbsoluteUri);
    }

    private async Task CompleteAsync(Task<Source> pending, Engine engine, ResolvedSpecifier resolved, ModuleLoadCompletion completion) {
        try { await pending.ConfigureAwait(false); } catch { /* Settle transports the original failure on the engine thread. */ }
        loop().Enqueue(_ => Settle(pending, engine, resolved, completion), TaskPriority.Normal);
    }

    private static void Settle(Task<Source> pending, Engine engine, ResolvedSpecifier resolved, ModuleLoadCompletion completion) {
        try {
            Source source = pending.GetAwaiter().GetResult();
            var location = resolved with { Key = source.Location, Uri = new Uri(source.Location) };
            completion.SetModule(ModuleFactory.BuildSourceTextModule(engine, location, source.Text));
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

    private sealed record Source(string Text, string Location);
}
