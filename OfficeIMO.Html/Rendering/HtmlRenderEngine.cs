using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// First-party dependency-free HTML layout entry point shared by image and PDF adapters.
/// </summary>
public static class HtmlRenderEngine {
    /// <summary>
    /// Parses and renders HTML into a backend-neutral continuous or paged visual document.
    /// </summary>
    public static HtmlRenderDocument Render(string html, HtmlRenderOptions? options = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        resolved.Validate();
        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
        var diagnostics = new HtmlDiagnosticReport();
        var resourceOptions = new HtmlResourcePipelineOptions {
            BaseUri = resolved.BaseUri,
            UrlPolicy = (resolved.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
            MediaContext = resolved.MediaContext
        };
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(document, resourceOptions);
        diagnostics.AddRange(manifest.Diagnostics.Diagnostics);
        var resources = new HtmlRenderResourceSet();
        AddPendingStylesheetDiagnostics(manifest, resources, diagnostics);
        HtmlCssPageRuleSet pageRules = HtmlCssPageSettingsResolver.Apply(document, resolved, diagnostics);
        resolved.Validate();
        IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document, resolved.MediaContext);
        return new HtmlRenderLayoutEngine(document, styles, resolved, diagnostics, resources, pageRules).Render();
    }

    /// <summary>
    /// Parses and renders HTML while asynchronously resolving policy-approved external resources through the configured resolver.
    /// </summary>
    public static async Task<HtmlRenderDocument> RenderAsync(string html, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        resolved.Validate();
        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
        var diagnostics = new HtmlDiagnosticReport();
        var resourceOptions = new HtmlResourcePipelineOptions {
            BaseUri = resolved.BaseUri,
            UrlPolicy = (resolved.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
            MediaContext = resolved.MediaContext
        };
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(document, resourceOptions);
        diagnostics.AddRange(manifest.Diagnostics.Diagnostics);
        HtmlRenderResourceSet resources = await HtmlRenderResourceLoader.LoadAsync(manifest, resolved, diagnostics, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderStylesheetApplier.Apply(document, resources, diagnostics);
        AddPendingStylesheetDiagnostics(manifest, resources, diagnostics);
        HtmlCssPageRuleSet pageRules = HtmlCssPageSettingsResolver.Apply(document, resolved, diagnostics);
        resolved.Validate();
        IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(document, resolved.MediaContext);
        return new HtmlRenderLayoutEngine(document, styles, resolved, diagnostics, resources, pageRules).Render();
    }

    /// <summary>
    /// Renders HTML through the shared first-party layout engine.
    /// </summary>
    public static HtmlRenderDocument RenderHtml(this string html, HtmlRenderOptions? options = null) => Render(html, options);

    private static void AddPendingStylesheetDiagnostics(HtmlResourceManifest manifest, HtmlRenderResourceSet resources, HtmlDiagnosticReport diagnostics) {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (HtmlResourceReference reference in manifest.Resources) {
            if (!reference.IsAllowed
                || reference.Kind != HtmlResourceKind.Stylesheet
                || reference.ResolvedSource.Length == 0
                || resources.TryGet(reference.Source, reference.ResolvedSource, out _)
                || resources.WasAttempted(reference.Source, reference.ResolvedSource)
                || !seen.Add(reference.ResolvedSource)) {
                continue;
            }

            diagnostics.Add(
                "OfficeIMO.Html.Renderer",
                HtmlRenderDiagnosticCodes.ExternalStylesheetPending,
                "An external stylesheet was not loaded; use the asynchronous renderer with a resource resolver.",
                HtmlDiagnosticSeverity.Warning,
                reference.Source,
                reference.ResolvedSource);
        }
    }
}
