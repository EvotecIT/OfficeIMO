using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// First-party dependency-free HTML layout entry point shared by image and PDF adapters.
/// </summary>
public static class HtmlRenderEngine {
    // Raw text entry points remain internal for renderer-focused tests and low-level package code.
    // They delegate to the native source model so parsing, normalization, trust, and media filtering
    // still have one owner.
    internal static HtmlRenderDocument Render(string html, HtmlRenderOptions? options = null) =>
        Render(HtmlConversionDocument.Parse(html), options);

    /// <summary>
    /// Renders a parsed HTML source into a backend-neutral continuous or paged visual document.
    /// </summary>
    public static HtmlRenderDocument Render(
        HtmlConversionDocument document,
        HtmlRenderOptions? options = null) =>
        Render(document, options, CancellationToken.None);

    internal static HtmlRenderDocument Render(
        HtmlConversionDocument document,
        HtmlRenderOptions? options,
        CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        resolved.BaseUri ??= document.BaseUri;
        resolved.Validate();
        HtmlRenderInputGuard.ValidateSource(document.SourceHtml, resolved);
        cancellationToken.ThrowIfCancellationRequested();
        IHtmlDocument renderDocument = document.CreateDocumentForRendering();
        return RenderDocument(
            renderDocument,
            resolved,
            initialDiagnostics: null,
            document.Limits,
            cancellationToken);
    }

    /// <summary>
    /// Renders a prepared HTML DOM without reparsing source text or mutating the caller's document.
    /// </summary>
    internal static HtmlRenderDocument Render(IHtmlDocument document, HtmlRenderOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        resolved.Validate();
        IHtmlDocument renderDocument = HtmlDocumentParser.CloneDocument(document);
        HtmlRenderInputGuard.ValidateSource(renderDocument.DocumentElement?.OuterHtml ?? string.Empty, resolved);
        return RenderDocument(
            renderDocument,
            resolved,
            initialDiagnostics: null,
            HtmlConversionLimits.CreateUntrustedProfile(),
            CancellationToken.None);
    }

    private static HtmlRenderDocument RenderDocument(
        IHtmlDocument document,
        HtmlRenderOptions resolved,
        IEnumerable<HtmlDiagnostic>? initialDiagnostics,
        HtmlConversionLimits limits,
        CancellationToken cancellationToken) {
        resolved.ResponsiveImageCandidateLimit = limits.MaxResponsiveImageCandidates;
        HtmlRenderInputGuard.ValidateDocument(document, resolved, cancellationToken);
        var diagnostics = new HtmlDiagnosticReport();
        if (initialDiagnostics != null) diagnostics.AddRange(initialDiagnostics);
        var resourceOptions = new HtmlResourcePipelineOptions {
            BaseUri = resolved.BaseUri,
            UrlPolicy = (resolved.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
            ResourceUrlPolicy = resolved.GetResourceUrlPolicy().Clone(),
            Limits = limits.Clone(),
            MaxResponsiveImageCandidates = resolved.ResponsiveImageCandidateLimit,
            MediaContext = resolved.MediaContext,
            MediaWidth = resolved.Mode == HtmlRenderMode.Paged ? resolved.PageWidth : resolved.ViewportWidth,
            MediaHeight = resolved.Mode == HtmlRenderMode.Paged ? resolved.PageHeight : resolved.ViewportHeight ?? 1056D
        };
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(document, resourceOptions);
        cancellationToken.ThrowIfCancellationRequested();
        diagnostics.AddRange(manifest.Diagnostics);
        HtmlCssByteBudget cssBudget = HtmlRenderStylesheetApplier.CreateBudget(document, limits);
        HtmlResourceSession resources = HtmlRenderResourceLoader.Load(
            manifest,
            resolved,
            diagnostics,
            cancellationToken,
            cssBudget);
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderStylesheetApplier.Apply(document, resources, resolved, cssBudget, diagnostics);
        AddPendingStylesheetDiagnostics(manifest, resources, diagnostics);
        OfficeIMO.Drawing.OfficeFontFaceCollection fonts = HtmlRenderFontFaceLoader.Load(document, resources, resolved, diagnostics);
        fonts.AddRange(resolved.Fonts);
        HtmlCssPageRuleSet pageRules = HtmlCssPageSettingsResolver.Apply(document, resolved, diagnostics);
        resolved.Validate();
        HtmlComputedStyleSet styles = HtmlComputedStyleEngine.ComputeForRendering(document, resolved, limits);
        cancellationToken.ThrowIfCancellationRequested();
        return new HtmlRenderLayoutEngine(
            document,
            styles,
            resolved,
            diagnostics,
            resources,
            pageRules,
            fonts,
            cancellationToken).Render();
    }

    /// <summary>
    /// Renders a parsed HTML source while asynchronously resolving policy-approved external resources through the configured resolver.
    /// </summary>
    public static async Task<HtmlRenderDocument> RenderAsync(HtmlConversionDocument document, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        resolved.BaseUri ??= document.BaseUri;
        resolved.Validate();
        HtmlRenderInputGuard.ValidateSource(document.SourceHtml, resolved);
        IHtmlDocument renderDocument = document.CreateDocumentForRendering();
        return await RenderDocumentAsync(
            renderDocument,
            resolved,
            initialDiagnostics: null,
            document.Limits,
            cancellationToken).ConfigureAwait(false);
    }

    internal static Task<HtmlRenderDocument> RenderAsync(string html, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) =>
        RenderAsync(HtmlConversionDocument.Parse(html), options, cancellationToken);

    /// <summary>
    /// Renders a prepared HTML DOM while resolving resources without reparsing or mutating the caller's document.
    /// </summary>
    internal static async Task<HtmlRenderDocument> RenderAsync(IHtmlDocument document, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        resolved.Validate();
        IHtmlDocument renderDocument = HtmlDocumentParser.CloneDocument(document);
        HtmlRenderInputGuard.ValidateSource(renderDocument.DocumentElement?.OuterHtml ?? string.Empty, resolved);
        return await RenderDocumentAsync(
            renderDocument,
            resolved,
            initialDiagnostics: null,
            HtmlConversionLimits.CreateUntrustedProfile(),
            cancellationToken).ConfigureAwait(false);
    }

    private static async Task<HtmlRenderDocument> RenderDocumentAsync(
        IHtmlDocument document,
        HtmlRenderOptions resolved,
        IEnumerable<HtmlDiagnostic>? initialDiagnostics,
        HtmlConversionLimits limits,
        CancellationToken cancellationToken) {
        resolved.ResponsiveImageCandidateLimit = limits.MaxResponsiveImageCandidates;
        HtmlRenderInputGuard.ValidateDocument(document, resolved, cancellationToken);
        var diagnostics = new HtmlDiagnosticReport();
        if (initialDiagnostics != null) diagnostics.AddRange(initialDiagnostics);
        var resourceOptions = new HtmlResourcePipelineOptions {
            BaseUri = resolved.BaseUri,
            UrlPolicy = (resolved.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
            ResourceUrlPolicy = resolved.GetResourceUrlPolicy().Clone(),
            Limits = limits.Clone(),
            MaxResponsiveImageCandidates = resolved.ResponsiveImageCandidateLimit,
            MediaContext = resolved.MediaContext,
            MediaWidth = resolved.Mode == HtmlRenderMode.Paged ? resolved.PageWidth : resolved.ViewportWidth,
            MediaHeight = resolved.Mode == HtmlRenderMode.Paged ? resolved.PageHeight : resolved.ViewportHeight ?? 1056D
        };
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(document, resourceOptions);
        diagnostics.AddRange(manifest.Diagnostics);
        HtmlCssByteBudget cssBudget = HtmlRenderStylesheetApplier.CreateBudget(document, limits);
        HtmlResourceSession resources = await HtmlRenderResourceLoader.LoadAsync(manifest, resolved, diagnostics, cancellationToken, cssBudget).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderStylesheetApplier.Apply(document, resources, resolved, cssBudget, diagnostics);
        AddPendingStylesheetDiagnostics(manifest, resources, diagnostics);
        OfficeIMO.Drawing.OfficeFontFaceCollection fonts = HtmlRenderFontFaceLoader.Load(document, resources, resolved, diagnostics);
        fonts.AddRange(resolved.Fonts);
        HtmlCssPageRuleSet pageRules = HtmlCssPageSettingsResolver.Apply(document, resolved, diagnostics);
        cancellationToken.ThrowIfCancellationRequested();
        resolved.Validate();
        HtmlComputedStyleSet styles = HtmlComputedStyleEngine.ComputeForRendering(document, resolved, limits);
        cancellationToken.ThrowIfCancellationRequested();
        return new HtmlRenderLayoutEngine(document, styles, resolved, diagnostics, resources, pageRules, fonts, cancellationToken).Render();
    }

    internal static HtmlRenderDocument RenderHtml(this string html, HtmlRenderOptions? options = null) => Render(html, options);

    internal static Task<HtmlRenderDocument> RenderHtmlAsync(this string html, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) =>
        RenderAsync(html, options, cancellationToken);

    private static void AddPendingStylesheetDiagnostics(HtmlResourceManifest manifest, HtmlResourceSession resources, HtmlDiagnosticReport diagnostics) {
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
