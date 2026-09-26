using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// First-party browser-free HTML layout entry point shared by image and PDF adapters.
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
        CancellationToken cancellationToken) =>
        Execute(
            document,
            HtmlRenderRequest.FromLegacy(options, HtmlRenderEncoder.DisplayList, HtmlRenderPageSet.All()),
            cancellationToken).Document;

    /// <summary>
    /// Resolves an explicit render request into a backend-neutral display list. The intended output
    /// adapter cannot change the request's CSS, layout, pagination, or page-set axes.
    /// </summary>
    public static HtmlRenderResult Execute(
        HtmlConversionDocument document,
        HtmlRenderRequest request,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (request == null) throw new ArgumentNullException(nameof(request));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderOptions resolved = PrepareOptions(document, request);
        return ExecuteWithDeadline(resolved, cancellationToken,
            operationCancellationToken => ExecuteCore(document, request, resolved, operationCancellationToken));
    }

    /// <summary>
    /// Renders a prepared HTML DOM without reparsing source text or mutating the caller's document.
    /// </summary>
    internal static HtmlRenderDocument Render(IHtmlDocument document, HtmlRenderOptions? options = null) =>
        Render(document, options, HtmlConversionLimits.CreateUntrustedProfile(),
            sourceAlreadyValidated: false, cancellationToken: CancellationToken.None);

    /// <summary>
    /// Renders a prepared DOM that already passed the owning conversion document's source checks.
    /// The supplied limits remain authoritative for DOM and CSS complexity during the render pass.
    /// </summary>
    internal static HtmlRenderDocument Render(
        IHtmlDocument document,
        HtmlRenderOptions? options,
        HtmlConversionLimits limits) =>
        Render(document, options, limits, sourceAlreadyValidated: true,
            cancellationToken: CancellationToken.None);

    /// <summary>
    /// Renders a prepared DOM under the owning conversion document's base URI, URL policies,
    /// and validated complexity limits.
    /// </summary>
    internal static HtmlRenderDocument Render(
        IHtmlDocument document,
        HtmlRenderOptions? options,
        HtmlConversionDocument owningDocument) {
        if (owningDocument == null) throw new ArgumentNullException(nameof(owningDocument));
        return Render(document, options, owningDocument, owningDocument.Limits);
    }

    /// <summary>
    /// Renders a prepared DOM under its owning document policies and an adapter-intersected limit profile.
    /// </summary>
    internal static HtmlRenderDocument Render(
        IHtmlDocument document,
        HtmlRenderOptions? options,
        HtmlConversionDocument owningDocument,
        HtmlConversionLimits limits) {
        if (owningDocument == null) throw new ArgumentNullException(nameof(owningDocument));
        if (limits == null) throw new ArgumentNullException(nameof(limits));
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        resolved.BaseUri ??= owningDocument.BaseUri;
        ApplyDocumentPolicies(owningDocument, resolved);
        return Render(document, resolved, HtmlConversionLimits.Intersect(owningDocument.Limits, limits));
    }

    internal static HtmlRenderDocument Render(
        IHtmlDocument document,
        HtmlRenderOptions? options,
        CancellationToken cancellationToken) =>
        Render(document, options, HtmlConversionLimits.CreateUntrustedProfile(),
            sourceAlreadyValidated: false, cancellationToken);

    private static HtmlRenderDocument Render(
        IHtmlDocument document,
        HtmlRenderOptions? options,
        HtmlConversionLimits limits,
        bool sourceAlreadyValidated,
        CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (limits == null) throw new ArgumentNullException(nameof(limits));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        HtmlConversionLimits effectiveLimits = limits.Clone();
        effectiveLimits.Validate();
        if (sourceAlreadyValidated) {
            resolved.MaxHtmlNodes = Math.Min(
                resolved.MaxHtmlNodes,
                effectiveLimits.MaxHtmlNodes ?? int.MaxValue);
        }
        resolved.Validate();
        return ExecuteWithDeadline(resolved, cancellationToken, operationCancellationToken => {
            IHtmlDocument renderDocument = HtmlDocumentParser.CloneDocument(document, operationCancellationToken);
            HtmlRenderInputGuard.ValidateFormState(null, renderDocument, resolved, operationCancellationToken);
            HtmlEditableLayoutProjector.CopyMarkers(document, renderDocument);
            if (!sourceAlreadyValidated) {
                HtmlRenderInputGuard.ValidateSource(
                    renderDocument.DocumentElement?.OuterHtml ?? string.Empty,
                    resolved);
            }
            return RenderDocument(
                renderDocument,
                resolved,
                initialDiagnostics: null,
                effectiveLimits,
                operationCancellationToken);
        });
    }

    private static HtmlRenderDocument RenderDocument(
        IHtmlDocument document,
        HtmlRenderOptions resolved,
        IEnumerable<HtmlDiagnostic>? initialDiagnostics,
        HtmlConversionLimits limits,
        CancellationToken cancellationToken) {
        resolved.ResponsiveImageCandidateLimit = limits.MaxResponsiveImageCandidates;
        resolved.ResponsiveImageSizesCharacterLimit = limits.MaxResponsiveImageSizesCharacters;
        var diagnostics = new HtmlDiagnosticReport();
        if (initialDiagnostics != null) diagnostics.AddRange(initialDiagnostics);
        HtmlSerializedShadowRootProjector.Apply(document, resolved, diagnostics, cancellationToken);
        HtmlRenderAdditionalStylesheetApplier.Apply(document, resolved.AdditionalStylesheets.ToList());
        HtmlCssRuleBlockScanner.ValidateDocument(document, limits);
        HtmlRenderInputGuard.ValidateDocument(document, resolved, cancellationToken);
        var resourceOptions = new HtmlResourcePipelineOptions {
            BaseUri = resolved.BaseUri,
            UrlPolicy = (resolved.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
            ResourceUrlPolicy = resolved.GetResourceUrlPolicy().Clone(),
            Limits = limits.Clone(),
            MaxResponsiveImageCandidates = resolved.ResponsiveImageCandidateLimit,
            MaxResponsiveImageSizesCharacters = resolved.ResponsiveImageSizesCharacterLimit,
            MediaContext = resolved.MediaContext,
            MediaWidth = resolved.CssMediaWidth,
            MediaHeight = resolved.CssMediaHeight,
            DevicePixelRatio = resolved.MediaFeatures.ResolutionDpi / HtmlRenderOptions.CssPixelsPerInch,
            DefaultFontSize = resolved.DefaultFontSize,
            MediaFeatures = resolved.MediaFeatures.Clone(),
            IncludeDocumentIcons = false
        };
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(document, resourceOptions);
        cancellationToken.ThrowIfCancellationRequested();
        diagnostics.AddRange(manifest.Diagnostics);
        HtmlCssByteBudget cssBudget = HtmlRenderStylesheetApplier.CreateBudget(document, limits, resolved);
        HtmlResourceSession resources = HtmlRenderResourceLoader.Load(
            manifest,
            resolved,
            diagnostics,
            limits,
            cancellationToken,
            cssBudget);
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderStylesheetApplier.Apply(document, resources, resolved, limits, cssBudget, diagnostics);
        HtmlCssRuleBlockScanner.ValidateDocument(document, limits);
        AddPendingStylesheetDiagnostics(manifest, resources, diagnostics);
        OfficeIMO.Drawing.OfficeFontFaceCollection fonts = HtmlRenderFontFaceLoader.Load(document, resources, resolved, limits, diagnostics);
        fonts.AddRange(resolved.Fonts);
        HtmlCssPageRuleSet pageRules = HtmlCssPageSettingsResolver.Apply(document, resolved, diagnostics);
        resolved.Validate();
        HtmlComputedStyleSet styles = HtmlComputedStyleEngine.ComputeForRendering(document, resolved, limits);
        if (HtmlCssPrintFitResolver.TryApplyWideRoot(document, styles, pageRules, resolved)) {
            resolved.Validate();
            styles = HtmlComputedStyleEngine.ComputeForRendering(document, resolved, limits);
        }
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderDocument rendered = new HtmlRenderLayoutEngine(
            document,
            styles,
            resolved,
            diagnostics,
            resources,
            pageRules,
            fonts,
            limits: limits,
            cancellationToken: cancellationToken).Render();
        return CompleteRender(rendered, resolved);
    }

    /// <summary>
    /// Renders a parsed HTML source while asynchronously resolving policy-approved external resources through the configured resolver.
    /// </summary>
    public static async Task<HtmlRenderDocument> RenderAsync(HtmlConversionDocument document, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) =>
        (await ExecuteAsync(
            document,
            HtmlRenderRequest.FromLegacy(options, HtmlRenderEncoder.DisplayList, HtmlRenderPageSet.All()),
            cancellationToken).ConfigureAwait(false)).Document;

    /// <summary>
    /// Asynchronously resolves resources and prepares an explicit render request as one retained
    /// backend-neutral result.
    /// </summary>
    public static async Task<HtmlRenderResult> ExecuteAsync(
        HtmlConversionDocument document,
        HtmlRenderRequest request,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (request == null) throw new ArgumentNullException(nameof(request));
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderOptions resolved = PrepareOptions(document, request);
        return await ExecuteWithDeadlineAsync(resolved, cancellationToken,
            operationCancellationToken => ExecuteCoreAsync(document, request, resolved, operationCancellationToken))
            .ConfigureAwait(false);
    }

    internal static HtmlRenderOptions PrepareOptions(HtmlConversionDocument document, HtmlRenderRequest request) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (request == null) throw new ArgumentNullException(nameof(request));
        HtmlRenderOptions resolved = request.ResolveOptions();
        resolved.BaseUri ??= document.BaseUri;
        ApplyDocumentPolicies(document, resolved);
        resolved.Validate();
        return resolved;
    }

    internal static HtmlRenderResult ExecuteCore(
        HtmlConversionDocument document,
        HtmlRenderRequest request,
        HtmlRenderOptions resolved,
        CancellationToken cancellationToken) {
        HtmlRenderInputGuard.ValidateSource(document.SourceHtml, resolved);
        cancellationToken.ThrowIfCancellationRequested();
        IHtmlDocument renderDocument = document.CreateDocumentForRendering();
        HtmlRenderInputGuard.ValidateFormState(document.SourceHtml.Length, renderDocument, resolved, cancellationToken);
        HtmlRenderDocument rendered = RenderDocument(
            renderDocument, resolved, initialDiagnostics: null, document.Limits, cancellationToken);
        return HtmlRenderRequestProcessor.Complete(request, rendered, resolved, cancellationToken);
    }

    internal static async Task<HtmlRenderResult> ExecuteCoreAsync(
        HtmlConversionDocument document,
        HtmlRenderRequest request,
        HtmlRenderOptions resolved,
        CancellationToken cancellationToken) {
        HtmlRenderInputGuard.ValidateSource(document.SourceHtml, resolved);
        cancellationToken.ThrowIfCancellationRequested();
        IHtmlDocument renderDocument = document.CreateDocumentForRendering();
        HtmlRenderInputGuard.ValidateFormState(document.SourceHtml.Length, renderDocument, resolved, cancellationToken);
        HtmlRenderDocument rendered = await RenderDocumentAsync(
            renderDocument, resolved, initialDiagnostics: null, document.Limits, cancellationToken).ConfigureAwait(false);
        return HtmlRenderRequestProcessor.Complete(request, rendered, resolved, cancellationToken);
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
        return await ExecuteWithDeadlineAsync(resolved, cancellationToken, async operationCancellationToken => {
            IHtmlDocument renderDocument = HtmlDocumentParser.CloneDocument(document, operationCancellationToken);
            HtmlRenderInputGuard.ValidateFormState(null, renderDocument, resolved, operationCancellationToken);
            HtmlRenderInputGuard.ValidateSource(renderDocument.DocumentElement?.OuterHtml ?? string.Empty, resolved);
            return await RenderDocumentAsync(
                renderDocument,
                resolved,
                initialDiagnostics: null,
                HtmlConversionLimits.CreateUntrustedProfile(),
                operationCancellationToken).ConfigureAwait(false);
        }).ConfigureAwait(false);
    }

    private static async Task<HtmlRenderDocument> RenderDocumentAsync(
        IHtmlDocument document,
        HtmlRenderOptions resolved,
        IEnumerable<HtmlDiagnostic>? initialDiagnostics,
        HtmlConversionLimits limits,
        CancellationToken cancellationToken) {
        resolved.ResponsiveImageCandidateLimit = limits.MaxResponsiveImageCandidates;
        resolved.ResponsiveImageSizesCharacterLimit = limits.MaxResponsiveImageSizesCharacters;
        var diagnostics = new HtmlDiagnosticReport();
        if (initialDiagnostics != null) diagnostics.AddRange(initialDiagnostics);
        HtmlSerializedShadowRootProjector.Apply(document, resolved, diagnostics, cancellationToken);
        HtmlRenderAdditionalStylesheetApplier.Apply(document, resolved.AdditionalStylesheets.ToList());
        HtmlCssRuleBlockScanner.ValidateDocument(document, limits);
        HtmlRenderInputGuard.ValidateDocument(document, resolved, cancellationToken);
        var resourceOptions = new HtmlResourcePipelineOptions {
            BaseUri = resolved.BaseUri,
            UrlPolicy = (resolved.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
            ResourceUrlPolicy = resolved.GetResourceUrlPolicy().Clone(),
            Limits = limits.Clone(),
            MaxResponsiveImageCandidates = resolved.ResponsiveImageCandidateLimit,
            MaxResponsiveImageSizesCharacters = resolved.ResponsiveImageSizesCharacterLimit,
            MediaContext = resolved.MediaContext,
            MediaWidth = resolved.CssMediaWidth,
            MediaHeight = resolved.CssMediaHeight,
            DevicePixelRatio = resolved.MediaFeatures.ResolutionDpi / HtmlRenderOptions.CssPixelsPerInch,
            DefaultFontSize = resolved.DefaultFontSize,
            MediaFeatures = resolved.MediaFeatures.Clone(),
            IncludeDocumentIcons = false
        };
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(document, resourceOptions);
        diagnostics.AddRange(manifest.Diagnostics);
        HtmlCssByteBudget cssBudget = HtmlRenderStylesheetApplier.CreateBudget(document, limits, resolved);
        HtmlResourceSession resources = await HtmlRenderResourceLoader.LoadAsync(manifest, resolved, diagnostics, limits, cancellationToken, cssBudget).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderStylesheetApplier.Apply(document, resources, resolved, limits, cssBudget, diagnostics);
        HtmlCssRuleBlockScanner.ValidateDocument(document, limits);
        AddPendingStylesheetDiagnostics(manifest, resources, diagnostics);
        OfficeIMO.Drawing.OfficeFontFaceCollection fonts = HtmlRenderFontFaceLoader.Load(document, resources, resolved, limits, diagnostics);
        fonts.AddRange(resolved.Fonts);
        HtmlCssPageRuleSet pageRules = HtmlCssPageSettingsResolver.Apply(document, resolved, diagnostics);
        cancellationToken.ThrowIfCancellationRequested();
        resolved.Validate();
        HtmlComputedStyleSet styles = HtmlComputedStyleEngine.ComputeForRendering(document, resolved, limits);
        if (HtmlCssPrintFitResolver.TryApplyWideRoot(document, styles, pageRules, resolved)) {
            resolved.Validate();
            styles = HtmlComputedStyleEngine.ComputeForRendering(document, resolved, limits);
        }
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderDocument rendered = new HtmlRenderLayoutEngine(document, styles, resolved, diagnostics, resources, pageRules, fonts,
            limits: limits, cancellationToken: cancellationToken).Render();
        return CompleteRender(rendered, resolved);
    }

    internal static HtmlRenderDocument RenderHtml(this string html, HtmlRenderOptions? options = null) => Render(html, options);

    internal static Task<HtmlRenderDocument> RenderHtmlAsync(this string html, HtmlRenderOptions? options = null, CancellationToken cancellationToken = default) =>
        RenderAsync(html, options, cancellationToken);

    private static HtmlRenderDocument CompleteRender(HtmlRenderDocument rendered, HtmlRenderOptions options) =>
        options.FidelityPolicy == HtmlRenderFidelityPolicy.RequireNoLoss
            ? rendered.RequireNoLoss()
            : rendered;

    internal static T ExecuteWithDeadline<T>(
        HtmlRenderOptions options,
        CancellationToken callerCancellationToken,
        Func<CancellationToken, T> operation) {
        using OfficeIMO.Drawing.OfficeImageExportExecutionScope execution =
            OfficeIMO.Drawing.OfficeImageExportExecutionScope.Start(options.RenderTimeout, callerCancellationToken);
        try {
            T result = operation(execution.Token);
            execution.ThrowIfCancellationRequested();
            return result;
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

    internal static async Task<T> ExecuteWithDeadlineAsync<T>(
        HtmlRenderOptions options,
        CancellationToken callerCancellationToken,
        Func<CancellationToken, Task<T>> operation) {
        using OfficeIMO.Drawing.OfficeImageExportExecutionScope execution =
            OfficeIMO.Drawing.OfficeImageExportExecutionScope.Start(options.RenderTimeout, callerCancellationToken);
        try {
            T result = await operation(execution.Token).ConfigureAwait(false);
            execution.ThrowIfCancellationRequested();
            return result;
        } catch (OperationCanceledException exception) when (execution.IsTimeoutCancellation(exception)) {
            throw execution.CreateTimeoutException(exception);
        }
    }

    internal static void ApplyDocumentPolicies(HtmlConversionDocument document, HtmlRenderOptions options) {
        HtmlUrlPolicy requestedHyperlinkPolicy = options.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile();
        HtmlUrlPolicy requestedResourcePolicy = options.ResourceUrlPolicy ?? requestedHyperlinkPolicy;
        options.UrlPolicy = HtmlUrlPolicy.Intersect(document.HyperlinkUrlPolicy, requestedHyperlinkPolicy);
        options.ResourceUrlPolicy = HtmlUrlPolicy.Intersect(document.ResourceUrlPolicy, requestedResourcePolicy);
    }

    internal static void AddPendingStylesheetDiagnostics(HtmlResourceManifest manifest, HtmlResourceSession resources, HtmlDiagnosticReport diagnostics) {
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
                reference.ResolvedSource,
                OfficeConversionLossKind.Omission);
        }
    }
}
