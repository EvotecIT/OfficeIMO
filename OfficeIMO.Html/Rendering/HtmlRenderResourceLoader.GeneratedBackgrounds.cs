using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

public sealed partial class HtmlResourceSession {
    private readonly HashSet<string> _requestIdentities = new(HtmlResourceSeenKeyComparer.Instance);
    private readonly Dictionary<string, HtmlResourceReference> _imagePlans = new(HtmlResourceIdentityComparer.Instance);
    private readonly Dictionary<string, HtmlResourceReference> _stylesheetImagePlans = new(StringComparer.Ordinal);

    internal bool TryAddRequestIdentity(string identity) => _requestIdentities.Add(identity);

    internal void RegisterImagePlan(HtmlResourceReference reference) {
        if (reference.Kind != HtmlResourceKind.Image) return;
        string materialized = GetMaterializedImageSource(reference);
        if (materialized.Length > 0) _imagePlans[materialized] = reference;
        if (reference.ResolutionBaseUri?.IsAbsoluteUri == true)
            _stylesheetImagePlans[reference.ResolutionBaseUri.AbsoluteUri + "\n" + reference.Source] = reference;
    }

    internal bool TryGetPlannedReference(string source, out HtmlResourceReference reference) =>
        _imagePlans.TryGetValue(source, out reference!);

    internal bool TryGetStylesheetImagePlan(string source, Uri baseUri, out HtmlResourceReference reference) =>
        TryGetPlannedReference(source, out reference)
        || _stylesheetImagePlans.TryGetValue(baseUri.AbsoluteUri + "\n" + source, out reference!);

    // Keep a rejected source identifiable until its effective generated use is known. It remains a
    // rejected plan and never enters dispatch; replacing every blocked URL with url("") loses that use.
    internal static string GetMaterializedImageSource(HtmlResourceReference reference) {
        if (reference.IsAllowed) return reference.ResolvedSource;
        if (Uri.TryCreate(reference.Source, UriKind.Absolute, out Uri? absolute)) return absolute.AbsoluteUri;
        if (reference.ResolutionBaseUri?.IsAbsoluteUri == true && Uri.TryCreate(reference.ResolutionBaseUri, reference.Source, out Uri? relative))
            return relative.AbsoluteUri;
        return reference.Source;
    }
}

internal static partial class HtmlRenderResourceLoader {
    /// <summary>Uses the operation session after external CSS has joined the generated cascade.</summary>
    internal static void LoadGeneratedBackgrounds(IHtmlDocument document, HtmlCssPageRuleSet pageRules,
        HtmlComputedStyleSet styles, HtmlResourcePipelineOptions resourceOptions, HtmlRenderOptions options,
        HtmlResourceSession resources, HtmlConversionLimits limits, CancellationToken cancellationToken) {
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildPageMarginBackgroundManifest(
            document, pageRules, styles, resourceOptions, cancellationToken, resources);
        int firstDiagnostic = resources.Diagnostics.Count;
        resources.Diagnostics.AddRange(manifest.Diagnostics);
        Load(manifest, options, resources.Diagnostics, limits, cancellationToken, existingSession: resources);
        resources.DeferMissingCssImageLoss(firstDiagnostic);
    }

    /// <summary>Dispatches final generated resources without resetting request, byte, or count budgets.</summary>
    internal static async Task LoadGeneratedBackgroundsAsync(IHtmlDocument document, HtmlCssPageRuleSet pageRules,
        HtmlComputedStyleSet styles, HtmlResourcePipelineOptions resourceOptions, HtmlRenderOptions options,
        HtmlResourceSession resources, HtmlConversionLimits limits, CancellationToken cancellationToken) {
        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildPageMarginBackgroundManifest(
            document, pageRules, styles, resourceOptions, cancellationToken, resources);
        int firstDiagnostic = resources.Diagnostics.Count;
        resources.Diagnostics.AddRange(manifest.Diagnostics);
        await LoadAsync(manifest, options, resources.Diagnostics, limits, cancellationToken, existingSession: resources).ConfigureAwait(false);
        resources.DeferMissingCssImageLoss(firstDiagnostic);
    }
}
