using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Globalization;

namespace OfficeIMO.Html;

public sealed partial class HtmlConversionDocument {
    /// <summary>
    /// Creates an independent editable source without elements hidden by computed
    /// <c>display: none</c>, zero opacity, or a closed <c>details</c> element
    /// in the requested screen or print environment.
    /// The returned source keeps its original stylesheet links and bounded import contract.
    /// A supported computed image maximum width is retained as an inline hint so editable
    /// targets do not expand a captured image to its unconstrained intrinsic size.
    /// </summary>
    public Task<HtmlVisibleContentResult> CreateVisibleContentDocumentResultAsync(
        HtmlRenderOptions? options = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        HtmlRenderOptions resolved = options?.Clone() ?? new HtmlRenderOptions();
        resolved.BaseUri ??= BaseUri;
        HtmlRenderEngine.ApplyDocumentPolicies(this, resolved);
        resolved.Validate();
        return HtmlRenderEngine.ExecuteWithDeadlineAsync(resolved, cancellationToken,
            token => CreateVisibleContentDocumentCoreAsync(resolved, token));
    }

    private async Task<HtmlVisibleContentResult> CreateVisibleContentDocumentCoreAsync(
        HtmlRenderOptions resolved, CancellationToken cancellationToken) {
        HtmlConversionLimits limits = _options.Limits.Clone();
        IHtmlDocument source = CreateDocumentForRendering();
        var diagnostics = new HtmlDiagnosticReport();
        HtmlSerializedShadowRootProjector.Apply(source, resolved, diagnostics, cancellationToken);
        var originalStyles = new HashSet<IElement>(source.QuerySelectorAll("style"));

        HtmlRenderAdditionalStylesheetApplier.Apply(source, resolved.AdditionalStylesheets.ToList());
        HtmlCssRuleBlockScanner.ValidateDocument(source, limits);
        var resourceOptions = new HtmlResourcePipelineOptions {
            BaseUri = resolved.BaseUri ?? BaseUri,
            UrlPolicy = (resolved.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
            ResourceUrlPolicy = resolved.GetResourceUrlPolicy().Clone(),
            Limits = limits.Clone(),
            MaxResponsiveImageCandidates = resolved.ResponsiveImageCandidateLimit,
            MaxResponsiveImageSizesCharacters = resolved.ResponsiveImageSizesCharacterLimit,
            MediaContext = resolved.MediaContext,
            MediaWidth = resolved.CssMediaWidth,
            MediaHeight = resolved.Mode == HtmlRenderMode.Paged ? resolved.PageHeight : resolved.ViewportHeight ?? 1056D,
            DevicePixelRatio = resolved.MediaFeatures.ResolutionDpi / HtmlRenderOptions.CssPixelsPerInch,
            DefaultFontSize = resolved.DefaultFontSize,
            MediaFeatures = resolved.MediaFeatures.Clone()
        };
        HtmlResourceManifest discovered = HtmlResourcePipeline.BuildManifest(source, resourceOptions);
        var stylesheets = new HtmlResourceManifest();
        foreach (HtmlResourceReference reference in discovered.Resources) {
            if (reference.Kind == HtmlResourceKind.Stylesheet) stylesheets.Add(reference);
        }
        foreach (HtmlResourceReference reference in stylesheets.Resources) {
            if (!reference.IsAllowed) {
                diagnostics.Add("OfficeIMO.Html",
                    reference.DiagnosticCode.Length == 0 ? "HtmlResourceRejectedByPolicy" : reference.DiagnosticCode,
                    "A linked stylesheet was rejected by the configured URL policy and cannot select visible content.",
                    HtmlDiagnosticSeverity.Warning, reference.Source, reference.ElementName + "[" + reference.AttributeName + "]",
                    OfficeConversionLossKind.Omission);
            }
        }
        HtmlCssByteBudget cssBudget = HtmlRenderStylesheetApplier.CreateBudget(source, limits, resolved);
        HtmlResourceSession resources = await HtmlRenderResourceLoader.LoadAsync(
            stylesheets, resolved, diagnostics, limits, cancellationToken, cssBudget).ConfigureAwait(false);
        HtmlRenderStylesheetApplier.Apply(source, resources, resolved, limits, cssBudget, diagnostics);
        HtmlCssRuleBlockScanner.ValidateDocument(source, limits);
        HtmlRenderEngine.AddPendingStylesheetDiagnostics(stylesheets, resources, diagnostics);
        cancellationToken.ThrowIfCancellationRequested();

        IElement[] appliedStyles = source.QuerySelectorAll("style")
            .Where(style => !originalStyles.Contains(style)).ToArray();
        HtmlComputedStyleSet styles = HtmlComputedStyleEngine.ComputeForRendering(source, resolved, limits);
        IElement? body = source.Body;
        int omitted = 0;
        if (body != null) {
            bool hiddenRoot = styles.Elements.Any(pair =>
                (ReferenceEquals(pair.Key, source.DocumentElement) || ReferenceEquals(pair.Key, body))
                && IsInvisibleForEditableContent(pair.Value));
            if (hiddenRoot) {
                omitted = body.QuerySelectorAll("*").Length + 1;
                foreach (INode child in body.ChildNodes.ToArray()) body.RemoveChild(child);
            } else {
                foreach (KeyValuePair<IElement, HtmlComputedStyle> pair in styles.Elements) {
                    IElement element = pair.Key;
                    if (!body.Contains(element) || !IsInvisibleForEditableContent(pair.Value)) continue;
                    element.Remove();
                    omitted++;
                }
            }
        }
        if (body != null) {
            foreach (IElement details in body.QuerySelectorAll("details:not([open])")) {
                if (!body.Contains(details)) continue;
                IElement? summary = details.Children.FirstOrDefault(child =>
                    child.LocalName.Equals("summary", StringComparison.OrdinalIgnoreCase));
                foreach (INode child in details.ChildNodes.ToArray()) {
                    if (ReferenceEquals(child, summary)) continue;
                    if (child is IElement) omitted++;
                    details.RemoveChild(child);
                }
            }
        }
        if (appliedStyles.Length > 0 && body != null) {
            foreach (IHtmlImageElement image in body.QuerySelectorAll("img").OfType<IHtmlImageElement>()) {
                if (!styles.Elements.TryGetValue(image, out HtmlComputedStyle? style)) continue;
                string maximum = style.GetValue("max-width").Trim();
                if (!IsPortableImageMaximumWidth(maximum)) continue;
                string existing = image.GetAttribute("style")?.Trim() ?? string.Empty;
                image.SetAttribute("style", existing.Length == 0
                    ? "max-width:" + maximum
                    : existing.TrimEnd(';') + ";max-width:" + maximum);
            }
        }
        foreach (IElement style in appliedStyles) style.Remove();
        cancellationToken.ThrowIfCancellationRequested();

        HtmlConversionDocument editable = FromDocument(NativeDomBridge.Import(source), _options);
        return new HtmlVisibleContentResult(editable, diagnostics.Diagnostics.ToArray(),
            omitted, appliedStyles.Count(style => style.HasAttribute("data-officeimo-source")));
    }

    private static bool IsInvisibleForEditableContent(HtmlComputedStyle style) {
        if (style.GetValue("display").Trim().Equals("none", StringComparison.OrdinalIgnoreCase)) return true;
        string opacity = style.GetValue("opacity").Trim();
        return double.TryParse(opacity, NumberStyles.Float,
            CultureInfo.InvariantCulture, out double alpha) && alpha <= 0D;
    }

    private static bool IsPortableImageMaximumWidth(string value) {
        bool pixels = value.EndsWith("px", StringComparison.OrdinalIgnoreCase);
        bool percentage = value.EndsWith("%", StringComparison.Ordinal);
        if (!pixels && !percentage) return false;
        string number = value.Substring(0, value.Length - (pixels ? 2 : 1));
        return double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
            && !double.IsNaN(parsed) && !double.IsInfinity(parsed) && parsed > 0D;
    }
}
