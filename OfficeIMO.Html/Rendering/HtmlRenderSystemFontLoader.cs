using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static partial class HtmlRenderSystemFontLoader {
    private const int MaximumFamilyAttempts = 32;
    private static readonly HtmlPseudoElementKind[] Pseudos = {
        HtmlPseudoElementKind.Before, HtmlPseudoElementKind.After, HtmlPseudoElementKind.Marker,
        HtmlPseudoElementKind.FootnoteCall, HtmlPseudoElementKind.FootnoteMarker,
        HtmlPseudoElementKind.FirstLetter, HtmlPseudoElementKind.FirstLine, HtmlPseudoElementKind.Placeholder
    };

    internal static void Load(IHtmlDocument document, HtmlComputedStyleSet styles,
        OfficeFontFaceCollection fonts, HtmlRenderOptions options, HtmlResourceSession resources,
        HtmlDiagnosticReport diagnostics, CancellationToken cancellationToken, HtmlCssPageRuleSet? pageRules = null) {
        IReadOnlyList<string> configuredDefaults = Pdf.HtmlPdfRenderedConverter.RetainConfiguredDefaultFonts(
            fonts, options, resources, diagnostics, cancellationToken);
        if (!options.AllowSystemFontFallback) return;
        var inherited = new Dictionary<IElement, (string Families, OfficeFontFaceDescriptor Descriptor)>();
        var attempted = new Dictionary<(string Family, OfficeFontFaceDescriptor Descriptor), bool>();
        var supplied = new HashSet<string>(fonts.Faces.Select(face => face.FamilyName), StringComparer.OrdinalIgnoreCase);
        supplied.UnionWith(configuredDefaults);
        var defaultFamilies = new HashSet<string>(OfficeFontFamilyParser.Parse(options.DefaultFontFamily), StringComparer.OrdinalIgnoreCase);
        var visibility = new Dictionary<IElement, DefaultFontVisibility>();
        var visibleText = new HashSet<IElement>();
        var defaults = new List<(IElement Element, string Families, OfficeFontFaceDescriptor Descriptor, bool NeedsText, bool DirectText)>();
        var disclosures = new HtmlDisclosureState();
        var counterStyles = new Lazy<HtmlCounterStyleRegistry>(() => HtmlCounterStyleRegistry.Parse(document, options));
        IElement[] elements = document.QuerySelectorAll("*").ToArray();
        bool limitReported = false;
        foreach (IElement element in elements) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!styles.Elements.TryGetValue(element, out HtmlComputedStyle? computed)) continue;
            (string Families, OfficeFontFaceDescriptor Descriptor) parent = element.ParentElement != null
                && inherited.TryGetValue(element.ParentElement, out var found)
                ? found : (options.DefaultFontFamily, OfficeFontFaceDescriptor.Regular);
            string tag = element.TagName.ToLowerInvariant();
            string fallback = HtmlRenderStyleResolver.UsesDocumentMonospaceFamily(tag, options) ? HtmlRenderStyleResolver.ResolveInitialFontFamily(tag, options) : parent.Families;
            string families = HtmlRenderStyleResolver.ResolveElementFontFamily(tag, computed, fallback, options);
            OfficeFontFaceDescriptor descriptor = HtmlRenderStyleResolver.ResolveFontFaceDescriptor(tag, computed, parent.Descriptor);
            inherited[element] = (families, descriptor);
            DefaultFontVisibility parentVisibility = element.ParentElement != null
                && visibility.TryGetValue(element.ParentElement, out DefaultFontVisibility inheritedVisibility)
                ? inheritedVisibility : DefaultFontVisibility.Visible;
            DefaultFontVisibility currentVisibility = ResolveDefaultFontVisibility(element, computed, parentVisibility, disclosures);
            visibility[element] = currentVisibility;
            LoadList(families, descriptor);
            if (currentVisibility.CanPaint && HasDirectDefaultFontText(element, computed, disclosures)) {
                visibleText.Add(element);
                defaults.Add((element, families, descriptor, false, true));
            }
            foreach (HtmlPseudoElementKind kind in Pseudos) {
                if (!styles.TryGetPseudoStyle(element, kind, out HtmlComputedStyle pseudo)) continue;
                string pseudoFamilies = HtmlRenderStyleResolver.ResolveFontFamily(string.Empty, pseudo, families, options.DefaultFontFamily);
                OfficeFontFaceDescriptor pseudoDescriptor = HtmlRenderStyleResolver.ResolveFontFaceDescriptor(string.Empty, pseudo, descriptor);
                LoadList(pseudoFamilies, pseudoDescriptor);
                if (!currentVisibility.Blocked && CanPaintDefaultPseudo(pseudo, currentVisibility)
                    && TryGetDefaultPseudoRequest(element, kind, pseudo, out bool needsText)) {
                    defaults.Add((element, pseudoFamilies, pseudoDescriptor, needsText, false));
                }
            }
        }

        // A first-line/marker style can own descendant text without a direct text node.
        // Aggregate that evidence once instead of repeatedly allocating TextContent.
        for (int index = elements.Length - 1; index >= 0; index--) {
            cancellationToken.ThrowIfCancellationRequested();
            IElement element = elements[index];
            if (visibleText.Contains(element) && element.ParentElement != null)
                visibleText.Add(element.ParentElement);
        }
        foreach (var request in defaults) {
            cancellationToken.ThrowIfCancellationRequested();
            if ((!request.NeedsText || visibleText.Contains(request.Element))
                && (!request.DirectText || !ScopedFacesCoverDirectText(request.Element, request.Families, request.Descriptor,
                    fonts, supplied, disclosures, styles.Elements[request.Element], counterStyles, cancellationToken)))
                LoadDefaults(request.Families, request.Descriptor);
        }

        if (pageRules != null && document.DocumentElement is IElement root && styles.Elements.TryGetValue(root, out HtmlComputedStyle? rootStyle)) {
            foreach (var request in pageRules.EnumerateMarginFontRequests(rootStyle, styles, options, cancellationToken))
                LoadList(request.Families, request.Descriptor);
            foreach (var request in pageRules.EnumerateMarginFontRequests(rootStyle, styles, options, cancellationToken, requireGeneratedContent: true))
                LoadDefaults(request.Families, request.Descriptor);
        }

        void LoadDefaults(string familyNames, OfficeFontFaceDescriptor descriptor) {
            // The regular face also supplies the containing inline strut when all
            // visible descendants happen to request bold or italic text.
            LoadList(familyNames, OfficeFontFaceDescriptor.Regular, rendererDefault: true);
            if (descriptor != OfficeFontFaceDescriptor.Regular)
                LoadList(familyNames, descriptor, rendererDefault: true);
        }

        void LoadList(string familyNames, OfficeFontFaceDescriptor descriptor, bool rendererDefault = false) {
            IReadOnlyList<string> families = OfficeFontFamilyParser.Parse(familyNames);
            if (!rendererDefault && !families.Any(family => OfficeSystemFontFamilyAliases.IsSystemUi(family) || OfficeSystemFontFamilyAliases.IsMath(family))) return;
            foreach (string family in families) {
                cancellationToken.ThrowIfCancellationRequested();
                // A default request cannot turn a preceding authored family into
                // permission to discover that family on the host.
                if (rendererDefault && !defaultFamilies.Contains(family)) continue;
                // A PDF target can permit library-selected mathematical fonts without
                // allowing a document's named host family or system-UI fallback list.
                // Its renderer-owned default remains permitted under the same opt-out.
                if (!rendererDefault && options.MathOnlySystemFontFallback && !OfficeSystemFontFamilyAliases.IsMath(family)) continue;
                if (supplied.Contains(family)) continue;
                if (OfficeSystemFontFamilyAliases.IsMath(family)
                    && supplied.Any(candidate => OfficeSystemFontFamilyAliases.MathFamilyRank(candidate) != int.MaxValue)) continue;
                var key = (family.ToLowerInvariant(), descriptor);
                if (attempted.TryGetValue(key, out bool previouslyLoaded)) {
                    if (previouslyLoaded) return;
                    continue;
                }
                if (attempted.Count >= MaximumFamilyAttempts) {
                    if (!limitReported) diagnostics.Add("OfficeIMO.Html.Renderer", "InstalledFontResolutionLimitExceeded",
                        "Installed-font resolution exceeded the operation's family-attempt limit.", HtmlDiagnosticSeverity.Warning,
                        family, "limit=" + MaximumFamilyAttempts);
                    limitReported = true;
                    return;
                }
                attempted.Add(key, false);
                long remaining = resources.RemainingFontBytes;
                string? error = "Installed fonts exceed the operation-wide resource byte limit.";
                int bytes = 0;
                if (fonts.TryAddInstalledFamily(family, descriptor,
                        (int)Math.Max(0L, Math.Min(remaining, int.MaxValue)), cancellationToken, out bytes, out error,
                        resources.MaxResourceBytes)) {
                    resources.AcceptDecodedFontBytes(bytes);
                    attempted[key] = true;
                    diagnostics.Add("OfficeIMO.Html.Renderer", "InstalledFontResolved",
                        rendererDefault ? "An installed face was retained for the configured default font family."
                            : "An installed font face was loaded for a generic fallback list.", HtmlDiagnosticSeverity.Info,
                        family, "weight=" + descriptor.Weight + ";decodedBytes=" + bytes);
                    return;
                } else if (error != null) {
                    string code = error.Contains("per-resource") ? HtmlRenderDiagnosticCodes.ResourceByteLimitExceeded
                        : error.Contains("limit") ? HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded
                        : HtmlRenderDiagnosticCodes.FontFormatUnsupported;
                    diagnostics.Add("OfficeIMO.Html.Renderer", code,
                        "An installed fallback font could not be loaded by the bounded font engine.",
                        HtmlDiagnosticSeverity.Warning, family, error);
                } else if (OfficeSystemFontFamilyAliases.IsSystemUi(family) || OfficeSystemFontFamilyAliases.IsMath(family)) {
                    diagnostics.Add("OfficeIMO.Html.Renderer", "InstalledFontUnavailable",
                        "No supported installed face was available for the requested generic family.",
                        HtmlDiagnosticSeverity.Warning, family);
                }
            }
        }
    }
}
