using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlRenderSystemFontLoader {
    private const int MaximumFamilyAttempts = 32;
    private static readonly HtmlPseudoElementKind[] Pseudos = {
        HtmlPseudoElementKind.Before, HtmlPseudoElementKind.After, HtmlPseudoElementKind.Marker,
        HtmlPseudoElementKind.FootnoteCall, HtmlPseudoElementKind.FootnoteMarker,
        HtmlPseudoElementKind.FirstLetter, HtmlPseudoElementKind.FirstLine, HtmlPseudoElementKind.Placeholder
    };

    internal static void Load(IHtmlDocument document, HtmlComputedStyleSet styles,
        OfficeFontFaceCollection fonts, HtmlRenderOptions options, HtmlResourceSession resources,
        HtmlDiagnosticReport diagnostics, CancellationToken cancellationToken) {
        if (!options.AllowSystemFontFallback) return;
        var inherited = new Dictionary<IElement, (string Families, OfficeFontFaceDescriptor Descriptor)>();
        var attempted = new Dictionary<(string Family, int Weight, OfficeFontSlant Slant), bool>();
        var supplied = new HashSet<string>(fonts.Faces.Select(face => face.FamilyName), StringComparer.OrdinalIgnoreCase);
        bool limitReported = false;
        foreach (IElement element in document.QuerySelectorAll("*")) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!styles.Elements.TryGetValue(element, out HtmlComputedStyle? computed)) continue;
            (string Families, OfficeFontFaceDescriptor Descriptor) parent = element.ParentElement != null
                && inherited.TryGetValue(element.ParentElement, out var found)
                ? found : (options.DefaultFontFamily, OfficeFontFaceDescriptor.Regular);
            string tag = element.TagName.ToLowerInvariant();
            string fallback = tag is "code" or "pre" or "kbd" or "samp" ? "Consolas" : parent.Families;
            string families = HtmlRenderStyleResolver.ResolveFontFamily(tag, computed, fallback);
            OfficeFontFaceDescriptor descriptor = HtmlRenderStyleResolver.ResolveFontFaceDescriptor(tag, computed, parent.Descriptor);
            inherited[element] = (families, descriptor);
            LoadList(families, descriptor);
            foreach (HtmlPseudoElementKind kind in Pseudos) {
                if (!styles.TryGetPseudoStyle(element, kind, out HtmlComputedStyle pseudo)) continue;
                LoadList(HtmlRenderCssValues.FontFamilyList(pseudo.GetValue("font-family"), families),
                    HtmlRenderStyleResolver.ResolveFontFaceDescriptor(string.Empty, pseudo, descriptor));
            }
        }

        void LoadList(string familyNames, OfficeFontFaceDescriptor descriptor) {
            IReadOnlyList<string> families = OfficeFontFamilyParser.Parse(familyNames);
            if (!families.Any(family => OfficeSystemFontFamilyAliases.IsSystemUi(family) || OfficeSystemFontFamilyAliases.IsMath(family))) return;
            foreach (string family in families) {
                cancellationToken.ThrowIfCancellationRequested();
                // A PDF target can permit library-selected mathematical fonts without
                // allowing a document's named host family or system-UI fallback list.
                if (options.MathOnlySystemFontFallback && !OfficeSystemFontFamilyAliases.IsMath(family)) continue;
                if (supplied.Contains(family)) continue;
                if (OfficeSystemFontFamilyAliases.IsMath(family)
                    && supplied.Any(candidate => OfficeSystemFontFamilyAliases.MathFamilyRank(candidate) != int.MaxValue)) continue;
                var key = (family.ToLowerInvariant(), descriptor.Weight, descriptor.Slant);
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
                        "An installed font face was loaded for a generic fallback list.", HtmlDiagnosticSeverity.Info,
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
