using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderStyleResolver {
    private readonly HashSet<string> _reportedGeneratedColors = new HashSet<string>(StringComparer.Ordinal);

    private void ReportUnsupportedGeneratedColor(string source, string property, string value) {
        if (_reportedGeneratedColors.Add(source + " " + property))
            ReportUnsupportedColor(source, property, value);
    }

    /// <summary>Resolves solid paint identically for element and non-DOM generated contexts.</summary>
    private static bool TryResolveBackgroundColor(HtmlComputedStyle computed, OfficeColor foreground, out OfficeColor? color) {
        string value = computed.GetValue("background-color");
        if (value.Length == 0) value = computed.GetValue("background");
        if (string.Equals(value.Trim(), "currentcolor", StringComparison.OrdinalIgnoreCase)) {
            color = foreground;
            return true;
        }
        if (HtmlRenderCssValues.TryColor(value, out OfficeColor parsed)) {
            color = parsed;
            return true;
        }
        color = null;
        return value.Length == 0;
    }

    /// <summary>Resolves generated paint through the same bounded layer parser as ordinary boxes.</summary>
    private void ResolvePageMarginBackground(HtmlComputedStyle computed, HtmlRenderBoxStyle style, string source) {
        if (computed.Properties.Any(pair => HtmlCssPageSettingsResolver.IsMarginBackgroundDeclaration(pair.Key)
            && ContainsCharacterUnit(pair.Value))) style.CharacterAdvance = _measureCharacterAdvance(style);
        if (TryResolveBackgroundColor(computed, style.Color, out OfficeColor? color))
            style.BackgroundColor = color;
        else
            ReportUnsupportedGeneratedColor(source, "background-color", computed.GetValue("background-color"));
        string shorthand = computed.GetValue("background");
        string sourceValue = computed.GetValue("background-image");
        style.BackgroundColorClip = ResolveSolidBackgroundClip(computed, shorthand,
            HtmlRenderCssValues.SplitTopLevelCommas(sourceValue.Length > 0 ? sourceValue : shorthand).Count, out bool supportedClip);
        if (!supportedClip && _reportedPageMarginDimensions.Add(source + " background-clip")) {
            _diagnostics.Add("OfficeIMO.Html", HtmlRenderDiagnosticCodes.BackgroundImageValueUnsupported,
                "A page-margin background clip used the border-box fallback.", HtmlDiagnosticSeverity.Warning,
                source, "background-clip=" + computed.GetValue("background-clip"), OfficeConversionLossKind.Approximation);
        }
        ApplyBackgroundLayers(computed, style, shorthand);
    }

    /// <summary>Uses the final image layer's recycled clip for solid color, including generated boxes.</summary>
    private static string ResolveSolidBackgroundClip(HtmlComputedStyle computed, string shorthand, int imageLayers, out bool supported) {
        IReadOnlyList<string> clips = HtmlRenderCssValues.SplitTopLevelCommas(computed.GetValue("background-clip"));
        (_, string fallback) = ExtractBackgroundBoxes(shorthand);
        string value = GetLayerValue(clips, Math.Max(0, imageLayers - 1), fallback, "border-box").Trim().ToLowerInvariant();
        supported = value is "border-box" or "padding-box" or "content-box";
        return HtmlRenderBackgroundLayer.NormalizeBox(value, "border-box");
    }
}
