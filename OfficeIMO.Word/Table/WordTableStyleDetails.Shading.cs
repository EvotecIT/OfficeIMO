using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Word;

public partial class WordTableStyleDetails {
    /// <summary>Native direct table shading fill as RGB hex or <c>auto</c>.</summary>
    /// <remarks>
    /// A leading # is accepted. Empty or null removes direct table shading and restores style inheritance.
    /// An explicit auto value clears inherited fill.
    /// </remarks>
    internal string ShadingFillValue {
        get => Shading.Attribute(W, "fill")?.ToUpperInvariant() ?? string.Empty;
        set {
            string normalized = NormalizeTableShadingColor(value, nameof(ShadingFillColor));
            if (normalized.Length == 0) { Shading.Remove(); return; }
            OfficeXmlNode shading = GetOrCreateTableShading();
            shading.SetAttribute(W, "fill", normalized);
            shading.SetAttribute(W, "themeFill", null);
            shading.SetAttribute(W, "themeFillTint", null);
            shading.SetAttribute(W, "themeFillShade", null);
        }
    }

    /// <summary>Gets or sets the direct table fill color. Null removes direct table shading.</summary>
    /// <remarks>With positive cell spacing, table fill colors the gaps.</remarks>
    public Color? ShadingFillColor {
        get => ParseTableShadingColor(ShadingFillValue);
        set => ShadingFillValue = value?.ToRgbHex() ?? string.Empty;
    }

    /// <summary>Native direct shading pattern color as RGB hex or <c>auto</c>.</summary>
    /// <remarks>Empty or null clears only the pattern color, retaining the fill and pattern.</remarks>
    internal string ShadingColorValue {
        get => Shading.Attribute(W, "color")?.ToUpperInvariant() ?? string.Empty;
        set {
            string normalized = NormalizeTableShadingColor(value, nameof(ShadingColor));
            OfficeXmlNode shading = normalized.Length == 0 ? Shading : GetOrCreateTableShading();
            if (shading.IsNone) return;
            shading.SetAttribute(W, "color", normalized.Length == 0 ? null : normalized);
            shading.SetAttribute(W, "themeColor", null);
            shading.SetAttribute(W, "themeTint", null);
            shading.SetAttribute(W, "themeShade", null);
        }
    }

    /// <summary>Gets or sets the direct shading pattern color. Null clears only that color.</summary>
    public Color? ShadingColor {
        get => ParseTableShadingColor(ShadingColorValue);
        set => ShadingColorValue = value?.ToRgbHex() ?? string.Empty;
    }

    /// <summary>Gets or sets the direct table shading pattern. Null removes direct table shading.</summary>
    public WordShadingPattern? ShadingPattern {
        get => Shading.Attribute(W, "val") is string token && WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out WordShadingPattern pattern) ? pattern : null;
        set {
            if (!value.HasValue) { Shading.Remove(); return; }
            string pattern = value.Value.ToOpenXmlToken();
            GetOrCreateTableShading().SetAttribute(W, "val", pattern);
        }
    }

    // Only the table's own w:shd counts, not one inside extension markup.
    private OfficeXmlNode Shading => _properties.Child(W, "shd");

    private OfficeXmlNode GetOrCreateTableShading() {
        OfficeXmlNode shading = Shading;
        if (shading.IsNone) {
            shading = _properties.AddOrdered(W, "shd", WTblPrView.Order);
            shading.SetAttribute(W, "val", "clear");
        }
        return shading;
    }

    private static string NormalizeTableShadingColor(string? value, string parameter) {
        string color = value?.Trim() ?? string.Empty;
        if (color.Length == 0 || string.Equals(color, "auto", StringComparison.OrdinalIgnoreCase)) return color.ToLowerInvariant();
        if (color.StartsWith("#", StringComparison.Ordinal)) color = color.Substring(1);
        if (color.Length != 6 || !Color.TryParseHex(color, out _))
            throw new ArgumentException("Table shading color must be six-digit RGB hex or auto.", parameter);
        return color.ToUpperInvariant();
    }

    private static Color? ParseTableShadingColor(string value) =>
        value.Length == 6 && Color.TryParseHex(value, out Color color) ? color : null;
}
