using DocumentFormat.OpenXml.Wordprocessing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Word;

public partial class WordTableStyleDetails {
    /// <summary>Gets or sets the direct table shading fill as RGB hex or <c>auto</c>.</summary>
    /// <remarks>
    /// A leading # is accepted. Empty or null removes direct table shading and restores style inheritance.
    /// An explicit auto value clears inherited fill. With positive cell spacing, table fill colors the gaps.
    /// </remarks>
    public string ShadingFillColorHex {
        get => _tableProperties.GetFirstChild<Shading>()?.Fill?.Value?.ToUpperInvariant() ?? string.Empty;
        set {
            string normalized = NormalizeTableShadingColor(value, nameof(ShadingFillColorHex));
            if (normalized.Length == 0) { _tableProperties.GetFirstChild<Shading>()?.Remove(); return; }
            Shading shading = GetOrCreateTableShading();
            shading.Fill = normalized;
            shading.ThemeFill = null;
            shading.ThemeFillTint = null;
            shading.ThemeFillShade = null;
        }
    }

    /// <summary>Gets or sets the direct table fill color. Null removes direct table shading.</summary>
    public Color? ShadingFillColor {
        get => ParseTableShadingColor(ShadingFillColorHex);
        set => ShadingFillColorHex = value?.ToRgbHex() ?? string.Empty;
    }

    /// <summary>Gets or sets the direct shading pattern color as RGB hex or <c>auto</c>.</summary>
    /// <remarks>Empty or null clears only the pattern color, retaining the fill and pattern.</remarks>
    public string ShadingColorHex {
        get => _tableProperties.GetFirstChild<Shading>()?.Color?.Value?.ToUpperInvariant() ?? string.Empty;
        set {
            string normalized = NormalizeTableShadingColor(value, nameof(ShadingColorHex));
            Shading? shading = normalized.Length == 0 ? _tableProperties.GetFirstChild<Shading>() : GetOrCreateTableShading();
            if (shading == null) return;
            if (normalized.Length == 0) shading.Color = null;
            else shading.Color = normalized;
            shading.ThemeColor = null;
            shading.ThemeTint = null;
            shading.ThemeShade = null;
        }
    }

    /// <summary>Gets or sets the direct shading pattern color. Null clears only that color.</summary>
    public Color? ShadingColor {
        get => ParseTableShadingColor(ShadingColorHex);
        set => ShadingColorHex = value?.ToRgbHex() ?? string.Empty;
    }

    /// <summary>Gets or sets the direct table shading pattern. Null removes direct table shading.</summary>
    public WordShadingPattern? ShadingPattern {
        get => _tableProperties.GetFirstChild<Shading>()?.Val?.Value.ToOfficeEnum();
        set {
            if (!value.HasValue) { _tableProperties.GetFirstChild<Shading>()?.Remove(); return; }
            ShadingPatternValues pattern = value.Value.ToOpenXml();
            GetOrCreateTableShading().Val = pattern;
        }
    }

    private Shading GetOrCreateTableShading() {
        Shading? shading = _tableProperties.GetFirstChild<Shading>();
        if (shading == null) {
            shading = new Shading { Val = ShadingPatternValues.Clear };
            _tableProperties.AddChild(shading, true);
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
