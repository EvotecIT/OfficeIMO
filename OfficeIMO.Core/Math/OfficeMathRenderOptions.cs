using System;

namespace OfficeIMO.Drawing;

/// <summary>Settings for dependency-free mathematical layout and drawing.</summary>
public sealed class OfficeMathRenderOptions {
    /// <summary>Base mathematical font, whose size is expressed in points.</summary>
    public OfficeFontInfo Font { get; set; } = new OfficeFontInfo("Cambria Math", 18D);

    /// <summary>Foreground color.</summary>
    public OfficeColor Color { get; set; } = OfficeColor.Black;

    /// <summary>Optional canvas background.</summary>
    public OfficeColor? BackgroundColor { get; set; }

    /// <summary>Padding around the expression in drawing units.</summary>
    public double Padding { get; set; } = 8D;

    /// <summary>Uses available static OpenType MATH constants for supported geometry. Disable
    /// to use caller-supplied script and rule settings throughout.</summary>
    public bool UseFontMathMetrics { get; set; } = true;

    /// <summary>Fallback relative scale applied to scripts, limits, and root indices.</summary>
    public double ScriptScale { get; set; } = 0.71D;

    /// <summary>Fallback minimum gap around compact fraction and decoration rules.
    /// Display fractions use three times this gap.</summary>
    public double RuleGap { get; set; } = 2D;

    /// <summary>Fallback thickness of fraction and bar rules; thickness of radical and box rules.</summary>
    public double RuleThickness { get; set; } = 1D;

    /// <summary>Horizontal and vertical gap between matrix cells.</summary>
    public double MatrixGap { get; set; } = 8D;

    /// <summary>Point-to-drawing-unit density used by both measurement and paint.
    /// The default 72 keeps one drawing unit per font point.</summary>
    public double Dpi { get; set; } = 72D;

    /// <summary>Uses display fractions with full-size children. Compact fractions and
    /// fractions inside scripts use available font script percentages or the fallback <see cref="ScriptScale"/>.</summary>
    public bool DisplayStyle { get; set; } = true;

    /// <summary>Scoped font programs used for deterministic advances and painted glyph bounds.</summary>
    public OfficeFontFaceCollection Fonts { get; set; } = new OfficeFontFaceCollection();

    // Computed token presentation supplied by document adapters. Public semantic text
    // and serialization remain untouched; measurement and paint use the same glyphs.
    internal Func<OfficeMathExpression, string?>? TokenPaintText { get; set; }

    /// <summary>Creates a detached copy.</summary>
    public OfficeMathRenderOptions Clone() => new OfficeMathRenderOptions {
        TokenPaintText = TokenPaintText,
        Font = Font,
        Color = Color,
        BackgroundColor = BackgroundColor,
        Padding = Padding,
        UseFontMathMetrics = UseFontMathMetrics,
        ScriptScale = ScriptScale,
        RuleGap = RuleGap,
        RuleThickness = RuleThickness,
        MatrixGap = MatrixGap,
        Dpi = Dpi,
        DisplayStyle = DisplayStyle,
        Fonts = Fonts?.Clone() ?? new OfficeFontFaceCollection()
    };

    internal void Validate() {
        Positive(Font.Size, nameof(Font));
        NonNegative(Padding, nameof(Padding));
        Positive(ScriptScale, nameof(ScriptScale));
        NonNegative(RuleGap, nameof(RuleGap));
        Positive(RuleThickness, nameof(RuleThickness));
        NonNegative(MatrixGap, nameof(MatrixGap));
        Positive(Dpi, nameof(Dpi));
    }

    private static void Positive(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) throw new ArgumentOutOfRangeException(name);
    }

    private static void NonNegative(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) throw new ArgumentOutOfRangeException(name);
    }
}
