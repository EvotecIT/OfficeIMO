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

    /// <summary>Relative scale applied to scripts, limits, and root indices.</summary>
    public double ScriptScale { get; set; } = 0.71D;

    /// <summary>Gap around fraction and decoration rules.</summary>
    public double RuleGap { get; set; } = 2D;

    /// <summary>Thickness of fraction, radical, bar, and box rules.</summary>
    public double RuleThickness { get; set; } = 1D;

    /// <summary>Horizontal and vertical gap between matrix cells.</summary>
    public double MatrixGap { get; set; } = 8D;

    /// <summary>Point-to-drawing-unit density used by both measurement and paint.
    /// The default 72 keeps one drawing unit per font point.</summary>
    public double Dpi { get; set; } = 72D;

    /// <summary>Uses display fractions with full-size children. Compact fractions and
    /// fractions inside scripts reduce children by <see cref="ScriptScale"/>.</summary>
    public bool DisplayStyle { get; set; } = true;

    /// <summary>Scoped font programs used for deterministic advances and painted glyph bounds.</summary>
    public OfficeFontFaceCollection Fonts { get; set; } = new OfficeFontFaceCollection();

    /// <summary>Creates a detached copy.</summary>
    public OfficeMathRenderOptions Clone() => new OfficeMathRenderOptions {
        Font = Font,
        Color = Color,
        BackgroundColor = BackgroundColor,
        Padding = Padding,
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
