using System;

namespace OfficeIMO.Drawing;

/// <summary>Immutable text-style overrides for a textual tab leader. Unspecified properties inherit the tab run.</summary>
public sealed class OfficeTextTabLeaderStyle {
    /// <summary>Creates a leader style using an absolute font size or a multiplier of the active tab run's size.</summary>
    /// <param name="fontSizePoints">Positive absolute font size before drawing scale and frame fitting; null uses <paramref name="fontSizeFactor"/>.</param>
    /// <param name="fontSizeFactor">Positive multiplier of the active font size when no absolute size is supplied.</param>
    /// <param name="color">Optional complete foreground color, including alpha.</param>
    /// <param name="fontFamily">Optional font family.</param>
    /// <param name="bold">Optional bold override.</param>
    /// <param name="italic">Optional italic override.</param>
    /// <param name="underlineStyle">Optional underline override; None removes inherited underline.</param>
    /// <param name="strikethroughStyle">Optional strike-through override; None removes inherited strike-through.</param>
    /// <param name="baseline">Optional baseline override.</param>
    /// <param name="backgroundColor">Optional background override; transparent removes inherited background paint.</param>
    /// <param name="inheritOpacity">Use the active run's alpha, including when color overrides its RGB.</param>
    /// <param name="opacity">Optional alpha from zero to one; takes precedence over inherited or color alpha.</param>
    public OfficeTextTabLeaderStyle(double? fontSizePoints = null, double fontSizeFactor = 1D, OfficeColor? color = null,
        string? fontFamily = null, bool? bold = null, bool? italic = null, OfficeTextDecorationStyle? underlineStyle = null,
        OfficeTextDecorationStyle? strikethroughStyle = null, OfficeTextBaseline? baseline = null, OfficeColor? backgroundColor = null,
        bool inheritOpacity = false, double? opacity = null) {
        if (fontSizePoints.HasValue && !PositiveFinite(fontSizePoints.Value)) throw new ArgumentOutOfRangeException(nameof(fontSizePoints));
        if (!PositiveFinite(fontSizeFactor)) throw new ArgumentOutOfRangeException(nameof(fontSizeFactor));
        if (opacity.HasValue && (opacity.Value < 0 || opacity.Value > 1 || double.IsNaN(opacity.Value))) throw new ArgumentOutOfRangeException(nameof(opacity));
        if (fontFamily != null && string.IsNullOrWhiteSpace(fontFamily)) throw new ArgumentException("A font family cannot be blank.", nameof(fontFamily));
        // Reuse the canonical run's decoration and baseline validation.
        _ = new OfficeRichTextRun(string.Empty, 1, OfficeColor.Black, underlineStyle: underlineStyle ?? OfficeTextDecorationStyle.None,
            strikethroughStyle: strikethroughStyle ?? OfficeTextDecorationStyle.None, baseline: baseline ?? OfficeTextBaseline.Normal);
        FontSizePoints = fontSizePoints; FontSizeFactor = fontSizeFactor; Color = color; FontFamily = fontFamily;
        Bold = bold; Italic = italic; UnderlineStyle = underlineStyle; StrikethroughStyle = strikethroughStyle;
        Baseline = baseline; BackgroundColor = backgroundColor;
        InheritOpacity = inheritOpacity; Opacity = opacity;
    }
    /// <summary>Absolute font size before drawing scale and frame fitting, or null for a relative size.</summary>
    public double? FontSizePoints { get; }
    /// <summary>Multiplier of the active tab run's font size when no absolute size is supplied.</summary>
    public double FontSizeFactor { get; }
    /// <summary>Foreground override; alpha follows <see cref="Opacity"/> and <see cref="InheritOpacity"/> when supplied.</summary>
    public OfficeColor? Color { get; }
    /// <summary>Font family override, or null to inherit.</summary>
    public string? FontFamily { get; }
    /// <summary>Bold override, or null to inherit.</summary>
    public bool? Bold { get; }
    /// <summary>Italic override, or null to inherit.</summary>
    public bool? Italic { get; }
    /// <summary>Underline override, or null to inherit.</summary>
    public OfficeTextDecorationStyle? UnderlineStyle { get; }
    /// <summary>Strike-through override, or null to inherit.</summary>
    public OfficeTextDecorationStyle? StrikethroughStyle { get; }
    /// <summary>Baseline override, or null to inherit.</summary>
    public OfficeTextBaseline? Baseline { get; }
    /// <summary>Background override, or null to inherit.</summary>
    public OfficeColor? BackgroundColor { get; }

    /// <summary>Whether the active run's alpha is retained when no opacity override is supplied.</summary>
    public bool InheritOpacity { get; }
    /// <summary>Independent alpha override from zero to one, or null.</summary>
    public double? Opacity { get; }
    internal OfficeTextTabLeaderStyle Scale(double scale) => FontSizePoints.HasValue
        ? Copy(Math.Max(double.Epsilon, Math.Min(double.MaxValue, FontSizePoints.Value * scale)), Color, BackgroundColor) : this;
    internal OfficeTextTabLeaderStyle Tint(OfficeColor color) => Copy(FontSizePoints,
        Color.HasValue ? OfficeColor.FromRgba(color.R, color.G, color.B, Color.Value.A) : null,
        BackgroundColor.HasValue ? OfficeColor.FromRgba(color.R, color.G, color.B, BackgroundColor.Value.A) : null);
    private OfficeTextTabLeaderStyle Copy(double? size, OfficeColor? color, OfficeColor? background) => new OfficeTextTabLeaderStyle(size, FontSizeFactor,
        color, FontFamily, Bold, Italic, UnderlineStyle, StrikethroughStyle, Baseline, background, InheritOpacity, Opacity);
    internal OfficeRichTextRun? Resolve(OfficeRichTextRun active, string glyph) {
        double size = FontSizePoints ?? active.FontSize * FontSizeFactor;
        if (!PositiveFinite(size)) return null;
        OfficeColor color = Color ?? active.Color;
        if (Opacity.HasValue) color = OfficeColorTransforms.WithAlpha(color, Opacity.Value);
        else if (InheritOpacity) color = OfficeColor.FromRgba(color.R, color.G, color.B, active.Color.A);
        return new OfficeRichTextRun(glyph, size, color, Bold ?? active.Bold, Italic ?? active.Italic,
            fontFamily: FontFamily ?? active.FontFamily, backgroundColor: BackgroundColor ?? active.BackgroundColor,
            underlineStyle: UnderlineStyle ?? active.UnderlineStyle, strikethroughStyle: StrikethroughStyle ?? active.StrikethroughStyle,
            baseline: Baseline ?? active.Baseline) { LinkUri = active.LinkUri };
    }
    private static bool PositiveFinite(double value) => value > 0 && !double.IsNaN(value) && !double.IsInfinity(value);
}
