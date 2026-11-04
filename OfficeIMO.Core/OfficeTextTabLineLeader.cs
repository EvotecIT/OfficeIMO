using System;

namespace OfficeIMO.Drawing;

/// <summary>Line patterns available inside a paragraph tab gap.</summary>
public enum OfficeTextTabLineLeaderStyle {
    /// <summary>No line paint.</summary>
    None,
    /// <summary>A continuous line.</summary>
    Solid,
    /// <summary>Round dots.</summary>
    Dotted,
    /// <summary>Short dashes.</summary>
    Dash,
    /// <summary>Long dashes.</summary>
    LongDash,
    /// <summary>Alternating dots and dashes.</summary>
    DotDash,
    /// <summary>Two dots between successive dashes.</summary>
    DotDotDash,
    /// <summary>A sinusoidal line.</summary>
    Wave
}

/// <summary>Immutable vector paint settings for a tab leader.</summary>
public sealed class OfficeTextTabLineLeader {
    // Drawing scale stays numerical so fitting does not re-quantize authored lengths.
    private double? _scaledWidthPoints;
    internal double? DrawingWidthPoints => _scaledWidthPoints ?? Width?.ToPoints();
    /// <summary>Creates line paint with an absolute width or a fraction of the active rendered font size.</summary>
    /// <param name="style">Pattern drawn within the gap.</param>
    /// <param name="doubleLine">Draw two parallel copies of the pattern.</param>
    /// <param name="width">Positive absolute width before drawing scale; null uses <paramref name="widthFontFraction"/>.</param>
    /// <param name="widthFontFraction">Positive fraction of the active rendered font size; the default is 0.05.</param>
    /// <param name="color">Paint color; null uses the tab's active text color.</param>
    public OfficeTextTabLineLeader(OfficeTextTabLineLeaderStyle style = OfficeTextTabLineLeaderStyle.Solid,
        bool doubleLine = false, OfficeLength? width = null, double widthFontFraction = .05D, OfficeColor? color = null) {
        if (!Enum.IsDefined(typeof(OfficeTextTabLineLeaderStyle), style)) throw new ArgumentOutOfRangeException(nameof(style));
        if (width.HasValue && width.Value <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
        if (!PositiveFinite(widthFontFraction)) throw new ArgumentOutOfRangeException(nameof(widthFontFraction));
        Style = style; DoubleLine = doubleLine; Width = width; WidthFontFraction = widthFontFraction; Color = color;
    }
    /// <summary>Pattern drawn in the gap.</summary>
    public OfficeTextTabLineLeaderStyle Style { get; }
    /// <summary>Whether two parallel copies are drawn.</summary>
    public bool DoubleLine { get; }
    /// <summary>Absolute width before drawing scale, or null for a font-relative width.</summary>
    public OfficeLength? Width { get; }
    /// <summary>Width as a fraction of the active rendered font size when no absolute width is supplied.</summary>
    public double WidthFontFraction { get; }
    /// <summary>Explicit paint color, or null for the active text color.</summary>
    public OfficeColor? Color { get; }
    internal OfficeTextTabLineLeader Scale(double scale) {
        if (!Width.HasValue) return this;
        // Keep extreme accepted widths representable so the bounded planner can
        // suppress unpaintable geometry without aborting the text-frame layout.
        double width = Math.Max(double.Epsilon, Math.Min(double.MaxValue, DrawingWidthPoints!.Value * scale));
        return new OfficeTextTabLineLeader(Style, DoubleLine, Width, WidthFontFraction, Color) {
            _scaledWidthPoints = width
        };
    }
    private static bool PositiveFinite(double value) => value > 0 && !double.IsNaN(value) && !double.IsInfinity(value);
}
