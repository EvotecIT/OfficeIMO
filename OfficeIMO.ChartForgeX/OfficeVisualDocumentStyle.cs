using System;
using global::ChartForgeX.Primitives;
using global::ChartForgeX.Rendering;
using global::ChartForgeX.Themes;

namespace OfficeIMO.ChartForgeX;

/// <summary>Coordinates chart preparation and document text through a point-based typography scale.</summary>
/// <remarks>Prepare at the intended placement size for readable labels. Resizing an already prepared artifact scales its text instead of reflowing it.</remarks>
public sealed class OfficeVisualDocumentStyle {
    private const double PointsPerPixel = 0.75D;

    /// <summary>Gets the shared Arial document style with 18-point headings and 11.25-point chart labels.</summary>
    public static OfficeVisualDocumentStyle Default { get; } = new OfficeVisualDocumentStyle();

    /// <summary>Creates an immutable document typography style. Sizes are expressed in points.</summary>
    public OfficeVisualDocumentStyle(string fontFamily = "Arial", double headingSizePoints = 18D,
        double captionSizePoints = 11D, double labelSizePoints = 11.25D, double subtitleSizePoints = 12D) {
        if (string.IsNullOrWhiteSpace(fontFamily)) throw new ArgumentException("A font family is required.", nameof(fontFamily));
        FontFamily = fontFamily;
        HeadingSizePoints = Positive(headingSizePoints, nameof(headingSizePoints));
        CaptionSizePoints = Positive(captionSizePoints, nameof(captionSizePoints));
        LabelSizePoints = Positive(labelSizePoints, nameof(labelSizePoints));
        SubtitleSizePoints = Positive(subtitleSizePoints, nameof(subtitleSizePoints));
    }

    /// <summary>Gets the font family used by chart preparation and document text.</summary>
    public string FontFamily { get; }
    /// <summary>Gets the heading size in points.</summary>
    public double HeadingSizePoints { get; }
    /// <summary>Gets the caption size in points.</summary>
    public double CaptionSizePoints { get; }
    /// <summary>Gets the chart axis, legend and data-label size in points.</summary>
    public double LabelSizePoints { get; }
    /// <summary>Gets the chart subtitle size in points.</summary>
    public double SubtitleSizePoints { get; }

    /// <summary>Creates a ChartForgeX context at the intended document size, using this typography and the shared paired palette.</summary>
    /// <param name="widthPoints">Intended document width in points.</param>
    /// <param name="heightPoints">Intended document height in points.</param>
    /// <param name="themeMode">Light or dark paired palette.</param>
    /// <param name="frame">Optional title, subtitle, legend and surface configuration.</param>
    /// <param name="theme">Optional palette and geometry. This style supplies its typography.</param>
    public VisualRenderContext CreateContext(double widthPoints = 450D, double heightPoints = 270D,
        VisualThemeMode themeMode = VisualThemeMode.Light, VisualFrame? frame = null, VisualTheme? theme = null) {
        Positive(widthPoints, nameof(widthPoints));
        Positive(heightPoints, nameof(heightPoints));
        var typography = new VisualTypography(FontFamily, HeadingSizePoints / PointsPerPixel,
            SubtitleSizePoints / PointsPerPixel, LabelSizePoints / PointsPerPixel,
            LabelSizePoints / PointsPerPixel, LabelSizePoints / PointsPerPixel);
        return new VisualRenderContext(
            new VisualLayoutOptions(new VisualSize(widthPoints / PointsPerPixel, heightPoints / PointsPerPixel), padding: 24D),
            (theme ?? VisualTheme.Graphite()).WithTypography(typography), themeMode, frame);
    }

    private static double Positive(double value, string name) {
        if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(name);
        return value;
    }
}
