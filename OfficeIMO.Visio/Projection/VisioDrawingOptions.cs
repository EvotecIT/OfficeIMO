using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

/// <summary>Controls the cached Visio page projection into point-sized shared Drawing scenes.</summary>
public sealed class VisioDrawingOptions {
    /// <summary>Optional source name recorded in the operation report.</summary>
    public string? SourceName { get; set; }

    /// <summary>Layer selection. The default selects printable content independently of screen visibility.</summary>
    public VisioLayerRenderMode LayerMode { get; set; } = VisioLayerRenderMode.Printable;

    /// <summary>Maximum source page count and pages in a single background composition chain. Default: 1,000.</summary>
    public int MaximumPages { get; set; } = 1000;

    /// <summary>Maximum shapes and connectors visited, including excluded layers. Default: 100,000.</summary>
    public int MaximumShapes { get; set; } = 100_000;

    /// <summary>Maximum cached or generated geometry points projected across all pages. Default: 1,000,000.</summary>
    public int MaximumGeometryPoints { get; set; } = 1_000_000;

    /// <summary>Maximum decoded bitmap pixels retained across all projected image instances. Default: 64 million.</summary>
    public long MaximumTotalImagePixels { get; set; } = 64_000_000;

    /// <summary>Maximum PNG payload bytes retained across all projected image instances. Default: 128 MiB.</summary>
    public long MaximumTotalImageBytes { get; set; } = 128 * 1024 * 1024;

    /// <summary>Reject the operation if its report contains an approximation, omission, or unassessed feature.</summary>
    public bool RequireNoLoss { get; set; }

    /// <summary>Caller-supplied font faces used by the returned scenes before platform fallback.</summary>
    public OfficeFontFaceCollection Fonts { get; } = new();

    /// <summary>Optional shared text shaping provider.</summary>
    public IOfficeTextShapingProvider? TextShapingProvider { get; set; }

    /// <summary>Optional BCP 47 language hint passed to the shared text shaper.</summary>
    public string? TextShapingLanguage { get; set; }

    // Bridges can assess clipping against the resources used by their final renderer.
    internal OfficeDrawingTextMetrics? LayoutMetrics { get; set; }

    /// <summary>Creates independent operation settings and a detached font collection.</summary>
    public VisioDrawingOptions Clone() {
        var copy = new VisioDrawingOptions {
            SourceName = SourceName, LayerMode = LayerMode, MaximumPages = MaximumPages,
            MaximumShapes = MaximumShapes, MaximumGeometryPoints = MaximumGeometryPoints,
            MaximumTotalImagePixels = MaximumTotalImagePixels, MaximumTotalImageBytes = MaximumTotalImageBytes,
            RequireNoLoss = RequireNoLoss, TextShapingProvider = TextShapingProvider,
            TextShapingLanguage = TextShapingLanguage, LayoutMetrics = LayoutMetrics
        };
        copy.Fonts.AddRange(Fonts);
        copy.Fonts.FontProgramProvider = Fonts.FontProgramProvider;
        copy.Fonts.FontVariationResolver = Fonts.FontVariationResolver;
        return copy;
    }

    internal VisioDrawingOptions Snapshot() {
        if (SourceName != null && string.IsNullOrWhiteSpace(SourceName)) throw new ArgumentException("Source name cannot be empty.", nameof(SourceName));
        if (!Enum.IsDefined(typeof(VisioLayerRenderMode), LayerMode)) throw new ArgumentOutOfRangeException(nameof(LayerMode));
        if (MaximumPages < 1) throw new ArgumentOutOfRangeException(nameof(MaximumPages));
        if (MaximumShapes < 1) throw new ArgumentOutOfRangeException(nameof(MaximumShapes));
        if (MaximumGeometryPoints < 1) throw new ArgumentOutOfRangeException(nameof(MaximumGeometryPoints));
        if (MaximumTotalImagePixels < 1) throw new ArgumentOutOfRangeException(nameof(MaximumTotalImagePixels));
        if (MaximumTotalImageBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaximumTotalImageBytes));
        return Clone();
    }
}
