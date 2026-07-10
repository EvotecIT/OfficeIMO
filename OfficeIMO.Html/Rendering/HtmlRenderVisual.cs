namespace OfficeIMO.Html;

/// <summary>
/// Immutable positioned visual emitted by the shared HTML layout engine.
/// </summary>
public abstract class HtmlRenderVisual {
    internal HtmlRenderVisual(HtmlRenderVisualKind kind, double x, double y, double width, double height, int paintOrder, string? linkUri, string? source) {
        ValidateFinite(x, nameof(x));
        ValidateFinite(y, nameof(y));
        ValidatePositive(width, nameof(width));
        ValidatePositive(height, nameof(height));
        Kind = kind;
        X = x;
        Y = y;
        Width = width;
        Height = height;
        PaintOrder = paintOrder;
        LinkUri = linkUri;
        Source = source;
    }

    /// <summary>Visual operation kind.</summary>
    public HtmlRenderVisualKind Kind { get; }

    /// <summary>Left coordinate in CSS pixels.</summary>
    public double X { get; }

    /// <summary>Top coordinate in CSS pixels.</summary>
    public double Y { get; }

    /// <summary>Visual width in CSS pixels.</summary>
    public double Width { get; }

    /// <summary>Visual height in CSS pixels.</summary>
    public double Height { get; }

    /// <summary>Stable paint order within the rendered page.</summary>
    public int PaintOrder { get; }

    /// <summary>Optional safe hyperlink associated with the visual.</summary>
    public string? LinkUri { get; }

    /// <summary>Optional source element description used by diagnostics and inspection.</summary>
    public string? Source { get; }

    internal abstract HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder);

    private static void ValidateFinite(double value, string parameterName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "Render coordinates must be finite.");
        }
    }

    private static void ValidatePositive(double value, string parameterName) {
        if (value <= 0D || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(parameterName, "Render dimensions must be finite positive numbers.");
        }
    }
}
