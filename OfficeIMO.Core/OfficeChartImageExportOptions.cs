namespace OfficeIMO.Drawing;

/// <summary>
/// Options for rendering one native Word, Excel or PowerPoint chart as an image.
/// </summary>
/// <remarks>
/// The chart is drawn on its own, without the page, worksheet or slide around it. Scheme colors and series
/// without explicit colors resolve against the theme of the document that owns the chart. Chart-specific
/// rendering settings belong on this type so every chart owner reads them the same way.
/// </remarks>
public sealed class OfficeChartImageExportOptions : OfficeImageExportOptions {
    /// <summary>Rendered chart width. Null uses the width of the chart's frame or anchor in its document.</summary>
    public OfficeLength? Width { get; set; }

    /// <summary>Rendered chart height. Null uses the height of the chart's frame or anchor in its document.</summary>
    public OfficeLength? Height { get; set; }

    internal OfficeChartImageExportOptions Clone() {
        OfficeChartImageExportOptions clone = CopyImageExportOptionsTo(new OfficeChartImageExportOptions());
        clone.Width = Width;
        clone.Height = Height;
        return clone;
    }

    internal void Validate() {
        ValidateImageExportOptions();
        if (Width is { } width && width.ToEmus() <= 0) throw new ArgumentOutOfRangeException(nameof(Width), "Chart image width must be positive.");
        if (Height is { } height && height.ToEmus() <= 0) throw new ArgumentOutOfRangeException(nameof(Height), "Chart image height must be positive.");
    }
}
