namespace OfficeIMO.Pdf;

/// <summary>
/// Visual style of a panel box used by panel paragraphs.
/// </summary>
public class PanelStyle {
    /// <summary>Background fill color. Set to null for no fill.</summary>
    public PdfColor? Background { get; set; }
    /// <summary>Border color. Set to null for no border.</summary>
    public PdfColor? BorderColor { get; set; }
    /// <summary>Border stroke width in points.</summary>
    public double BorderWidth { get; set; } = 0.5;
    /// <summary>Vertical padding inside the panel (points).</summary>
    public double PaddingY { get; set; } = 6;
    /// <summary>Horizontal padding inside the panel (points).</summary>
    public double PaddingX { get; set; } = 6;
    /// <summary>Optional maximum width for the panel box (points). When set, the box can be centered or right-aligned.</summary>
    public double? MaxWidth { get; set; }
    /// <summary>Horizontal alignment of the panel box within the content area.</summary>
    public PdfAlign Align { get; set; } = PdfAlign.Left;
    /// <summary>When true, the entire panel is kept on one page; otherwise it can split across pages.</summary>
    public bool KeepTogether { get; set; }
}
