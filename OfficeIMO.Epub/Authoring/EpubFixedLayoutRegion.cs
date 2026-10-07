namespace OfficeIMO.Epub;

/// <summary>
/// An absolute border-box rectangle for an existing top-level XHTML body element.
/// Coordinates are CSS pixels from the canvas's top-left corner, independent of reading direction.
/// </summary>
public sealed class EpubFixedLayoutRegion {
    /// <summary>Creates a region; its target and bounds are validated when configuring the page.</summary>
    public EpubFixedLayoutRegion(string elementId, decimal left, decimal top, decimal width, decimal height) {
        ElementId = elementId ?? throw new ArgumentNullException(nameof(elementId));
        Left = left; Top = top; Width = width; Height = height;
    }
    /// <summary>The existing element's unqualified HTML id attribute.</summary>
    public string ElementId { get; }
    /// <summary>Distance from the canvas's left edge in CSS pixels.</summary>
    public decimal Left { get; }
    /// <summary>Distance from the canvas's top edge in CSS pixels.</summary>
    public decimal Top { get; }
    /// <summary>Border-box width in CSS pixels, including padding and borders.</summary>
    public decimal Width { get; }
    /// <summary>Border-box height in CSS pixels, including padding and borders.</summary>
    public decimal Height { get; }
}
