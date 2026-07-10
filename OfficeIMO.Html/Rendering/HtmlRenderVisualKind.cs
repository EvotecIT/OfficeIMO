namespace OfficeIMO.Html;

/// <summary>
/// Identifies the backend-neutral visual operation produced by HTML layout.
/// </summary>
public enum HtmlRenderVisualKind {
    /// <summary>Vector shape such as a background or border.</summary>
    Shape,

    /// <summary>Positioned searchable text.</summary>
    Text,

    /// <summary>Positioned raster or vector image.</summary>
    Image
}
