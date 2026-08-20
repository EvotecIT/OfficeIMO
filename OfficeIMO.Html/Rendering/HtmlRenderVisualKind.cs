namespace OfficeIMO.Html;

/// <summary>
/// Identifies the backend-neutral visual operation produced by HTML layout.
/// </summary>
public enum HtmlRenderVisualKind {
    /// <summary>Vector shape such as a background or border.</summary>
    Shape = 0,

    /// <summary>Positioned searchable text.</summary>
    Text = 1,

    /// <summary>Positioned raster or vector image.</summary>
    Image = 2,

    /// <summary>Clipped repeating image pattern.</summary>
    ImagePattern = 3,

    /// <summary>Ordered child visuals clipped as one paint group.</summary>
    ClipGroup = 4,

    /// <summary>Ordered child visuals clipped by shared Drawing path geometry.</summary>
    PathClipGroup = 5,

    /// <summary>Ordered child visuals painted through an affine transform and isolated opacity.</summary>
    EffectGroup = 6,

    /// <summary>Positioned shared vector drawing.</summary>
    Drawing = 7,

    /// <summary>Paint-neutral semantic ownership group.</summary>
    SemanticGroup = 8,

    /// <summary>Paint-neutral positioned fragments sharing one logical extraction string.</summary>
    LogicalTextGroup = 9,

    /// <summary>Standard HTML form semantics with ordered static fallback visuals.</summary>
    FormField = 10,

    /// <summary>Paint-neutral navigation destination for an element without searchable text visuals.</summary>
    BookmarkAnchor = 11,

    /// <summary>Paint-neutral editable layout region for native destination projection.</summary>
    LayoutRegion = 12
}
