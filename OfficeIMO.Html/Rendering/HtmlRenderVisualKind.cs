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
    Image,

    /// <summary>Clipped repeating image pattern.</summary>
    ImagePattern,

    /// <summary>Ordered child visuals clipped as one paint group.</summary>
    ClipGroup,

    /// <summary>Ordered child visuals clipped by shared Drawing path geometry.</summary>
    PathClipGroup,

    /// <summary>Ordered child visuals painted through an affine transform and isolated opacity.</summary>
    EffectGroup,

    /// <summary>Positioned shared vector drawing.</summary>
    Drawing,

    /// <summary>Paint-neutral semantic ownership group.</summary>
    SemanticGroup,

    /// <summary>Paint-neutral positioned fragments sharing one logical extraction string.</summary>
    LogicalTextGroup
}
