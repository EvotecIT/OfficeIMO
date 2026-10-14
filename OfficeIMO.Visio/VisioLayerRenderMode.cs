namespace OfficeIMO.Visio;

/// <summary>Controls which layer members appear in SVG and raster previews.</summary>
public enum VisioLayerRenderMode {
    /// <summary>Uses each layer's Visible flag. This is the default screen preview.</summary>
    Visible = 0,

    /// <summary>Uses each layer's Print flag, independently of screen visibility.</summary>
    Printable = 1,

    /// <summary>Includes all layer members regardless of their Visible and Print flags.</summary>
    All = 2
}
