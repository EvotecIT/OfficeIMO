namespace OfficeIMO.ChartForgeX;

/// <summary>Controls how ChartForgeX SVG features that OfficeIMO.Drawing cannot preserve are handled.</summary>
public enum OfficeVisualSvgPolicy {
    /// <summary>Keep the imported vector scene and report any unsupported SVG features.</summary>
    PreserveVector,
    /// <summary>Rasterize the artifact when the SVG importer reports unsupported features.</summary>
    RasterizeWhenNeeded,
    /// <summary>Reject an artifact unless its SVG can be represented completely as an Office drawing.</summary>
    RequireVector
}
