namespace OfficeIMO.Drawing;

/// <summary>
/// Image formats supported by OfficeIMO dependency-free export pipelines.
/// </summary>
public enum OfficeImageExportFormat {
    /// <summary>Portable Network Graphics raster output.</summary>
    Png = 0,

    /// <summary>Scalable Vector Graphics XML output.</summary>
    Svg = 1,

    /// <summary>Joint Photographic Experts Group raster output.</summary>
    Jpeg = 2,

    /// <summary>Tagged Image File Format raster output.</summary>
    Tiff = 3,

    /// <summary>Lossless WebP raster output.</summary>
    Webp = 4
}
