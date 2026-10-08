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
    Webp = 4,
    /// <summary>Windows V4 bitmap with explicit alpha masks.</summary>
    Bmp = 5,
    /// <summary>Binary bilevel Portable Bitmap; pixels are composited over white before thresholding.</summary>
    Pbm = 6,
    /// <summary>Uncompressed 32-bit Truevision TGA with straight alpha.</summary>
    Tga = 7,
    /// <summary>Windows icon with a PNG-compressed entry.</summary>
    Icon = 8
}
