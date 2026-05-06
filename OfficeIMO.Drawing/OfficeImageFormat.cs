namespace OfficeIMO.Drawing;

/// <summary>
/// Image formats understood by the shared OfficeIMO drawing layer.
/// </summary>
public enum OfficeImageFormat {
    /// <summary>Unknown or unsupported image format.</summary>
    Unknown,
    /// <summary>Portable Network Graphics image.</summary>
    Png,
    /// <summary>JPEG image.</summary>
    Jpeg,
    /// <summary>Graphics Interchange Format image.</summary>
    Gif,
    /// <summary>Bitmap image.</summary>
    Bmp,
    /// <summary>Tagged Image File Format image.</summary>
    Tiff,
    /// <summary>Scalable Vector Graphics image.</summary>
    Svg,
    /// <summary>Enhanced Metafile image.</summary>
    Emf,
    /// <summary>Windows Metafile image.</summary>
    Wmf,
    /// <summary>Icon image.</summary>
    Icon,
    /// <summary>PCX image.</summary>
    Pcx
}
