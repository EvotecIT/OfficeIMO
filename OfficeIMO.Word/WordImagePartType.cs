using OfficeIMO.Drawing;

namespace OfficeIMO.Word;

/// <summary>
/// Enumeration of additional image types supported by the library.
/// </summary>
public enum WordImagePartType {
    /// <summary>
    /// Bitmap image type.
    /// </summary>
    Bmp,

    /// <summary>
    /// GIF image type.
    /// </summary>
    Gif,

    /// <summary>
    /// JPEG image type.
    /// </summary>
    Jpeg,

    /// <summary>
    /// PNG image type.
    /// </summary>
    Png,

    /// <summary>
    /// TIFF image type.
    /// </summary>
    Tiff,

    /// <summary>
    /// Enhanced metafile image type.
    /// </summary>
    Emf,

    /// <summary>
    /// Windows metafile image type.
    /// </summary>
    Wmf,

    /// <summary>
    /// Scalable Vector Graphics image type.
    /// </summary>
    Svg
}

/// <summary>
/// Extension helpers for <see cref="WordImagePartType"/> values.
/// </summary>
public static class WordImagePartTypeExtensions {
    /// <summary>
    /// Converts the custom image part type to the Open XML content type string.
    /// </summary>
    /// <param name="customType">The custom image part type.</param>
    /// <returns>The corresponding content type value.</returns>
    public static string ToOpenXmlImagePartType(this WordImagePartType customType) =>
        OfficeImageInfo.GetMimeType(customType.ToOfficeImageFormat());

    private static OfficeImageFormat ToOfficeImageFormat(this WordImagePartType customType) {
        return customType switch {
            WordImagePartType.Bmp => OfficeImageFormat.Bmp,
            WordImagePartType.Gif => OfficeImageFormat.Gif,
            WordImagePartType.Jpeg => OfficeImageFormat.Jpeg,
            WordImagePartType.Png => OfficeImageFormat.Png,
            WordImagePartType.Tiff => OfficeImageFormat.Tiff,
            WordImagePartType.Emf => OfficeImageFormat.Emf,
            WordImagePartType.Wmf => OfficeImageFormat.Wmf,
            WordImagePartType.Svg => OfficeImageFormat.Svg,
            _ => throw new ArgumentOutOfRangeException(nameof(customType), customType, null)
        };
    }
}
