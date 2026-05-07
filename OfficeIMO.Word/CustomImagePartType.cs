namespace OfficeIMO.Word;

/// <summary>
/// Enumeration of additional image types supported by the library.
/// </summary>
public enum CustomImagePartType {
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
/// Extension helpers for <see cref="CustomImagePartType"/> values.
/// </summary>
public static class CustomImagePartTypeExtensions {
    /// <summary>
    /// Converts the custom image part type to the Open XML content type string.
    /// </summary>
    /// <param name="customType">The custom image part type.</param>
    /// <returns>The corresponding content type value.</returns>
    public static string ToOpenXmlImagePartType(this CustomImagePartType customType) {
        return customType switch {
            CustomImagePartType.Bmp => "image/bmp",
            CustomImagePartType.Gif => "image/gif",
            CustomImagePartType.Jpeg => "image/jpeg",
            CustomImagePartType.Png => "image/png",
            CustomImagePartType.Tiff => "image/tiff",
            CustomImagePartType.Emf => "image/x-emf",
            CustomImagePartType.Wmf => "image/x-wmf",
            CustomImagePartType.Svg => "image/svg+xml",
            _ => throw new ArgumentOutOfRangeException(nameof(customType), customType, null)
        };
    }
}
