using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word;

public enum CustomImagePartType {
    Bmp,
    Gif,
    Jpeg,
    Png,
    Tiff
}

public static class CustomImagePartTypeExtensions {
    /// <summary>
    /// Converts the custom image type enumeration into an Open XML content type string.
    /// </summary>
    public static string ToOpenXmlImagePartType(this CustomImagePartType customType) {
        return customType switch {
            CustomImagePartType.Bmp => "image/bmp",
            CustomImagePartType.Gif => "image/gif",
            CustomImagePartType.Jpeg => "image/jpeg",
            CustomImagePartType.Png => "image/png",
            CustomImagePartType.Tiff => "image/tiff",
            _ => throw new ArgumentOutOfRangeException(nameof(customType), customType, null)
        };
    }
}
