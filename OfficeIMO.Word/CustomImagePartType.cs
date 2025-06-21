using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word;

public enum CustomImagePartType {
    Bmp,
    Gif,
    Jpeg,
    Png,
    Tiff,
    Emf
}

public static class CustomImagePartTypeExtensions {
    public static string ToOpenXmlImagePartType(this CustomImagePartType customType) {
        return customType switch {
            CustomImagePartType.Bmp => "image/bmp",
            CustomImagePartType.Gif => "image/gif",
            CustomImagePartType.Jpeg => "image/jpeg",
            CustomImagePartType.Png => "image/png",
            CustomImagePartType.Tiff => "image/tiff",
            CustomImagePartType.Emf => "image/x-emf",
            _ => throw new ArgumentOutOfRangeException(nameof(customType), customType, null)
        };
    }
}
