using static OfficeIMO.Word.OpenXml.WordXml;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Admits a rectangular inline VML picture without extending support to floating shapes or picture effects.</summary>
    private static bool IsNativeInlineVmlPicture(OfficeXmlNode shape) {
        if (shape.IsNone || !shape.Is(V, "shape") || !shape.Parent.IsW("pict") || shape.Attr("type") != "#_x0000_t75" ||
            shape.Ancestor(V, "group").Exists || shape.Elements().Count() != 1 ||
            shape.FirstChild is not { IsNone: false } image || !image.Is(V, "imagedata") ||
            WordVmlImageGeometry.ReadDimension(shape.Attr("style"), "width") is not double width || width <= 0D ||
            WordVmlImageGeometry.ReadDimension(shape.Attr("style"), "height") is not double height || height <= 0D)
            return false;

        // Current image style mapping has no VML crop/effect or floating-frame
        // projection. Such sources retain their established bounded fallback.
        if (shape.Attributes().Any(attribute => attribute.LocalName is not ("id" or "type" or "style" or "alt")) ||
            image.Attributes().Any(attribute => attribute.LocalName is not ("id" or "title"))) return false;
        try {
            return WordConverterNodes.VmlShapeStyle(shape).All(pair =>
                pair.Key.Equals("width", StringComparison.OrdinalIgnoreCase) ||
                pair.Key.Equals("height", StringComparison.OrdinalIgnoreCase) ||
                (pair.Key.Equals("mso-wrap-style", StringComparison.OrdinalIgnoreCase) &&
                    pair.Value.Equals("square", StringComparison.OrdinalIgnoreCase)));
        } catch (ArgumentException) {
            // Duplicate declarations do not establish an ordinary inline frame.
            return false;
        }
    }

    /// <summary>Keeps the running-story admission and paragraph image projection on the same VML profile.</summary>
    private static bool IsNativeRunningVmlPicture(OfficeXmlNode picture) {
        OfficeXmlNode[] shapes = picture.Elements(V, "shape").ToArray();
        return shapes.Length == 1 && IsNativeInlineVmlPicture(shapes[0]) &&
            picture.Elements().All(child => child.Is(V, "shapetype") || child == shapes[0]);
    }
}
