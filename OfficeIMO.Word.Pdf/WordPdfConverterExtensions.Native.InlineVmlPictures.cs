using System;
using System.Linq;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Admits a rectangular inline VML picture without extending support to floating shapes or picture effects.</summary>
    private static bool IsNativeInlineVmlPicture(V.Shape? shape) {
        if (shape == null || shape.Parent is not W.Picture || shape.Type?.Value != "#_x0000_t75" ||
            shape.Ancestors<V.Group>().Any() || shape.ChildElements.Count != 1 ||
            shape.ChildElements[0] is not V.ImageData image ||
            WordVmlImageGeometry.ReadDimension(shape.Style?.Value, "width") is not double width || width <= 0D ||
            WordVmlImageGeometry.ReadDimension(shape.Style?.Value, "height") is not double height || height <= 0D)
            return false;

        // Current image style mapping has no VML crop/effect or floating-frame
        // projection. Such sources retain their established bounded fallback.
        if (shape.GetAttributes().Any(attribute => attribute.LocalName is not ("id" or "type" or "style" or "alt")) ||
            image.GetAttributes().Any(attribute => attribute.LocalName is not ("id" or "title"))) return false;
        try {
            return ImageShapeStyleHelper.GetStyle(shape).All(pair =>
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
    private static bool IsNativeRunningVmlPicture(W.Picture picture) {
        V.Shape[] shapes = picture.Elements<V.Shape>().ToArray();
        return shapes.Length == 1 && IsNativeInlineVmlPicture(shapes[0]) &&
            picture.ChildElements.All(child => child is V.Shapetype || ReferenceEquals(child, shapes[0]));
    }
}
