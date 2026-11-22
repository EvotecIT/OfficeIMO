using System.Globalization;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>
    /// Projects a rectangular source picture without emitting it or changing story admission.
    /// The caller retains a stable original-source identity across deferred materializations.
    /// </summary>
    private static bool TryPrepareNativeAnchoredImageItem(object sourceIdentity, WordImage image,
        byte[] bytes, double width, double height, WordSection? section,
        PdfCore.PdfPageSize pageSize, PdfCore.PdfPageMargins margins, WordPdfExportOptions? options,
        string source, out PdfCore.PdfCanvasAnchoredImageItem? item, out string? unsupportedReason, bool runningStory = false) {
        item = null;
        unsupportedReason = null;
        if (image._Image == null ||
            !TryResolveNativeAnchoredImageSource(image._Image, section, pageSize, margins, width, height,
                out WordDrawingLayoutInfo layout, out NativeAnchoredImagePlacement placement,
                out PdfCore.PdfImageTextExclusion? exclusion, out unsupportedReason)) return false;

        var content = new PdfCore.PdfPageCanvas();
        PdfCore.PdfImageStyle imageStyle = CreateNativeImageStyle(image, options, source);
        string? linkUri = GetNativeImageLink(image);
        string? linkContents = linkUri != null && !string.IsNullOrWhiteSpace(image.AltText) ? image.AltText : null;
        bool behindText = IsNativeBehindText(layout);
        // Wrapped text keeps clear of the picture; in a header or footer it stays in the story's layer, under the body.
        if (behindText || runningStory)
            content.Image(bytes, global::OfficeIMO.OfficeLength.Points(0D), global::OfficeIMO.OfficeLength.Points(0D), global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(height), style: imageStyle,
                    linkUri: linkUri, linkContents: linkContents);
        else
            content.ForegroundImage(bytes, 0D, 0D, width, height,
                zOrder: layout.RelativeHeight, style: imageStyle, linkUri: linkUri, linkContents: linkContents);
        IReadOnlyList<PdfCore.PdfCanvasItem> paint = content.Items;
        if (behindText)
            paint = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasBehindTextItem(paint, layout.RelativeHeight) };
        // The story's in-front layer orders it by z-order with the story's other in-front drawings.
        else if (runningStory)
            paint = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasInFrontOfTextItem(paint, layout.RelativeHeight) };
        item = new PdfCore.PdfCanvasAnchoredImageItem(sourceIdentity, new PdfCore.PdfCanvasBlock(paint),
            width, height, new PdfCore.PdfAnchoredImagePosition(placement.X, placement.Y,
                placement.MarginX, margins.Left.ToPoints(), placement.ParagraphY), exclusion);
        return true;
    }

    /// <summary>
    /// Uses one source eligibility decision for running-story admission and actual
    /// descriptor construction. It reads geometry only and never prepares media.
    /// </summary>
    private static bool TryResolveNativeAnchoredImageSource(W.Drawing drawing, WordSection? section,
        PdfCore.PdfPageSize pageSize, PdfCore.PdfPageMargins margins, double width, double height,
        out WordDrawingLayoutInfo layout, out NativeAnchoredImagePlacement placement,
        out PdfCore.PdfImageTextExclusion? exclusion, out string? unsupportedReason) {
        placement = default;
        exclusion = null;
        unsupportedReason = null;
        if (drawing.Anchor is not DW.Anchor anchor ||
            !WordDrawingLayoutReader.TryRead(drawing, out layout) ||
            !TryResolveNativeAnchoredImageGeometry(layout, section, pageSize, margins, width, height,
                out placement)) {
            layout = null!;
            unsupportedReason = "The picture anchor cannot be resolved in the current page and paragraph frame.";
            return false;
        }
        if (!TryReadNativeAnchoredImageExclusion(anchor, out exclusion, out unsupportedReason)) return false;
        if (exclusion == null) return true;
        var picture = anchor.GetFirstChild<A.Graphic>()?.GraphicData?
            .GetFirstChild<DocumentFormat.OpenXml.Drawing.Pictures.Picture>();
        if (picture == null || layout.IsGroup || drawing.Ancestors<W.TableCell>().Any() ||
            drawing.Ancestors<W.TextBoxContent>().Any() ||
            picture.ShapeProperties?.Transform2D?.Rotation?.Value is > 0 or < 0) {
            unsupportedReason = "Rotated, grouped, table-cell and text-box picture wrapping is outside the supported picture wrapping.";
            return false;
        }
        if (exclusion.Kind != PdfCore.PdfImageTextWrapKind.TopAndBottom &&
            exclusion.Side != PdfCore.PdfImageTextWrapSide.BothSides) {
            unsupportedReason = "The requested wrapping side is retained as source data but is not yet qualified for native picture layout.";
            return false;
        }
        // A wrap polygon already encloses the picture's effects; rectangular wrapping adds the effect extent.
        if (exclusion.Polygon == null) exclusion = AddNativeEffectExtentToExclusion(anchor, exclusion);
        return true;
    }

    /// <summary>
    /// Reads persisted rectangular wrapping and explicit anchor distances only.
    /// Child distances remain unsupported until their normative precedence is independently qualified;
    /// missing source values are not replaced with Word defaults or fitted layout gaps.
    /// </summary>
    private static bool TryReadNativeAnchoredImageExclusion(DW.Anchor anchor,
        out PdfCore.PdfImageTextExclusion? exclusion, out string? unsupportedReason) {
        exclusion = null;
        unsupportedReason = null;
        if (anchor.ChildElements.Count(child => child is DW.WrapNone or DW.WrapSquare or DW.WrapTopBottom or
                DW.WrapTight or DW.WrapThrough) != 1) {
            unsupportedReason = "The picture must declare exactly one wrapping primitive.";
            return false;
        }
        if (anchor.GetFirstChild<DW.WrapNone>() != null) return true;
        OpenXmlElement wrapping;
        PdfCore.PdfImageTextWrapKind kind;
        PdfCore.PdfImageTextWrapSide side;
        List<(double X, double Y)>? polygon = null;
        if (anchor.GetFirstChild<DW.WrapSquare>() is DW.WrapSquare square) {
            wrapping = square;
            kind = PdfCore.PdfImageTextWrapKind.Square;
            if (!TryReadNativeImageWrapSide(square, out side)) {
                unsupportedReason = "The square wrapping side is missing or invalid.";
                return false;
            }
        } else if (anchor.GetFirstChild<DW.WrapTopBottom>() is DW.WrapTopBottom topBottom) {
            wrapping = topBottom;
            kind = PdfCore.PdfImageTextWrapKind.TopAndBottom;
            side = PdfCore.PdfImageTextWrapSide.BothSides;
        } else {
            wrapping = (OpenXmlElement?)anchor.GetFirstChild<DW.WrapTight>() ?? anchor.GetFirstChild<DW.WrapThrough>()!;
            kind = wrapping is DW.WrapTight ? PdfCore.PdfImageTextWrapKind.Tight : PdfCore.PdfImageTextWrapKind.Through;
            if (!TryReadNativeImageWrapSide(wrapping, out side)) {
                unsupportedReason = "The wrapping side is missing or invalid.";
                return false;
            }
            if (!TryReadNativeWrapPolygon(wrapping.GetFirstChild<DW.WrapPolygon>(), out polygon)) {
                unsupportedReason = "The wrap polygon is missing or has fewer than three points.";
                return false;
            }
        }
        if (wrapping.GetAttributes().Any(attribute => attribute.NamespaceUri == string.Empty &&
                attribute.LocalName is "distL" or "distT" or "distR" or "distB")) {
            unsupportedReason = "Wrap-child text distances are not qualified; their precedence over anchor distances is not inferred.";
            return false;
        }
        if (!TryReadNativeImageWrapDistance(anchor, "distT", out double top) ||
            !TryReadNativeImageWrapDistance(anchor, "distB", out double bottom)) {
            unsupportedReason = "The anchor omits or has invalid top or bottom text distances.";
            return false;
        }
        double left = 0D, right = 0D;
        if (kind != PdfCore.PdfImageTextWrapKind.TopAndBottom &&
            (!TryReadNativeImageWrapDistance(anchor, "distL", out left) ||
             !TryReadNativeImageWrapDistance(anchor, "distR", out right))) {
            unsupportedReason = "The anchor omits or has invalid left or right text distances.";
            return false;
        }
        exclusion = new PdfCore.PdfImageTextExclusion(kind, side, left, top, right, bottom, polygon);
        return true;
    }

    private static bool TryReadNativeImageWrapSide(OpenXmlElement wrapping, out PdfCore.PdfImageTextWrapSide side) {
        switch (ReadNativeImageWrapAttribute(wrapping, "wrapText")) {
            case "bothSides": side = PdfCore.PdfImageTextWrapSide.BothSides; return true;
            case "left": side = PdfCore.PdfImageTextWrapSide.Left; return true;
            case "right": side = PdfCore.PdfImageTextWrapSide.Right; return true;
            case "largest": side = PdfCore.PdfImageTextWrapSide.Largest; return true;
            default: side = default; return false;
        }
    }

    /// <summary>
    /// Word stores wrap polygon points in a 21600 by 21600 space over the picture's extent, and the
    /// points may lie outside it when the polygon encloses the picture's effects.
    /// </summary>
    private static bool TryReadNativeWrapPolygon(DW.WrapPolygon? source, out List<(double X, double Y)>? polygon) {
        polygon = null;
        if (source == null) return false;
        var points = new List<(double X, double Y)>();
        foreach (OpenXmlElement point in source.ChildElements) {
            if (point is not DW.StartPoint and not DW.LineTo) continue;
            if (!long.TryParse(ReadNativeImageWrapAttribute(point, "x"), NumberStyles.Integer, CultureInfo.InvariantCulture, out long x) ||
                !long.TryParse(ReadNativeImageWrapAttribute(point, "y"), NumberStyles.Integer, CultureInfo.InvariantCulture, out long y)) return false;
            points.Add((x / 21600D, y / 21600D));
        }
        // The closing point repeats the start.
        if (points.Count > 1 && points[0] == points[points.Count - 1]) points.RemoveAt(points.Count - 1);
        if (points.Count < 3) return false;
        polygon = points;
        return true;
    }

    /// <summary>Word keeps text clear of a drawing's effect extent as well as its wrap distances.</summary>
    private static PdfCore.PdfImageTextExclusion AddNativeEffectExtentToExclusion(DW.Anchor anchor, PdfCore.PdfImageTextExclusion exclusion) {
        if (anchor.GetFirstChild<DW.EffectExtent>() is not { } extent) return exclusion;
        static double Points(Int64Value? value) => Math.Max(0L, value?.Value ?? 0L) / 12700D;
        return new PdfCore.PdfImageTextExclusion(exclusion.Kind, exclusion.Side,
            exclusion.DistanceLeft + Points(extent.LeftEdge), exclusion.DistanceTop + Points(extent.TopEdge),
            exclusion.DistanceRight + Points(extent.RightEdge), exclusion.DistanceBottom + Points(extent.BottomEdge),
            exclusion.Polygon);
    }

    /// <summary>
    /// Word paints a wrapped drawing by z-order whatever its behindDoc flag; only a drawing that
    /// does not wrap text goes behind the text.
    /// </summary>
    private static bool IsNativeBehindText(WordDrawingLayoutInfo layout) =>
        layout.BehindDocument && layout.Wrap == WordDrawingWrapKind.None;

    /// <summary>Reads optional source markup without turning an absent attribute into an exception.</summary>
    private static string? ReadNativeImageWrapAttribute(OpenXmlElement source, string name) =>
        source.GetAttributes().Where(attribute => attribute.NamespaceUri == string.Empty && attribute.LocalName == name)
            .Select(attribute => attribute.Value).FirstOrDefault();

    /// <summary>Converts one explicit unsigned EMU distance using the source format's units.</summary>
    private static bool TryReadNativeImageWrapDistance(OpenXmlElement source, string name, out double points) {
        points = 0D;
        if (!uint.TryParse(ReadNativeImageWrapAttribute(source, name), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out uint emus)) return false;
        points = emus / 12700D;
        return true;
    }

}
