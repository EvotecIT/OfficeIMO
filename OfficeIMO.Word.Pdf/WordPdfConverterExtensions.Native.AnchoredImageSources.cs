using System.Collections.Generic;
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
        string source, out PdfCore.PdfCanvasAnchoredImageItem? item, out string? unsupportedReason) {
        item = null;
        unsupportedReason = null;
        if (image._Image == null ||
            !TryResolveNativeAnchoredImageSource(image._Image, section, pageSize, margins, width, height,
                out WordDrawingLayoutInfo layout, out NativeAnchoredImagePlacement placement,
                out PdfCore.PdfImageTextExclusion? exclusion, out unsupportedReason)) return false;

        var content = new PdfCore.PdfPageCanvas();
        PdfCore.PdfImageStyle imageStyle = CreateNativeImageStyle(image, options, source);
        if (layout.BehindDocument)
            content.Image(bytes, global::OfficeIMO.OfficeLength.Points(0D), global::OfficeIMO.OfficeLength.Points(0D), global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(height), style: imageStyle);
        else
            content.ForegroundImage(bytes, 0D, 0D, width, height,
                zOrder: layout.RelativeHeight, style: imageStyle);
        IReadOnlyList<PdfCore.PdfCanvasItem> paint = content.Items;
        if (layout.BehindDocument)
            paint = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasBehindTextItem(paint, layout.RelativeHeight) };
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
            picture.ShapeProperties?.Transform2D?.Rotation?.Value is > 0 or < 0 ||
            !HasZeroNativeImageWrapEffectExtents(anchor)) {
            unsupportedReason = "Rotated, grouped, table-cell, text-box or effect-expanded picture wrapping is outside the rectangular source profile.";
            return false;
        }
        if (exclusion.Kind == PdfCore.PdfImageTextWrapKind.Square &&
            exclusion.Side != PdfCore.PdfImageTextWrapSide.BothSides) {
            unsupportedReason = "The requested square wrapping side is retained as source data but is not yet qualified for native picture layout.";
            return false;
        }
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
        if (anchor.GetFirstChild<DW.WrapSquare>() is DW.WrapSquare square) {
            wrapping = square;
            kind = PdfCore.PdfImageTextWrapKind.Square;
            switch (ReadNativeImageWrapAttribute(square, "wrapText")) {
                case "bothSides": side = PdfCore.PdfImageTextWrapSide.BothSides; break;
                case "left": side = PdfCore.PdfImageTextWrapSide.Left; break;
                case "right": side = PdfCore.PdfImageTextWrapSide.Right; break;
                case "largest": side = PdfCore.PdfImageTextWrapSide.Largest; break;
                default:
                    unsupportedReason = "The square wrapping side is missing or invalid.";
                    return false;
            }
        } else if (anchor.GetFirstChild<DW.WrapTopBottom>() is DW.WrapTopBottom topBottom) {
            wrapping = topBottom;
            kind = PdfCore.PdfImageTextWrapKind.TopAndBottom;
            side = PdfCore.PdfImageTextWrapSide.BothSides;
        } else {
            unsupportedReason = "The picture wrapping primitive is outside the rectangular source profile.";
            return false;
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
        if (kind == PdfCore.PdfImageTextWrapKind.Square &&
            (!TryReadNativeImageWrapDistance(anchor, "distL", out left) ||
             !TryReadNativeImageWrapDistance(anchor, "distR", out right))) {
            unsupportedReason = "The anchor omits or has invalid left or right text distances.";
            return false;
        }
        exclusion = new PdfCore.PdfImageTextExclusion(kind, side, left, top, right, bottom);
        return true;
    }

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

    /// <summary>Rejects unqualified expanded wrap boundaries without choosing child/parent precedence.</summary>
    private static bool HasZeroNativeImageWrapEffectExtents(DW.Anchor anchor) {
        foreach (DW.EffectExtent extent in anchor.Descendants<DW.EffectExtent>())
            foreach (string name in new[] { "l", "t", "r", "b" })
                if (!long.TryParse(ReadNativeImageWrapAttribute(extent, name), NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out long emus) || emus != 0L) return false;
        return true;
    }
}
