using System.Collections.Generic;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private readonly record struct NativeAnchoredImagePlacement(double X, double Y, bool MarginX, bool ParagraphY);

    /// <summary>Shares supported picture-anchor geometry between running-story admission and emission.</summary>
    private static bool TryResolveNativeAnchoredImagePlacement(WordDrawingLayoutInfo layout,
        WordSection? section, PdfCore.PdfPageSize pageSize, PdfCore.PdfPageMargins margins, double width, double height,
        out NativeAnchoredImagePlacement placement) {
        placement = default;
        if (layout.Wrap != WordDrawingWrapKind.None) return false;
        return TryResolveNativeAnchoredImageGeometry(layout, section, pageSize, margins, width, height,
            out placement);
    }

    /// <summary>Resolves source anchor geometry independently of the wrapping admission policy.</summary>
    private static bool TryResolveNativeAnchoredImageGeometry(WordDrawingLayoutInfo layout,
        WordSection? section, PdfCore.PdfPageSize pageSize, PdfCore.PdfPageMargins margins, double width, double height,
        out NativeAnchoredImagePlacement placement) {
        placement = default;
        if (layout.Placement != WordDrawingPlacementKind.Anchored ||
            width <= 0D || height <= 0D || double.IsNaN(width) || double.IsInfinity(width) ||
            double.IsNaN(height) || double.IsInfinity(height)) return false;
        bool marginX = layout.HorizontalRelativeFrom is "margin" or "column";
        bool marginY = layout.VerticalRelativeFrom == "margin";
        bool paragraphY = layout.VerticalRelativeFrom == "paragraph";
        if ((marginX || marginY) && section == null) return false;
        // A flowing multi-column position needs the actual column chosen by pagination.
        if (layout.HorizontalRelativeFrom == "column" && section?._sectionProperties
                .GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Columns>() is { } columns &&
            (columns.ColumnCount?.Value > 1 || columns.ChildElements.Count > 1)) return false;
        string? horizontal = marginX ? "margin" : layout.HorizontalRelativeFrom;
        string? vertical = paragraphY ? "paragraph" : layout.VerticalRelativeFrom;
        if (!TryGetNativeGroupAxis(horizontal, layout.HorizontalOffsetPoints, layout.HorizontalAlignment,
                pageSize.Width, margins.Left, margins.Right, width, false, out double x) ||
            !TryGetNativeGroupAxis(vertical, layout.VerticalOffsetPoints, layout.VerticalAlignment,
                pageSize.Height, margins.Top, margins.Bottom, height, true, out double y)) return false;
        placement = new NativeAnchoredImagePlacement(x, y, marginX, paragraphY);
        return true;
    }

    /// <summary>Retains picture anchors and supported rectangular wrapping on their source paragraph.</summary>
    private static bool TryAddNativeAnchoredImage(INativePdfFlow pdf, WordImage image, byte[] bytes,
        double width, double height, PdfCore.PdfParagraphStyle style, PdfCore.PdfPageCanvas? target,
        WordParagraph? paragraph, WordPdfExportOptions? options, string source, WordSection? anchorSection = null) {
        if (image._Image?.Anchor == null ||
            !WordDrawingLayoutReader.TryRead(image._Image, out WordDrawingLayoutInfo layout))
            return false;

        WordSection? section = paragraph?.GetBodySection() ?? anchorSection;
        PdfCore.PdfPageMargins margins = section == null ? PdfCore.PdfPageMargins.Uniform(0D)
            : layout.Wrap == WordDrawingWrapKind.None ? GetNativeMargins(section, options)
            : GetNativeMargins(section, options, (0D, 0D));
        if (layout.Wrap != WordDrawingWrapKind.None) {
            // Visible run views retain their original source children, so the
            // drawing identity is stable across deferred story materializations.
            if (!TryPrepareNativeAnchoredImageItem(image._Image, image, bytes, width, height, section,
                    pdf.PageSize, margins, options, source, out PdfCore.PdfCanvasAnchoredImageItem? wrapped,
                    out string? unsupportedReason)) {
                RecordNativeImageWrapApproximation(options, image._Image, source,
                    "The wrapped picture was placed in document flow. " + unsupportedReason);
                return false;
            }
            PdfCore.PdfPageCanvas wrappedCanvas = target ?? new PdfCore.PdfPageCanvas();
            if (target == null && style.AnchoredCanvas != null) wrappedCanvas.AddItems(style.AnchoredCanvas.Items);
            wrappedCanvas.AddItems(new PdfCore.PdfCanvasItem[] { wrapped! });
            if (target == null) style.AnchoredCanvas = new PdfCore.PdfCanvasBlock(wrappedCanvas.Items);
            PdfCore.PdfAnchoredImagePosition position = wrapped!.Position;
            RecordNativeWrappedImagePageClipping(options, image._Image, source, pdf.PageSize,
                width, height, position);
            return true;
        }

        if (!TryResolveNativeAnchoredImagePlacement(layout, section, pdf.PageSize, margins, width, height,
                out NativeAnchoredImagePlacement placement)) return false;

        var imageCanvas = new PdfCore.PdfPageCanvas();
        PdfCore.PdfImageStyle imageStyle = CreateNativeImageStyle(image, options, source);
        // Effects retain signed producer offsets without weakening the public canvas coordinate contract.
        imageCanvas.Effect(OfficeTransform.Translate(placement.X, placement.Y), 1D, content => {
            if (layout.BehindDocument)
                content.Image(bytes, 0D, 0D, width, height, style: imageStyle);
            else
                content.ForegroundImage(bytes, 0D, 0D, width, height,
                    zOrder: layout.RelativeHeight, style: imageStyle);
        });
        IReadOnlyList<PdfCore.PdfCanvasItem> items = imageCanvas.Items;
        if (placement.MarginX) items = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasMarginAnchorItem(items, margins.Left) };
        if (placement.ParagraphY) items = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasParagraphAnchorItem(items) };
        if (layout.BehindDocument) items = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasBehindTextItem(items, layout.RelativeHeight) };
        PdfCore.PdfPageCanvas canvas = target ?? new PdfCore.PdfPageCanvas();
        if (target == null && style.AnchoredCanvas != null) canvas.AddItems(style.AnchoredCanvas.Items);
        canvas.AddItems(items);
        if (target == null) style.AnchoredCanvas = new PdfCore.PdfCanvasBlock(canvas.Items);
        if (options != null && (placement.X < 0D || placement.X + width > pdf.PageSize.Width + 0.001D ||
                !placement.ParagraphY && (placement.Y < 0D || placement.Y + height > pdf.PageSize.Height + 0.001D)))
            AddNativeExportWarning(options, "native-anchored-image-clipped", source,
                "The anchored image extends beyond the page; the page boundary clips its visible appearance.");
        return true;
    }
}
