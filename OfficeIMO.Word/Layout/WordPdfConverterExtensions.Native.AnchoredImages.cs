using System.Collections.Generic;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private readonly record struct NativeAnchoredImagePlacement(double X, double Y, bool MarginX, bool ParagraphY);

    /// <summary>Shares supported picture-anchor geometry between running-story admission and emission.</summary>
    private static bool TryResolveNativeAnchoredImagePlacement(WordDrawingLayoutSnapshot layout,
        WordSection? section, PdfCore.PageSize pageSize, PdfCore.PageMargins margins, double width, double height,
        out NativeAnchoredImagePlacement placement) {
        placement = default;
        if (layout.Placement != WordDrawingPlacementKind.Anchored || layout.Wrap != WordDrawingWrapKind.None ||
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

    /// <summary>Retains non-wrapping image anchors on the page that contains their paragraph.</summary>
    private static bool TryAddNativeAnchoredImage(INativePdfFlow pdf, WordImage image, byte[] bytes,
        double width, double height, PdfCore.PdfParagraphStyle style, PdfCore.PdfPageCanvas? target,
        WordParagraph? paragraph, WordToPdfOptions? options, string source, WordSection? anchorSection = null) {
        if (image._Image?.Anchor == null ||
            !WordDrawingLayoutReader.TryRead(image._Image, out WordDrawingLayoutSnapshot layout))
            return false;

        WordSection? section = paragraph?.GetBodySection() ?? anchorSection;
        PdfCore.PageMargins margins = section == null ? PdfCore.PageMargins.Uniform(0D) : GetNativeMargins(section, options);
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
            AddNativeExportWarning(options, "NativeAnchoredImageClipped", source,
                "The anchored image extends beyond the page; the page boundary clips its visible appearance.");
        return true;
    }
}
