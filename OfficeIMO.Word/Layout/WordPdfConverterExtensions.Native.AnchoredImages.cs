using System.Collections.Generic;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Retains non-wrapping image anchors on the page that contains their paragraph.</summary>
    private static bool TryAddNativeAnchoredImage(INativePdfFlow pdf, WordImage image, byte[] bytes,
        double width, double height, PdfCore.PdfParagraphStyle style, PdfCore.PdfPageCanvas? target,
        WordParagraph? paragraph, WordToPdfOptions? options, string source) {
        if (image._Image?.Anchor == null ||
            !WordDrawingLayoutReader.TryRead(image._Image, out WordDrawingLayoutSnapshot layout) ||
            layout.Wrap != WordDrawingWrapKind.None || layout.UsesSimplePosition ||
            (image.CropTop ?? 0) != 0 || (image.CropBottom ?? 0) != 0 ||
            (image.CropLeft ?? 0) != 0 || (image.CropRight ?? 0) != 0 || (image.Rotation ?? 0) != 0)
            return false;

        WordSection? section = paragraph?.GetBodySection();
        bool marginX = layout.HorizontalRelativeFrom is "margin" or "column";
        bool marginY = layout.VerticalRelativeFrom == "margin";
        bool paragraphY = layout.VerticalRelativeFrom == "paragraph";
        if ((marginX || marginY) && section == null) return false;
        // A flowing multi-column position needs the actual column chosen by pagination.
        if (layout.HorizontalRelativeFrom == "column" && section?._sectionProperties
                .GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Columns>() is { } columns &&
            (columns.ColumnCount?.Value > 1 || columns.ChildElements.Count > 1)) return false;
        PdfCore.PageMargins margins = section == null ? PdfCore.PageMargins.Uniform(0D) : GetNativeMargins(section, options);
        string? horizontal = marginX ? "margin" : layout.HorizontalRelativeFrom;
        string? vertical = paragraphY ? "paragraph" : layout.VerticalRelativeFrom;
        if (!TryGetNativeGroupAxis(horizontal, layout.HorizontalOffsetPoints, layout.HorizontalAlignment,
                pdf.PageSize.Width, margins.Left, margins.Right, width, false, out double x) ||
            !TryGetNativeGroupAxis(vertical, layout.VerticalOffsetPoints, layout.VerticalAlignment,
                pdf.PageSize.Height, margins.Top, margins.Bottom, height, true, out double y)) return false;

        var imageCanvas = new PdfCore.PdfPageCanvas();
        // Effects retain signed producer offsets without weakening the public canvas coordinate contract.
        imageCanvas.Effect(OfficeTransform.Translate(x, y), 1D, content => {
            if (layout.BehindDocument)
                content.Image(bytes, 0D, 0D, width, height,
                    horizontalFlip: image.HorizontalFlip ?? false, verticalFlip: image.VerticalFlip ?? false);
            else
                content.ForegroundImage(bytes, 0D, 0D, width, height,
                    horizontalFlip: image.HorizontalFlip ?? false, verticalFlip: image.VerticalFlip ?? false,
                    zOrder: layout.RelativeHeight);
        });
        IReadOnlyList<PdfCore.PdfCanvasItem> items = imageCanvas.Items;
        if (marginX) items = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasMarginAnchorItem(items, margins.Left) };
        if (paragraphY) items = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasParagraphAnchorItem(items) };
        if (layout.BehindDocument) items = new PdfCore.PdfCanvasItem[] { new PdfCore.PdfCanvasBehindTextItem(items, layout.RelativeHeight) };
        PdfCore.PdfPageCanvas canvas = target ?? new PdfCore.PdfPageCanvas();
        if (target == null && style.AnchoredCanvas != null) canvas.AddItems(style.AnchoredCanvas.Items);
        canvas.AddItems(items);
        if (target == null) style.AnchoredCanvas = new PdfCore.PdfCanvasBlock(canvas.Items);
        if (options != null && (x < 0D || x + width > pdf.PageSize.Width + 0.001D ||
                !paragraphY && (y < 0D || y + height > pdf.PageSize.Height + 0.001D)))
            AddNativeExportWarning(options, "NativeAnchoredImageClipped", source,
                "The anchored image extends beyond the page; the page boundary clips its visible appearance.");
        return true;
    }
}
