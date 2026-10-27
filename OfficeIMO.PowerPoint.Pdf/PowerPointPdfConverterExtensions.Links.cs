using System;
using System.Linq;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static string? ResolveSlideLinkDestination(Uri? link, int slideNumber,
        PowerPointPdfExportOptions options, string owner) {
        if (link == null || link.IsAbsoluteUri) return null;
        if (!PptCore.PowerPointHyperlinkResolver.TryParseSlideFragment(link, out int targetSlide)) {
            AddWarning(options, slideNumber, "relative-hyperlink",
                "Skipped a relative PowerPoint " + owner + " hyperlink that is not a supported slide target.",
                lossKind: OfficeConversionLossKind.Omission);
            return null;
        }

        string destination = PptCore.PowerPointHyperlinkResolver.GetSlideDestinationName(targetSlide);
        if (options.SlideLinkDestinations.Contains(destination)) return destination;
        AddWarning(options, slideNumber, "internal-slide-target-omitted",
            "Skipped a PowerPoint " + owner + " hyperlink because its target slide is absent from the PDF output.",
            lossKind: OfficeConversionLossKind.Omission);
        return null;
    }

    private static void RemoveUnavailableDrawingSlideLinks(OfficeDrawing drawing, PowerPointPdfExportOptions options,
        int? diagnosticSlideNumber = null) {
        foreach (OfficeDrawingRichText text in drawing.Elements.OfType<OfficeDrawingRichText>()) {
            foreach (OfficeRichTextRun run in text.Runs.Concat(text.Paragraphs.SelectMany(paragraph => paragraph.Runs))) {
                if (run.LinkDestinationName == null || options.SlideLinkDestinations.Contains(run.LinkDestinationName)) continue;
                if (diagnosticSlideNumber.HasValue)
                    AddWarning(options, diagnosticSlideNumber.Value, "internal-slide-target-omitted",
                        "Skipped a PowerPoint text run hyperlink because its target slide is absent from the PDF output.",
                        lossKind: OfficeConversionLossKind.Omission);
                run.LinkDestinationName = null;
            }
        }
        foreach (OfficeDrawingGroup group in drawing.Elements.OfType<OfficeDrawingGroup>())
            RemoveUnavailableDrawingSlideLinks(group.InnerDrawing, options, diagnosticSlideNumber);
        foreach (OfficeDrawingEffectGroup group in drawing.Elements.OfType<OfficeDrawingEffectGroup>())
            RemoveUnavailableDrawingSlideLinks(group.InnerDrawing, options, diagnosticSlideNumber);
    }

    private static void RenderShapeSlideLink(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointShape shape,
        double x, double y, double width, double height, int slideNumber, PowerPointPdfExportOptions options) {
        string? destination = ResolveSlideLinkDestination(shape.Hyperlink, slideNumber, options, "shape");
        if (destination != null) RenderSlideLinkFrame(canvas, shape, destination, x, y, width, height);
    }

    private static void RenderSlideLinkFrame(PdfCore.PdfPageCanvas canvas, PptCore.PowerPointShape shape,
        string destination, double x, double y, double width, double height) {
        double centerX = x + width / 2D, centerY = y + height / 2D;
        if (IsLineShape(shape)) {
            // Horizontal/vertical lines have zero extent, and AddLine's one-EMU
            // extent rounds to zero in PDF annotations. Include the visible stroke
            // while retaining the original rotation/flip center.
            double stroke = Math.Max(0.002D, shape.OutlineWidthPoints
                ?? (ParseOfficeColor(shape.OutlineColor).HasValue ? 1D : 0D));
            x -= stroke / 2D; y -= stroke / 2D;
            width += stroke; height += stroke;
        }
        if (width <= 0 || height <= 0) return;
        string? contents = string.IsNullOrWhiteSpace(shape.AltText) ? null : shape.AltText;
        void Add(PdfCore.PdfPageCanvas target) => target.LinkToNamedDestination(destination, x, y, width, height, contents);
        var frame = new OfficeImageFrameTransform(shape.Rotation ?? 0D, centerX, centerY,
            shape.HorizontalFlip == true, shape.VerticalFlip == true);
        if (frame.HasTransform) canvas.Effect(frame.CreateDestinationTransform(), 1D, Add);
        else Add(canvas);
    }
}
