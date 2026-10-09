using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    private static bool ProjectTextFitting(OdgShape shape, bool lineCaption, bool autoGrow, HashSet<string> losses) {
        OdfTextFitMode? mode;
        try { mode = shape.TextFitMode; }
        catch (NotSupportedException) { losses.Add("text-fitting"); return false; }
        if (mode is null or OdfTextFitMode.None) return false;
        if (mode == OdfTextFitMode.ShrinkToFit && IsFixedTextFrame(shape) && !lineCaption && !autoGrow) return true;
        losses.Add("text-fitting");
        return false;
    }

    private static bool IsFixedTextFrame(OdgShape shape) => shape.ElementName == "rect" ||
        shape.ElementName == "frame" && shape.TextRoot.Name == OdfNamespaces.Draw + "text-box";

    private static void AddFixedFrameText(OfficeDrawing target, IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        double width, double height, OfficeTextAreaAlignment areaAlignment, OfficeTextVerticalAlignment vertical,
        bool wrapText, OfficeTextPadding padding, bool shrinkToFit, OdfConversionReport report, string feature,
        CancellationToken cancellationToken, OfficeDrawingTextMetrics? layoutMetrics, bool reportHorizontalClipping = false,
        bool normalizeHorizontalPaint = false) {
        cancellationToken.ThrowIfCancellationRequested();
        double contentWidth = width - padding.Horizontal, contentHeight = height - padding.Vertical;
        if (contentWidth <= 0D || contentHeight <= 0D) {
            ReportLayout(clipped: true);
            return; // The shared text frame requires a positive padded content rectangle.
        }
        target.AddRichTextParagraphsCore(paragraphs, 0, 0, width, height, vertical, wrapText, padding, shrinkToFit, areaAlignment);
        var text = (OfficeDrawingRichText)target.Elements[target.Elements.Count - 1];
        text.NormalizeHorizontalPaint = normalizeHorizontalPaint;
        OfficeDrawingTextMetrics metrics = ResolveTextMetrics(target, layoutMetrics, cancellationToken);
        // The shared no-wrap layout permits ordinary body overhang while retaining
        // the saved content width for paragraph margins, tabs and layout limits.
        OfficeRichTextBlockLayout layout = OfficeDrawingTextLayout.CreateForProjectionWithMetrics(text, contentWidth, contentHeight, metrics);
        bool layoutClipped = layout.Clipped || OfficeDrawingTextLayout.IsVerticalPaintClipped(
            layout, contentHeight, vertical, metrics.MeasurePaintBounds, metrics.MeasureSegmentPaintBounds);
        if (reportHorizontalClipping) {
            var paint = OfficeDrawingTextLayout.HorizontalPaintEnvelope(layout, metrics.MeasureHorizontalPaintBounds);
            layoutClipped |= paint.Left < -.01D || paint.Right > contentWidth + .01D;
        }
        if (shrinkToFit) {
            // Intrinsic center/right placement can hide natural width in the reported area extent.
            // The shared left-area owner measures the same fitting floor without that anchor offset.
            OfficeRichTextBlockLayout natural = areaAlignment == OfficeTextAreaAlignment.Left ? layout :
                OfficeDrawingTextLayout.CreateWithMetrics(text.Clone().WithTextAreaAlignment(OfficeTextAreaAlignment.Left),
                    contentWidth, contentHeight, metrics);
            layoutClipped |= natural.Clipped || natural.Width > contentWidth + .01D;
        }
        ReportLayout(layoutClipped);

        void ReportLayout(bool clipped) {
            if (shrinkToFit) report.Add(feature + ":text-fitting", OdfConversionMappingStatus.Approximated,
                message: "Text shrinks within its saved frame using shared font layout. The largest font has a six-point fitting floor; absolute paragraph metrics remain fixed, and native metrics or later font-provider changes can alter fitting.");
            if (clipped) ReportTextClipped(report, feature);
        }
    }

    private static void ReportTextClipped(OdfConversionReport report, string feature) =>
        report.Add(feature + ":text-clipped", OdfConversionMappingStatus.Unsupported,
            message: "The shared text layout clips content or text paint, or horizontal growth leaves paint outside the projected padded frame with the current conversion font resources. Retained ODF text is not fully contained in that frame. Later font-provider changes can alter this decision.");
}
