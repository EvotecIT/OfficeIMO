using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    private static bool PrepareTextBoxLayout(OdgShape shape) => IsOrdinaryTextBox(shape) &&
        (HasInstanceTextBoxSize(shape) || shape.ReadGraphicProperty(OdfNamespaces.Draw + "auto-grow-height") == "true" ||
            shape.ReadGraphicProperty(OdfNamespaces.Draw + "auto-grow-width") == "true");

    private static bool ReadProjectionGrowth(OdgShape shape, string name, HashSet<string> losses) {
        string? value = shape.ReadGraphicProperty(OdfNamespaces.Draw + name);
        if (value is not (null or "false" or "true")) losses.Add("text-auto-size");
        return value == "true";
    }

    private static bool HasInstanceTextBoxSize(OdgShape shape) => shape.TextRoot.Name == OdfNamespaces.Draw + "text-box" &&
        new[] { "min-height", "min-width", "max-height", "max-width" }.Any(name =>
            shape.TextRoot.Attribute(OdfNamespaces.Fo + name) != null);

    private static OfficeDrawing GrowTextBox(OdgShape shape, OfficeDrawing target, IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        ref double width, ref double height, OfficeTextPadding padding, OfficeTextAreaAlignment areaAlignment,
        OfficeTextVerticalAlignment vertical, bool wrapText, bool growHeight, bool growWidth,
        HashSet<string> losses, OdfConversionReport report, string feature,
        CancellationToken cancellationToken, OfficeDrawingTextMetrics? layoutMetrics, TextBoxConstraints? constraints,
        out bool normalizeHorizontalPaint) {
        normalizeHorizontalPaint = false;
        if (!growHeight && !growWidth) return target;
        bool fitting;
        try { fitting = shape.TextFitMode is not (null or OdfTextFitMode.None); }
        catch (NotSupportedException) { fitting = true; }
        if (!IsOrdinaryTextBox(shape) ||
            shape.ReadGraphicProperty(OdfNamespaces.Draw + "auto-grow-width") is not ("false" or "true") ||
            shape.ReadGraphicProperty(OdfNamespaces.Draw + "auto-grow-height") is not ("false" or "true") ||
            vertical != OfficeTextVerticalAlignment.Top ||
            fitting || losses.Contains("writing-mode") || losses.Contains("wrap-option") || losses.Contains("text-area-alignment") ||
            losses.Contains("text-auto-size") || losses.Contains("text-size-constraints") ||
            losses.Contains("text-chain-flow") ||
            shape.Element.Attribute(OdfNamespaces.Style + "rel-width") != null ||
            shape.Element.Attribute(OdfNamespaces.Style + "rel-height") != null ||
            (string?)shape.Element.Attribute(OdfNamespaces.Text + "anchor-type") is not (null or "page")) {
            losses.Add("text-auto-size");
            return target;
        }
        cancellationToken.ThrowIfCancellationRequested();
        OfficeDrawingTextMetrics metrics = ResolveTextMetrics(target, layoutMetrics, cancellationToken);
        // ODF 1.4 section 20.98: horizontal major flow grows first. A maximum
        // width can leave wrapping; only then measure growth in the other axis.
        if (growWidth) {
            double requiredWidth = OfficeDrawingTextLayout.RequiredParagraphFrameWidth(paragraphs, metrics, cancellationToken);
            width = Math.Max(width, requiredWidth + padding.Horizontal);
            if (constraints?.MaximumWidth is double maximumWidth) width = Math.Min(width, maximumWidth);
        }
        double contentWidth = width - padding.Horizontal;
        if (growHeight && contentWidth > 0D) {
            OfficeRichTextBlockLayout measured;
            if (growWidth) {
                var measuringText = new OfficeDrawingRichText(Array.Empty<OfficeRichTextRun>(), 0, 0,
                    contentWidth, double.MaxValue, wrapText: wrapText).WithParagraphs(paragraphs).WithTextAreaAlignment(areaAlignment);
                measuringText.NormalizeHorizontalPaint = true;
                measured = OfficeDrawingTextLayout.CreateWithMetrics(measuringText, contentWidth, double.MaxValue, metrics);
            }
            else measured = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(paragraphs, contentWidth,
                double.MaxValue, metrics, wrapText, cancellationToken: cancellationToken);
            double requiredHeight = OfficeDrawingTextLayout.RequiredFrameHeight(measured, metrics.MeasurePaintBounds, metrics.MeasureSegmentPaintBounds);
            height = Math.Max(height, requiredHeight + padding.Vertical);
            if (constraints?.MaximumHeight is double maximumHeight) height = Math.Min(height, maximumHeight);
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (double.IsNaN(width) || double.IsInfinity(width) || double.IsNaN(height) || double.IsInfinity(height))
            throw new NotSupportedException("Text-box growth exceeds the finite drawing canvas.");
        target = ResizeTextCanvas(target, width, height);
        normalizeHorizontalPaint = growWidth;
        report.Add(feature + ":text-auto-size", OdfConversionMappingStatus.Approximated,
            message: "A horizontal top-aligned text box grows using shared paragraph, font and paint measurement. Width grows before height; wrapping is measured at the resolved width. " +
                "Instance minima replace saved dimensions; matching maxima cap the corresponding enabled growth. Without an instance minimum the saved dimension remains the floor. " +
                "Native metrics and later font-provider changes can alter growth. ODF geometry is unchanged.");
        return target;
    }
}
