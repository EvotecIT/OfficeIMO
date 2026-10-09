using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    private readonly struct LineTextFrame {
        internal LineTextFrame((double Left, double Top, double Right, double Bottom) bounds, OfficeTransform transform) {
            Bounds = bounds; Transform = transform;
        }
        internal (double Left, double Top, double Right, double Bottom) Bounds { get; }
        internal OfficeTransform Transform { get; }
    }

    private static void ProjectLineText(OdgShape shape, OfficeDrawing target, OdfConversionReport report,
        IReadOnlyList<OfficePathCommand> path, OfficeTransform transform, DrawingFieldContext fields,
        CancellationToken cancellationToken, OfficeDrawingTextMetrics? layoutMetrics) {
        try {
            int budget = OfficeTextLayoutEngine.MaximumLayoutTextCharacters;
            if (!HasDrawingText(shape.TextRoot, ref budget)) return;
            if (!OdgShape.IsTranslationTransform(transform)) {
                if (!IsLineCaptionRotation(transform) || path.Count != 2 || path[1].Kind != OfficePathCommandKind.LineTo ||
                    shape.IsConnector && (shape.StartShapeId != null || shape.EndShapeId != null))
                    throw new NotSupportedException("Line-label rotation supports one detached straight segment; attached, scaled, reflected, sheared or curved text frames require native placement.");
                // Native rotation materializes line coordinates without scaling fonts.
                // Connector captions remain horizontal around the rotated route bounds.
                path = new[] { OfficePathCommand.MoveTo(transform.TransformPoint(path[0].Point)),
                    OfficePathCommand.LineTo(transform.TransformPoint(path[1].Point)) };
                transform = OfficeTransform.Identity;
            }
            // Native connector captions use the curve's true bounds, not its control-point hull
            // or endpoint rectangle. Keep extrema calculation in the shared geometry owner.
            var bounds = OfficePathGeometry.Bounds(path);
            if (shape.ElementName == "line") {
                OfficePoint start = path[0].Point, end = path[path.Count - 1].Point;
                double dx = end.X - start.X, dy = end.Y - start.Y;
                double length = OfficeGeometry.Distance(start, end);
                bounds = (0, 0, length, 0);
                // Native draw:line text follows the directed line. Connector text stays horizontal.
                transform = OfficeTransform.RotateDegrees(Math.Atan2(dy, dx) * 180 / Math.PI)
                    .Then(OfficeTransform.Translate(start.X, start.Y)).Then(transform);
            }
            ProjectText(shape, target, report, 1, 1, fields, new LineTextFrame(bounds, transform), cancellationToken, layoutMetrics);
        } catch (Exception exception) when (exception is ArgumentException or InvalidDataException or NotSupportedException or OverflowException or FormatException) {
            report.Add("shape:" + shape.Name + ":text", OdfConversionMappingStatus.Skipped, message: exception.Message);
        }
    }

    private static bool IsLineCaptionRotation(OfficeTransform transform) {
        const double tolerance = 1e-12;
        return Math.Abs(transform.M11 * transform.M11 + transform.M12 * transform.M12 - 1) < tolerance &&
            Math.Abs(transform.M21 * transform.M21 + transform.M22 * transform.M22 - 1) < tolerance &&
            Math.Abs(transform.M11 * transform.M21 + transform.M12 * transform.M22) < tolerance &&
            Math.Abs(transform.M11 * transform.M22 - transform.M12 * transform.M21 - 1) < tolerance;
    }

    // Native line labels anchor to route bounds, including bends. Their unwrapped text may exceed
    // a zero-width/height route, so measure in the shared owner and expand the physical text frame.
    private static void AddLineLabel(OdgShape shape, OfficeDrawing target, IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        LineTextFrame placement, OfficeTextVerticalAlignment vertical,
        string? wrap, OfficeTextPadding padding, OdfConversionReport report, string feature,
        CancellationToken cancellationToken, OfficeDrawingTextMetrics? layoutMetrics) {
        if (wrap is not (null or "wrap" or "no-wrap"))
            throw new NotSupportedException("Unsupported native line-label wrapping option.");
        string? horizontal = shape.ReadGraphicProperty(OdfNamespaces.Draw + "textarea-horizontal-align");
        if (horizontal is not (null or "justify" or "left" or "center" or "right"))
            throw new NotSupportedException("Unsupported native line-label text-area horizontal alignment.");
        if (paragraphs.All(paragraph => paragraph.Runs.All(run => run.Text.Length == 0) &&
            (paragraph.Label == null || paragraph.Label.Run.Text.Length == 0)))
            throw new NotSupportedException("A line label without cached display text requires native field refresh.");

        // Intrinsic width must not include the alignment offset of a centered/right paragraph.
        var intrinsic = paragraphs.Select(IntrinsicLineLabelParagraph).ToArray();
        OfficeDrawingTextMetrics metrics = ResolveTextMetrics(target, layoutMetrics, cancellationToken);
        var measured = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(intrinsic, double.MaxValue, double.MaxValue,
            metrics, wrap: false, cancellationToken: cancellationToken);
        if (measured.Clipped)
            throw new NotSupportedException("Line-label text exceeds the shared paragraph layout limits.");
        var anchor = LineLabelAnchor(placement.Bounds, padding);
        double anchorWidth = anchor.Right - anchor.Left, anchorHeight = anchor.Bottom - anchor.Top;
        // Justified text areas span the route; the other alignments position an intrinsic
        // paragraph block. Paragraph alignment still applies independently inside that block.
        double innerWidth = Math.Max(horizontal is null or "justify" ? Math.Max(anchorWidth, measured.Width) : measured.Width, .001);
        double innerHeight = Math.Max(Math.Max(anchorHeight, measured.Height), .001);
        // Keep the declared padding on the shared frame. Native route insets determine its
        // placement before the frame grows to contain the caption's logical line boxes.
        double width = innerWidth + padding.Horizontal, height = innerHeight + padding.Vertical;
        double x = (horizontal switch {
            "left" => anchor.Left,
            "right" => anchor.Right - innerWidth,
            _ => anchor.Left + (anchorWidth - innerWidth) / 2
        }) - padding.Left;
        double y = vertical switch {
            OfficeTextVerticalAlignment.Top => anchor.Top - padding.Top,
            OfficeTextVerticalAlignment.Bottom => anchor.Bottom - innerHeight - padding.Top,
            _ => anchor.Top + (anchorHeight - innerHeight) / 2 - padding.Top
        };
        // Grow only the paint canvas for overflowing ink, retaining the logical text frame,
        // its alignment and native anchor. Use the actual paragraph alignment for overhang.
        var aligned = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(paragraphs, innerWidth, innerHeight,
            metrics, wrap: false, cancellationToken: cancellationToken);
        if (aligned.Clipped || OfficeDrawingTextLayout.IsVerticalPaintClipped(
            aligned, innerHeight, vertical, metrics.MeasurePaintBounds, metrics.MeasureSegmentPaintBounds))
            ReportTextClipped(report, feature);
        var horizontalPaint = OfficeDrawingTextLayout.HorizontalPaintEnvelope(aligned, metrics.MeasureHorizontalPaintBounds);
        double left = Math.Ceiling(Math.Max(0, -padding.Left - horizontalPaint.Left));
        double right = Math.Ceiling(Math.Max(0, padding.Left + horizontalPaint.Right - width));
        var painted = OfficeDrawingTextLayout.IncludePaintedHeight(measured, double.MaxValue, metrics.MeasurePaintBounds, metrics.MeasureSegmentPaintBounds);
        double top = painted.ContentOffsetY, extra = painted.Height - measured.Height;
        var label = new OfficeDrawing(width + left + right, height + extra);
        CopyDrawingResources(target, label);
        label.AddRichTextParagraphs(paragraphs, left, top, width, height, vertical, wrapText: false, padding: padding);
        OfficeTransform transform = OfficeTransform.Translate(x - left, y - top).Then(placement.Transform);
        if (OdgShape.IsTranslationTransform(transform) && transform.OffsetX >= 0 && transform.OffsetY >= 0 &&
            transform.OffsetX + label.Width <= target.Width && transform.OffsetY + label.Height <= target.Height)
            target.AddDrawing(label, transform.OffsetX, transform.OffsetY);
        else target.AddEffectDrawing(label, transform);
    }

    private static (double Left, double Top, double Right, double Bottom) LineLabelAnchor(
        (double Left, double Top, double Right, double Bottom) route, OfficeTextPadding padding) {
        double width = route.Right - route.Left, height = route.Bottom - route.Top;
        double left = padding.Left, right = padding.Right, top = padding.Top, bottom = padding.Bottom;
        // Positive route dimensions balance excessive side distances, stopping at zero.
        // Zero-width routes also bypass vertical balancing. Degenerate line frames instead
        // normalize their inset endpoints, which can put a top caption above the route.
        if (width > 0 && padding.Horizontal >= width) {
            double excess = (padding.Horizontal - width) / 2;
            left = Math.Max(0, left - excess); right = Math.Max(0, right - excess);
        }
        if (width > 0 && height > 0 && padding.Vertical >= height) {
            double excess = (padding.Vertical - height) / 2;
            top = Math.Max(0, top - excess); bottom = Math.Max(0, bottom - excess);
        }
        double x1 = route.Left + left, x2 = route.Right - right;
        double y1 = route.Top + top, y2 = route.Bottom - bottom;
        return (Math.Min(x1, x2), Math.Min(y1, y2), Math.Max(x1, x2), Math.Max(y1, y2));
    }

    private static OfficeRichTextParagraph IntrinsicLineLabelParagraph(OfficeRichTextParagraph paragraph) {
        var intrinsic = paragraph.Label == null ? new OfficeRichTextParagraph(paragraph.Runs,
            OfficeTextAlignment.Left, paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor) :
            new OfficeRichTextParagraph(paragraph.Runs, paragraph.Label, OfficeTextAlignment.Left,
                paragraph.LineHeight, paragraph.Margins, paragraph.Indent, paragraph.LineHeightFactor);
        return paragraph.TabStops == null ? intrinsic : intrinsic.WithTabStops(paragraph.TabStops);
    }
}
