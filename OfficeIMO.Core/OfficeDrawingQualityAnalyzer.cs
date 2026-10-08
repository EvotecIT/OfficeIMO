using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>
/// Performs dependency-free quality checks over shared drawing scenes before format-specific rendering.
/// </summary>
public static class OfficeDrawingQualityAnalyzer {
    private const int MaximumTextOverlapIssues = 4096;
    private const int MaximumTextOverlapComparisons = 1_000_000;
    /// <summary>
    /// Analyzes a drawing for reusable visual quality issues such as element overflow and text overlap.
    /// </summary>
    /// <param name="drawing">Drawing scene to analyze.</param>
    /// <param name="options">Optional quality-check tolerances.</param>
    /// <returns>Quality report with structured issues.</returns>
    public static OfficeDrawingQualityReport Analyze(OfficeDrawing drawing, OfficeDrawingQualityOptions? options = null) {
        if (drawing == null) {
            throw new ArgumentNullException(nameof(drawing));
        }

        return Analyze(drawing, drawing.Width, drawing.Height, options);
    }

    /// <summary>Analyzes rendered element rectangles against an explicit target canvas. This does not
    /// measure glyph ink, shadows, filter extents or content hidden by clipping. Coordinates use drawing units.</summary>
    public static OfficeDrawingQualityReport Analyze(OfficeDrawing drawing, double canvasWidth, double canvasHeight,
        OfficeDrawingQualityOptions? options = null, CancellationToken cancellationToken = default) {
        return AnalyzeAtOffset(drawing, 0D, 0D, canvasWidth, canvasHeight, options, cancellationToken);
    }

    internal static OfficeDrawingQualityReport AnalyzeAtOffset(OfficeDrawing drawing, double canvasLeft, double canvasTop,
        double canvasWidth, double canvasHeight, OfficeDrawingQualityOptions? options, CancellationToken cancellationToken) {
        if (drawing == null) throw new ArgumentNullException(nameof(drawing));
        if (double.IsNaN(canvasWidth) || double.IsInfinity(canvasWidth) || canvasWidth <= 0D) throw new ArgumentOutOfRangeException(nameof(canvasWidth));
        if (double.IsNaN(canvasHeight) || double.IsInfinity(canvasHeight) || canvasHeight <= 0D) throw new ArgumentOutOfRangeException(nameof(canvasHeight));
        cancellationToken.ThrowIfCancellationRequested();
        options ??= OfficeDrawingQualityOptions.Default;
        var issues = new List<OfficeDrawingQualityIssue>();
        var textBoxes = new List<(int Index, string Text, DrawingBounds Bounds)>();

        IReadOnlyList<OfficeDrawingElement> elements = drawing.Elements;
        for (int i = 0; i < elements.Count; i++) {
            AppendElementQuality(elements[i], i, OfficeTransform.Identity, canvasLeft, canvasTop, canvasWidth, canvasHeight, options, issues, textBoxes, cancellationToken);
        }

        if (options.DetectTextOverlap) {
            AddTextOverlapIssues(textBoxes, options.OverlapTolerance, issues, cancellationToken);
        }

        return new OfficeDrawingQualityReport(issues);
    }

    private static void AppendElementQuality(OfficeDrawingElement element, int rootIndex, OfficeTransform transform,
        double canvasLeft, double canvasTop, double canvasWidth, double canvasHeight, OfficeDrawingQualityOptions options, List<OfficeDrawingQualityIssue> issues,
        List<(int Index, string Text, DrawingBounds Bounds)> textBoxes, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (element is OfficeDrawingEffectGroup group) {
            // The intermediate canvas is storage, not painted geometry. Only its children
            // contribute bounds; compose their transforms before testing the target canvas.
            OfficeTransform childTransform = group.Transform.Then(transform);
            foreach (OfficeDrawingElement child in group.InnerDrawing.Elements)
                AppendElementQuality(child, rootIndex, childTransform, canvasLeft, canvasTop, canvasWidth, canvasHeight, options, issues, textBoxes, token);
            return;
        }
        DrawingBounds local = GetBounds(element);
        var transformed = transform.TransformRectangleBounds(local.Left, local.Top, local.Right - local.Left, local.Bottom - local.Top);
        var bounds = new DrawingBounds(transformed.Left, transformed.Top, transformed.Right, transformed.Bottom);
        var relative = new DrawingBounds(bounds.Left - canvasLeft, bounds.Top - canvasTop, bounds.Right - canvasLeft, bounds.Bottom - canvasTop);
        if (IsOutsideCanvas(relative, canvasWidth, canvasHeight, options.BoundsTolerance)) {
            issues.Add(new OfficeDrawingQualityIssue(OfficeDrawingQualityIssueKind.ElementOutsideBounds,
                FormatBoundsMessage(relative, canvasWidth, canvasHeight), rootIndex));
        }
        if (element is OfficeDrawingText text) textBoxes.Add((rootIndex, text.Text, bounds));
        else if (element is OfficeDrawingRichText richText) textBoxes.Add((rootIndex, richText.PlainText, bounds));
    }

    private static void AddTextOverlapIssues(IReadOnlyList<(int Index, string Text, DrawingBounds Bounds)> textBoxes, double tolerance, List<OfficeDrawingQualityIssue> issues, CancellationToken token) {
        var ordered = textBoxes.OrderBy(item => item.Bounds.Left).ToArray();
        int comparisons = 0;
        int overlapIssues = 0;
        for (int i = 0; i < ordered.Length; i++) {
            for (int j = i + 1; j < ordered.Length &&
                ordered[j].Bounds.Left < ordered[i].Bounds.Right - tolerance; j++) {
                token.ThrowIfCancellationRequested();
                if (++comparisons > MaximumTextOverlapComparisons)
                    throw new NotSupportedException("Drawing text overlap analysis exceeds its work limit.");
                if (!Overlaps(ordered[i].Bounds, ordered[j].Bounds, tolerance)) {
                    continue;
                }

                issues.Add(new OfficeDrawingQualityIssue(
                    OfficeDrawingQualityIssueKind.TextOverlap,
                    "Text box '" + Shorten(ordered[i].Text) + "' overlaps text box '" + Shorten(ordered[j].Text) + "'.",
                    ordered[i].Index,
                    ordered[j].Index));
                if (++overlapIssues > MaximumTextOverlapIssues)
                    throw new NotSupportedException("Drawing text overlap analysis exceeds its diagnostic limit.");
            }
        }
    }

    private static DrawingBounds GetBounds(OfficeDrawingElement element) {
        if (element is OfficeDrawingText text) {
            return GetRotatedBounds(
                text.X,
                text.Y,
                text.Width,
                text.Height,
                text.RotationDegrees,
                text.RotationCenterX,
                text.RotationCenterY);
        }

        if (element is OfficeDrawingRichText richText) {
            return GetRotatedBounds(
                richText.X,
                richText.Y,
                richText.Width,
                richText.Height,
                richText.RotationDegrees,
                richText.RotationCenterX,
                richText.RotationCenterY);
        }

        if (element is OfficeDrawingShape shape) {
            return new DrawingBounds(shape.X, shape.Y, shape.X + shape.Shape.Width, shape.Y + shape.Shape.Height);
        }

        if (element is OfficeDrawingImage image) {
            (double left, double top, double right, double bottom) = image.Projection.GetDestinationBounds();
            return new DrawingBounds(left, top, right, bottom);
        }

        if (element is OfficeDrawingImagePattern imagePattern) {
            OfficeImagePlacement area = imagePattern.Layout.Area;
            return new DrawingBounds(area.X, area.Y, area.X + area.Width, area.Y + area.Height);
        }

        if (element is OfficeDrawingTilingPattern tilingPattern) {
            OfficeImagePlacement area = tilingPattern.Area;
            return new DrawingBounds(area.X, area.Y, area.X + area.Width, area.Y + area.Height);
        }

        if (element is OfficeDrawingGroup group) {
            if (group.FrameTransform.HasValue && group.FrameTransform.Value.HasTransform) {
                OfficeTransform transform = group.FrameTransform.Value.CreateDestinationTransform();
                (double left, double top, double right, double bottom) = transform.TransformRectangleBounds(group.X, group.Y, group.ClipPath.Width, group.ClipPath.Height);
                return new DrawingBounds(left, top, right, bottom);
            }

            return new DrawingBounds(group.X, group.Y, group.X + group.ClipPath.Width, group.Y + group.ClipPath.Height);
        }

        if (element is OfficeDrawingLink link) {
            return new DrawingBounds(link.X, link.Y, link.X + link.Width, link.Y + link.Height);
        }

        return new DrawingBounds(0D, 0D, 0D, 0D);
    }

    private static DrawingBounds GetRotatedBounds(double x, double y, double width, double height, double rotationDegrees, double centerX, double centerY) {
        (double left, double top, double right, double bottom) = OfficeGeometry.GetRotatedRectangleBounds(
            x,
            y,
            width,
            height,
            rotationDegrees,
            centerX,
            centerY);
        return new DrawingBounds(left, top, right, bottom);
    }

    private static bool IsOutsideCanvas(DrawingBounds bounds, double width, double height, double tolerance) {
        return bounds.Left < -tolerance
               || bounds.Top < -tolerance
               || bounds.Right > width + tolerance
               || bounds.Bottom > height + tolerance;
    }

    private static bool Overlaps(DrawingBounds left, DrawingBounds right, double tolerance) {
        double overlapWidth = Math.Min(left.Right, right.Right) - Math.Max(left.Left, right.Left);
        double overlapHeight = Math.Min(left.Bottom, right.Bottom) - Math.Max(left.Top, right.Top);
        return overlapWidth > tolerance && overlapHeight > tolerance;
    }

    private static string FormatBoundsMessage(DrawingBounds bounds, double width, double height) {
        return string.Format(
            CultureInfo.InvariantCulture,
            "Element bounds [{0:0.###},{1:0.###},{2:0.###},{3:0.###}] exceed drawing bounds [0,0,{4:0.###},{5:0.###}].",
            bounds.Left,
            bounds.Top,
            bounds.Right,
            bounds.Bottom,
            width,
            height);
    }

    private static string Shorten(string value) {
        if (value.Length <= 32) {
            return value;
        }

        return value.Substring(0, 29) + "...";
    }

    private readonly struct DrawingBounds {
        public DrawingBounds(double left, double top, double right, double bottom) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        public double Left { get; }

        public double Top { get; }

        public double Right { get; }

        public double Bottom { get; }
    }
}
