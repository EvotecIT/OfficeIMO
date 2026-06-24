using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Drawing;

/// <summary>
/// Performs dependency-free quality checks over shared drawing scenes before format-specific rendering.
/// </summary>
public static class OfficeDrawingQualityAnalyzer {
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

        options ??= OfficeDrawingQualityOptions.Default;
        var issues = new List<OfficeDrawingQualityIssue>();
        var textBoxes = new List<(int Index, OfficeDrawingText Text, DrawingBounds Bounds)>();

        IReadOnlyList<OfficeDrawingElement> elements = drawing.Elements;
        for (int i = 0; i < elements.Count; i++) {
            OfficeDrawingElement element = elements[i];
            DrawingBounds bounds = GetBounds(element);
            if (IsOutsideCanvas(bounds, drawing.Width, drawing.Height, options.BoundsTolerance)) {
                issues.Add(new OfficeDrawingQualityIssue(
                    OfficeDrawingQualityIssueKind.ElementOutsideBounds,
                    FormatBoundsMessage(bounds, drawing.Width, drawing.Height),
                    i));
            }

            if (element is OfficeDrawingText text) {
                textBoxes.Add((i, text, bounds));
            }
        }

        if (options.DetectTextOverlap) {
            AddTextOverlapIssues(textBoxes, options.OverlapTolerance, issues);
        }

        return new OfficeDrawingQualityReport(issues);
    }

    private static void AddTextOverlapIssues(IReadOnlyList<(int Index, OfficeDrawingText Text, DrawingBounds Bounds)> textBoxes, double tolerance, List<OfficeDrawingQualityIssue> issues) {
        for (int i = 0; i < textBoxes.Count; i++) {
            for (int j = i + 1; j < textBoxes.Count; j++) {
                if (!Overlaps(textBoxes[i].Bounds, textBoxes[j].Bounds, tolerance)) {
                    continue;
                }

                issues.Add(new OfficeDrawingQualityIssue(
                    OfficeDrawingQualityIssueKind.TextOverlap,
                    "Text box '" + Shorten(textBoxes[i].Text.Text) + "' overlaps text box '" + Shorten(textBoxes[j].Text.Text) + "'.",
                    textBoxes[i].Index,
                    textBoxes[j].Index));
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

        if (element is OfficeDrawingShape shape) {
            return new DrawingBounds(shape.X, shape.Y, shape.X + shape.Shape.Width, shape.Y + shape.Shape.Height);
        }

        return new DrawingBounds(0D, 0D, 0D, 0D);
    }

    private static DrawingBounds GetRotatedBounds(double x, double y, double width, double height, double rotationDegrees, double centerX, double centerY) {
        if (Math.Abs(rotationDegrees) <= 0.000001D) {
            return new DrawingBounds(x, y, x + width, y + height);
        }

        double radians = OfficeGeometry.DegreesToRadians(rotationDegrees);
        OfficePoint topLeft = OfficeGeometry.RotatePoint(new OfficePoint(x, y), centerX, centerY, radians);
        OfficePoint topRight = OfficeGeometry.RotatePoint(new OfficePoint(x + width, y), centerX, centerY, radians);
        OfficePoint bottomRight = OfficeGeometry.RotatePoint(new OfficePoint(x + width, y + height), centerX, centerY, radians);
        OfficePoint bottomLeft = OfficeGeometry.RotatePoint(new OfficePoint(x, y + height), centerX, centerY, radians);
        double left = Math.Min(Math.Min(topLeft.X, topRight.X), Math.Min(bottomRight.X, bottomLeft.X));
        double top = Math.Min(Math.Min(topLeft.Y, topRight.Y), Math.Min(bottomRight.Y, bottomLeft.Y));
        double right = Math.Max(Math.Max(topLeft.X, topRight.X), Math.Max(bottomRight.X, bottomLeft.X));
        double bottom = Math.Max(Math.Max(topLeft.Y, topRight.Y), Math.Max(bottomRight.Y, bottomLeft.Y));
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
