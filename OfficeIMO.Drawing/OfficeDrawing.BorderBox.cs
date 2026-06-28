using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeDrawing {
    /// <summary>
    /// Adds a filled rectangle and independent edge borders in paint order.
    /// </summary>
    public OfficeDrawing AddBorderBox(double x, double y, double width, double height, OfficeColor? fillColor, OfficeBorderBox borders, OfficeTransform? transform = null) {
        ValidateFiniteNonNegative(x, nameof(x));
        ValidateFiniteNonNegative(y, nameof(y));
        ValidatePositiveFinite(width, nameof(width));
        ValidatePositiveFinite(height, nameof(height));
        if (x + width > Width || y + height > Height) {
            throw new ArgumentOutOfRangeException(nameof(width), "Border box must fit inside the drawing bounds.");
        }

        if (fillColor.HasValue && fillColor.Value.A > 0) {
            OfficeShape fill = OfficeShape.Rectangle(width, height);
            fill.FillColor = fillColor.Value;
            fill.StrokeWidth = 0D;
            fill.Transform = transform;
            AddShape(fill, x, y);
        }

        AddHorizontalBorderLine(x, y, width, height, 0D, width, borders.Top, transform);
        AddVerticalBorderLine(x, y, width, height, width, height, borders.Right, transform);
        AddHorizontalBorderLine(x, y, width, height, height, width, borders.Bottom, transform);
        AddVerticalBorderLine(x, y, width, height, 0D, height, borders.Left, transform);
        AddDiagonalBorderLine(x, y, width, height, borders.DiagonalDown, diagonalDown: true, transform);
        AddDiagonalBorderLine(x, y, width, height, borders.DiagonalUp, diagonalDown: false, transform);
        return this;
    }

    private void AddHorizontalBorderLine(double originX, double originY, double boxWidth, double boxHeight, double y, double width, OfficeBorderSide? side, OfficeTransform? transform) {
        if (!side.HasValue || !side.Value.IsVisible) {
            return;
        }

        AddBorderLine(originX, originY, boxWidth, boxHeight, 0D, y, width, y, side.Value, transform);
    }

    private void AddVerticalBorderLine(double originX, double originY, double boxWidth, double boxHeight, double x, double height, OfficeBorderSide? side, OfficeTransform? transform) {
        if (!side.HasValue || !side.Value.IsVisible) {
            return;
        }

        AddBorderLine(originX, originY, boxWidth, boxHeight, x, 0D, x, height, side.Value, transform);
    }

    private void AddDiagonalBorderLine(double originX, double originY, double boxWidth, double boxHeight, OfficeBorderSide? side, bool diagonalDown, OfficeTransform? transform) {
        if (!side.HasValue || !side.Value.IsVisible) {
            return;
        }

        if (diagonalDown) {
            AddBorderLine(originX, originY, boxWidth, boxHeight, 0D, 0D, boxWidth, boxHeight, side.Value, transform);
            return;
        }

        AddBorderLine(originX, originY, boxWidth, boxHeight, 0D, boxHeight, boxWidth, 0D, side.Value, transform);
    }

    private void AddBorderLine(double originX, double originY, double boxWidth, double boxHeight, double x1, double y1, double x2, double y2, OfficeBorderSide side, OfficeTransform? transform) {
        if (side.LineKind == OfficeBorderLineKind.Double) {
            double separation = side.DoubleLineSeparation > 0D ? side.DoubleLineSeparation : Math.Max(1D, side.Width * 3D);
            if (OfficeGeometry.TryGetParallelLineOffsets(x1, y1, x2, y2, separation, out double offsetX, out double offsetY)) {
                AddSingleBorderLine(originX, originY, boxWidth, boxHeight, x1 - offsetX, y1 - offsetY, x2 - offsetX, y2 - offsetY, side, transform);
                AddSingleBorderLine(originX, originY, boxWidth, boxHeight, x1 + offsetX, y1 + offsetY, x2 + offsetX, y2 + offsetY, side, transform);
            }

            return;
        }

        AddSingleBorderLine(originX, originY, boxWidth, boxHeight, x1, y1, x2, y2, side, transform);
    }

    private void AddSingleBorderLine(double originX, double originY, double boxWidth, double boxHeight, double x1, double y1, double x2, double y2, OfficeBorderSide side, OfficeTransform? transform) {
        x1 = Clamp(x1, 0D, boxWidth);
        y1 = Clamp(y1, 0D, boxHeight);
        x2 = Clamp(x2, 0D, boxWidth);
        y2 = Clamp(y2, 0D, boxHeight);
        if (Math.Abs(x1 - x2) < 0.000001D && Math.Abs(y1 - y2) < 0.000001D) {
            return;
        }

        double minX = Math.Min(x1, x2);
        double minY = Math.Min(y1, y2);
        OfficeShape line = OfficeShape.Line(x1, y1, x2, y2);
        line.StrokeColor = side.Color;
        line.StrokeWidth = side.Width;
        line.StrokeDashStyle = side.DashStyle;
        line.Transform = RebaseTransform(transform, minX, minY);
        AddShape(line, originX + minX, originY + minY);
    }

    private static OfficeTransform? RebaseTransform(OfficeTransform? transform, double offsetX, double offsetY) {
        if (!transform.HasValue || transform.Value == OfficeTransform.Identity || (Math.Abs(offsetX) < 0.000001D && Math.Abs(offsetY) < 0.000001D)) {
            return transform;
        }

        OfficeTransform value = transform.Value;
        OfficePoint transformedOffset = value.TransformPoint(new OfficePoint(offsetX, offsetY));
        return new OfficeTransform(
            value.M11,
            value.M12,
            value.M21,
            value.M22,
            transformedOffset.X - offsetX,
            transformedOffset.Y - offsetY);
    }

    private static double Clamp(double value, double min, double max) =>
        value < min ? min : value > max ? max : value;
}
