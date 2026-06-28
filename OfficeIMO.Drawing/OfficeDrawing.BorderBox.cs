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

        AddHorizontalBorderLine(x, y, width, borders.Top, transform);
        AddVerticalBorderLine(x + Math.Max(0D, width - 1D), y, height, borders.Right, transform);
        AddHorizontalBorderLine(x, y + Math.Max(0D, height - 1D), width, borders.Bottom, transform);
        AddVerticalBorderLine(x, y, height, borders.Left, transform);
        AddDiagonalBorderLine(x, y, width, height, borders.DiagonalDown, diagonalDown: true, transform);
        AddDiagonalBorderLine(x, y, width, height, borders.DiagonalUp, diagonalDown: false, transform);
        return this;
    }

    private void AddHorizontalBorderLine(double x, double y, double width, OfficeBorderSide? side, OfficeTransform? transform) {
        if (!side.HasValue || !side.Value.IsVisible) {
            return;
        }

        AddBorderLine(x, y, x + width, y, side.Value, transform);
    }

    private void AddVerticalBorderLine(double x, double y, double height, OfficeBorderSide? side, OfficeTransform? transform) {
        if (!side.HasValue || !side.Value.IsVisible) {
            return;
        }

        AddBorderLine(x, y, x, y + height, side.Value, transform);
    }

    private void AddDiagonalBorderLine(double x, double y, double width, double height, OfficeBorderSide? side, bool diagonalDown, OfficeTransform? transform) {
        if (!side.HasValue || !side.Value.IsVisible) {
            return;
        }

        if (diagonalDown) {
            AddBorderLine(x, y, x + width, y + height, side.Value, transform);
            return;
        }

        AddBorderLine(x, y + height, x + width, y, side.Value, transform);
    }

    private void AddBorderLine(double x1, double y1, double x2, double y2, OfficeBorderSide side, OfficeTransform? transform) {
        if (side.LineKind == OfficeBorderLineKind.Double) {
            double separation = side.DoubleLineSeparation > 0D ? side.DoubleLineSeparation : Math.Max(1D, side.Width * 3D);
            if (OfficeGeometry.TryGetParallelLineOffsets(x1, y1, x2, y2, separation, out double offsetX, out double offsetY)) {
                AddSingleBorderLine(x1 - offsetX, y1 - offsetY, x2 - offsetX, y2 - offsetY, side, transform);
                AddSingleBorderLine(x1 + offsetX, y1 + offsetY, x2 + offsetX, y2 + offsetY, side, transform);
            }

            return;
        }

        AddSingleBorderLine(x1, y1, x2, y2, side, transform);
    }

    private void AddSingleBorderLine(double x1, double y1, double x2, double y2, OfficeBorderSide side, OfficeTransform? transform) {
        OfficeShape line = OfficeShape.Line(x1, y1, x2, y2);
        line.StrokeColor = side.Color;
        line.StrokeWidth = side.Width;
        line.StrokeDashStyle = side.DashStyle;
        line.Transform = transform;
        AddShape(line, Math.Min(x1, x2), Math.Min(y1, y2));
    }
}
