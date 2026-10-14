using System;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointTableCell {
    private static void ValidateBorderStyle(PowerPointTableCellBorders borders, string color, double? widthPoints) {
        if ((borders & ~(PowerPointTableCellBorders.All | PowerPointTableCellBorders.Diagonals)) != 0)
            throw new ArgumentOutOfRangeException(nameof(borders));
        if (string.IsNullOrWhiteSpace(color)) throw new ArgumentException("Border color cannot be null or empty.", nameof(color));
        if (widthPoints.HasValue) PowerPointShape.ToDrawingLineWidth(widthPoints.Value, nameof(widthPoints));
    }

    /// <summary>Explicitly hides the selected borders, overriding inherited table styling.
    /// <see cref="ClearBorders"/> removes the local declarations and allows style inheritance instead.</summary>
    public void HideBorders(PowerPointTableCellBorders borders) {
        if ((borders & ~(PowerPointTableCellBorders.All | PowerPointTableCellBorders.Diagonals)) != 0)
            throw new ArgumentOutOfRangeException(nameof(borders));
        if (borders == PowerPointTableCellBorders.None) return;
        var properties = EnsureProperties();
        if (borders.HasFlag(PowerPointTableCellBorders.Left))
            SetBorderLineFill(properties.LeftBorderLineProperties ??= new A.LeftBorderLineProperties(), new A.NoFill());
        if (borders.HasFlag(PowerPointTableCellBorders.Top))
            SetBorderLineFill(properties.TopBorderLineProperties ??= new A.TopBorderLineProperties(), new A.NoFill());
        if (borders.HasFlag(PowerPointTableCellBorders.Right))
            SetBorderLineFill(properties.RightBorderLineProperties ??= new A.RightBorderLineProperties(), new A.NoFill());
        if (borders.HasFlag(PowerPointTableCellBorders.Bottom))
            SetBorderLineFill(properties.BottomBorderLineProperties ??= new A.BottomBorderLineProperties(), new A.NoFill());
        if (borders.HasFlag(PowerPointTableCellBorders.DiagonalDown))
            SetBorderLineFill(properties.TopLeftToBottomRightBorderLineProperties ??= new A.TopLeftToBottomRightBorderLineProperties(), new A.NoFill());
        if (borders.HasFlag(PowerPointTableCellBorders.DiagonalUp))
            SetBorderLineFill(properties.BottomLeftToTopRightBorderLineProperties ??= new A.BottomLeftToTopRightBorderLineProperties(), new A.NoFill());
    }

    private static void SetBorderLineFill(A.LinePropertiesType line, OpenXmlElement fill) {
        line.RemoveAllChildren<A.NoFill>();
        line.RemoveAllChildren<A.SolidFill>();
        line.RemoveAllChildren<A.GradientFill>();
        line.RemoveAllChildren<A.PatternFill>();
        line.InsertAt(fill, 0);
    }
}
