using OfficeIMO.Drawing;

namespace OfficeIMO.Excel.Utilities {
    internal static class ExcelWorksheetDrawingObjectVisual {
        // Geometry is supplied by each export projection; shape and text styling stay shared.
        internal static ExcelVisualDrawingObject ToVisual(this ExcelWorksheetDrawingObjectInfo drawing,
            double x, double y, double width, double height, string source) =>
            new ExcelVisualDrawingObject(
                    drawing.Name,
                    drawing.Order,
                    drawing.ShapePresetName,
                    drawing.ShapeKind!.Value,
                    drawing.HorizontalFlip,
                    drawing.VerticalFlip,
                    drawing.RotationDegrees,
                    x,
                    y,
                    width,
                    height,
                    drawing.FillColorArgb,
                    drawing.StrokeColorArgb,
                    drawing.StrokeWidth,
                    drawing.StrokeDashStyle,
                    drawing.StrokeLineCap,
                    drawing.StrokeLineJoin,
                    drawing.Text,
                    drawing.TextAlignment,
                    drawing.TextVerticalAlignment,
                    drawing.TextColorArgb,
                    drawing.TextFontFamily,
                    drawing.TextFontSize.HasValue ? OfficeLength.Points(drawing.TextFontSize.Value) : null,
                    drawing.TextFontStyle,
                    drawing.TextWrap,
                    drawing.TextShrinkToFit,
                    drawing.TextResizeShapeToFit,
                    drawing.TextOrientation,
                    drawing.TextInsetLeft,
                    drawing.TextInsetTop,
                    drawing.TextInsetRight,
                    drawing.TextInsetBottom,
                    drawing.Glow,
                    drawing.Shadow,
                    source);
    }
}
