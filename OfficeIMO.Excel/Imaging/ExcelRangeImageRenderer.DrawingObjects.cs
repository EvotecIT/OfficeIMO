using System;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeImageRenderer {
        private static void RenderRasterDrawingObjects(OfficeRasterCanvas canvas, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options) {
            foreach (ExcelVisualDrawingObject drawingObject in snapshot.DrawingObjects) {
                RenderRasterDrawingObject(canvas, drawingObject, options);
            }
        }

        private static void AppendSvgDrawingObjects(StringBuilder builder, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options) {
            foreach (ExcelVisualDrawingObject drawingObject in snapshot.DrawingObjects) {
                AppendSvgDrawingObject(builder, drawingObject, options);
            }
        }

        private static void RenderRasterDrawingObject(OfficeRasterCanvas canvas, ExcelVisualDrawingObject drawingObject, ExcelImageExportOptions options) {
            double scale = options.Scale;
            OfficeDrawing drawing = CreateOfficeDrawing(drawingObject, scale);
            OfficeRasterImage drawingImage = OfficeDrawingRasterRenderer.Render(drawing);
            canvas.DrawImage(
                drawingImage,
                drawingObject.X * scale,
                drawingObject.Y * scale,
                drawingObject.Width * scale,
                drawingObject.Height * scale);
        }

        private static void AppendSvgDrawingObject(StringBuilder builder, ExcelVisualDrawingObject drawingObject, ExcelImageExportOptions options) {
            double scale = options.Scale;
            double x = drawingObject.X * scale;
            double y = drawingObject.Y * scale;
            double width = drawingObject.Width * scale;
            double height = drawingObject.Height * scale;
            OfficeDrawing drawing = CreateOfficeDrawing(drawingObject, scale);
            builder.AppendNestedSvg(x, y, width, height, OfficeSvgFormatting.ExtractSvgInner(OfficeDrawingSvgExporter.ToSvg(drawing)));
        }

        private static OfficeDrawing CreateOfficeDrawing(ExcelVisualDrawingObject drawingObject, double scale) {
            double width = Math.Max(1D, drawingObject.Width * scale);
            double height = Math.Max(1D, drawingObject.Height * scale);
            var drawing = new OfficeDrawing(width, height);
            OfficeShape shape = CreateOfficeShape(drawingObject, width, height);
            shape.FillColor = ResolveArgb(drawingObject.FillColorArgb);
            shape.StrokeColor = ResolveArgb(drawingObject.StrokeColorArgb);
            shape.StrokeWidth = drawingObject.StrokeWidth <= 0D ? 0D : Math.Max(1D, drawingObject.StrokeWidth * scale);
            drawing.AddShape(shape, 0D, 0D);

            if (!string.IsNullOrWhiteSpace(drawingObject.Text)) {
                double padding = Math.Min(8D * scale, Math.Max(2D, Math.Min(width, height) / 8D));
                double textWidth = Math.Max(1D, width - (padding * 2D));
                double textHeight = Math.Max(1D, height - (padding * 2D));
                double fontSize = Math.Max(7D, Math.Min(11D * scale, textHeight * 0.55D));
                drawing.AddText(
                    drawingObject.Text,
                    padding,
                    padding,
                    textWidth,
                    textHeight,
                    new OfficeFontInfo("Calibri", fontSize),
                    OfficeColor.FromRgb(31, 41, 55),
                    OfficeTextAlignment.Center);
            }

            return drawing;
        }

        private static OfficeShape CreateOfficeShape(ExcelVisualDrawingObject drawingObject, double width, double height) =>
            OfficeShapePresets.TryCreate(
                drawingObject.ShapePresetName,
                width,
                height,
                drawingObject.HorizontalFlip,
                drawingObject.VerticalFlip,
                out OfficeShape? shape) && shape != null
                ? shape
                : OfficeShape.Rectangle(width, height);
    }
}
