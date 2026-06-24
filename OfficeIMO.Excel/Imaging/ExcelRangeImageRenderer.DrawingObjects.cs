using System;
using System.Collections.Generic;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeImageRenderer {
        private static void RenderRasterDrawingObjects(OfficeRasterCanvas canvas, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options) {
            foreach (ExcelVisualDrawingObject drawingObject in snapshot.DrawingObjects) {
                RenderRasterDrawingObject(canvas, drawingObject, options, diagnostics: null);
            }
        }

        private static void AppendSvgDrawingObjects(StringBuilder builder, ExcelRangeVisualSnapshot snapshot, ExcelImageExportOptions options) {
            foreach (ExcelVisualDrawingObject drawingObject in snapshot.DrawingObjects) {
                AppendSvgDrawingObject(builder, drawingObject, options, diagnostics: null);
            }
        }

        private static void RenderRasterDrawingObject(OfficeRasterCanvas canvas, ExcelVisualDrawingObject drawingObject, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics) {
            AddRotatedTextApproximationDiagnostic(drawingObject, diagnostics);
            double scale = options.Scale;
            DrawingObjectScene scene = CreateOfficeDrawing(drawingObject, scale);
            OfficeRasterImage drawingImage = OfficeDrawingRasterRenderer.Render(scene.Drawing);
            canvas.DrawImage(
                drawingImage,
                (drawingObject.X * scale) - scene.OffsetX,
                (drawingObject.Y * scale) - scene.OffsetY,
                scene.Drawing.Width,
                scene.Drawing.Height);
        }

        private static void AppendSvgDrawingObject(StringBuilder builder, ExcelVisualDrawingObject drawingObject, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics) {
            AddRotatedTextApproximationDiagnostic(drawingObject, diagnostics);
            double scale = options.Scale;
            double x = drawingObject.X * scale;
            double y = drawingObject.Y * scale;
            DrawingObjectScene scene = CreateOfficeDrawing(drawingObject, scale);
            builder.AppendNestedSvg(
                x - scene.OffsetX,
                y - scene.OffsetY,
                scene.Drawing.Width,
                scene.Drawing.Height,
                OfficeSvgFormatting.ExtractSvgInner(OfficeDrawingSvgExporter.ToSvg(scene.Drawing)));
        }

        private static DrawingObjectScene CreateOfficeDrawing(ExcelVisualDrawingObject drawingObject, double scale) {
            double width = Math.Max(1D, drawingObject.Width * scale);
            double height = Math.Max(1D, drawingObject.Height * scale);
            OfficeShape shape = CreateOfficeShape(drawingObject, width, height);
            shape.FillColor = ResolveArgb(drawingObject.FillColorArgb);
            shape.StrokeColor = ResolveArgb(drawingObject.StrokeColorArgb);
            shape.StrokeWidth = drawingObject.StrokeWidth <= 0D ? 0D : Math.Max(1D, drawingObject.StrokeWidth * scale);
            double offsetX = 0D;
            double offsetY = 0D;
            if (drawingObject.HasRotation) {
                shape.Transform = OfficeTransform.RotateDegrees(drawingObject.RotationDegrees, width / 2D, height / 2D);
                ExpandRotatedShapeBounds(width, height, drawingObject.RotationDegrees, shape.StrokeWidth, out offsetX, out offsetY);
            }

            var drawing = new OfficeDrawing(width + (offsetX * 2D), height + (offsetY * 2D));
            drawing.AddShape(shape, offsetX, offsetY);

            if (!string.IsNullOrWhiteSpace(drawingObject.Text)) {
                double padding = Math.Min(8D * scale, Math.Max(2D, Math.Min(width, height) / 8D));
                double textWidth = Math.Max(1D, width - (padding * 2D));
                double textHeight = Math.Max(1D, height - (padding * 2D));
                double fontSize = Math.Max(7D, Math.Min(11D * scale, textHeight * 0.55D));
                drawing.AddText(
                    drawingObject.Text,
                    offsetX + padding,
                    offsetY + padding,
                    textWidth,
                    textHeight,
                    new OfficeFontInfo("Calibri", fontSize),
                    ResolveArgb(drawingObject.TextColorArgb) ?? OfficeColor.FromRgb(31, 41, 55),
                    drawingObject.TextAlignment,
                    verticalAlignment: drawingObject.TextVerticalAlignment,
                    rotationDegrees: drawingObject.RotationDegrees,
                    rotationCenterX: offsetX + width / 2D,
                    rotationCenterY: offsetY + height / 2D);
            }

            return new DrawingObjectScene(drawing, offsetX, offsetY);
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

        private static void AddRotatedTextApproximationDiagnostic(ExcelVisualDrawingObject drawingObject, List<OfficeImageExportDiagnostic>? diagnostics) {
            if (diagnostics == null || !drawingObject.HasRotation || string.IsNullOrWhiteSpace(drawingObject.Text)) {
                return;
            }

            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation,
                "Worksheet drawing object text is rendered through shared Drawing rotation without Excel-exact text-box metrics.",
                drawingObject.Source));
        }

        private static void ExpandRotatedShapeBounds(double width, double height, double rotationDegrees, double strokeWidth, out double offsetX, out double offsetY) {
            (double left, double top, double right, double bottom) = OfficeGeometry.GetRotatedRectangleBounds(
                0D,
                0D,
                width,
                height,
                rotationDegrees,
                width / 2D,
                height / 2D);
            double strokePadding = strokeWidth > 0D ? strokeWidth : 0D;
            offsetX = Math.Max(0D, Math.Max(-left, right - width)) + strokePadding;
            offsetY = Math.Max(0D, Math.Max(-top, bottom - height)) + strokePadding;
        }

        private readonly struct DrawingObjectScene {
            internal DrawingObjectScene(OfficeDrawing drawing, double offsetX, double offsetY) {
                Drawing = drawing;
                OffsetX = offsetX;
                OffsetY = offsetY;
            }

            internal OfficeDrawing Drawing { get; }

            internal double OffsetX { get; }

            internal double OffsetY { get; }
        }
    }
}
