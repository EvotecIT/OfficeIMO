using System;
using System.Collections.Generic;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal static partial class ExcelRangeImageRenderer {
        private static void RenderRasterDrawingObject(OfficeRasterCanvas canvas, ExcelVisualDrawingObject drawingObject, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics) {
            AddRotatedTextApproximationDiagnostic(drawingObject, diagnostics);
            AddTextAutoFitUnsupportedDiagnostic(drawingObject, diagnostics);
            AddTextVerticalOrientationUnsupportedDiagnostic(drawingObject, diagnostics);
            double scale = options.Scale;
            DrawingObjectScene scene = CreateOfficeDrawing(drawingObject, scale);
            ExcelImageExportOptions nestedOptions = options.Clone();
            nestedOptions.Scale = 1D;
            OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(
                scene.Drawing.Width,
                scene.Drawing.Height,
                OfficeImageExportFormat.Png,
                nestedOptions,
                drawingObject.Source);
            if (plan.Diagnostic != null) diagnostics?.Add(plan.Diagnostic);
            OfficeRasterImage drawingImage = OfficeDrawingRasterRenderer.Render(scene.Drawing, plan.Limit.Scale);
            canvas.DrawImage(
                drawingImage,
                (drawingObject.X * scale) - scene.OffsetX,
                (drawingObject.Y * scale) - scene.OffsetY,
                scene.Drawing.Width,
                scene.Drawing.Height);
        }

        private static void AppendSvgDrawingObject(StringBuilder builder, ExcelVisualDrawingObject drawingObject, ExcelImageExportOptions options, List<OfficeImageExportDiagnostic>? diagnostics) {
            AddRotatedTextApproximationDiagnostic(drawingObject, diagnostics);
            AddTextAutoFitUnsupportedDiagnostic(drawingObject, diagnostics);
            AddTextVerticalOrientationUnsupportedDiagnostic(drawingObject, diagnostics);
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
            shape.StrokeDashStyle = drawingObject.StrokeDashStyle;
            shape.StrokeLineCap = drawingObject.StrokeLineCap;
            shape.StrokeLineJoin = drawingObject.StrokeLineJoin;
            shape.Glow = ScaleGlow(drawingObject.Glow, scale);
            shape.Shadow = ScaleShadow(drawingObject.Shadow, scale);
            double offsetX = 0D;
            double offsetY = 0D;
            if (drawingObject.HasRotation) {
                shape.Transform = OfficeTransform.RotateDegrees(drawingObject.RotationDegrees, width / 2D, height / 2D);
                ExpandRotatedShapeBounds(width, height, drawingObject.RotationDegrees, shape.StrokeWidth, out offsetX, out offsetY);
            }

            ExpandEffectBounds(shape, ref offsetX, ref offsetY);
            var drawing = new OfficeDrawing(width + (offsetX * 2D), height + (offsetY * 2D));
            drawing.AddShape(shape, offsetX, offsetY);

            if (!string.IsNullOrWhiteSpace(drawingObject.Text)) {
                double insetLeft = Math.Max(0D, drawingObject.TextInsetLeft * scale);
                double insetTop = Math.Max(0D, drawingObject.TextInsetTop * scale);
                double insetRight = Math.Max(0D, drawingObject.TextInsetRight * scale);
                double insetBottom = Math.Max(0D, drawingObject.TextInsetBottom * scale);
                double textWidth = Math.Max(1D, width - insetLeft - insetRight);
                double textHeight = Math.Max(1D, height - insetTop - insetBottom);
                OfficeFontInfo font = ResolveDrawingTextFont(drawingObject, scale, textHeight);
                double textRotationDegrees = ResolveDrawingTextRotationDegrees(drawingObject);
                drawing.AddText(
                    drawingObject.Text,
                    offsetX + insetLeft,
                    offsetY + insetTop,
                    textWidth,
                    textHeight,
                    font,
                    ResolveArgb(drawingObject.TextColorArgb) ?? OfficeColor.FromRgb(31, 41, 55),
                    drawingObject.TextAlignment,
                    verticalAlignment: drawingObject.TextVerticalAlignment,
                    rotationDegrees: textRotationDegrees,
                    rotationCenterX: offsetX + width / 2D,
                    rotationCenterY: offsetY + height / 2D,
                    wrapText: drawingObject.TextWrap,
                    shrinkToFit: drawingObject.TextShrinkToFit,
                    stackedText: IsSupportedStackedTextOrientation(drawingObject));
            }

            return new DrawingObjectScene(drawing, offsetX, offsetY);
        }

        private static OfficeFontInfo ResolveDrawingTextFont(ExcelVisualDrawingObject drawingObject, double scale, double textHeight) {
            string family = string.IsNullOrWhiteSpace(drawingObject.TextFontFamily)
                ? "Calibri"
                : drawingObject.TextFontFamily!;
            double fontSize = drawingObject.TextFontSize.HasValue && drawingObject.TextFontSize.Value > 0D
                ? Math.Max(1D, drawingObject.TextFontSize.Value * scale)
                : Math.Max(7D, Math.Min(11D * scale, textHeight * 0.55D));
            return new OfficeFontInfo(family, fontSize, drawingObject.TextFontStyle);
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
            if (diagnostics == null || !UsesDrawingTextRotation(drawingObject) || string.IsNullOrWhiteSpace(drawingObject.Text)) {
                return;
            }

            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation,
                "Worksheet drawing object text is rendered through shared Drawing rotation without Excel-exact text-box metrics.",
                drawingObject.Source));
        }

        private static void AddTextVerticalOrientationUnsupportedDiagnostic(ExcelVisualDrawingObject drawingObject, List<OfficeImageExportDiagnostic>? diagnostics) {
            if (diagnostics == null || IsSupportedTextOrientation(drawingObject) || string.IsNullOrWhiteSpace(drawingObject.Text)) {
                return;
            }

            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                ExcelImageExportDiagnosticCodes.DrawingShapeTextVerticalOrientationUnsupported,
                "Worksheet drawing object text requested a non-horizontal orientation; image export renders it as horizontal text inside the authored shape bounds.",
                drawingObject.Source));
        }

        private static bool IsSupportedTextOrientation(ExcelVisualDrawingObject drawingObject) =>
            drawingObject.TextOrientation == ExcelDrawingTextOrientation.Horizontal ||
            drawingObject.TextOrientation == ExcelDrawingTextOrientation.Vertical270 ||
            IsSupportedStackedTextOrientation(drawingObject);

        private static bool IsSupportedStackedTextOrientation(ExcelVisualDrawingObject drawingObject) =>
            drawingObject.TextOrientation == ExcelDrawingTextOrientation.Vertical;

        private static bool UsesDrawingTextRotation(ExcelVisualDrawingObject drawingObject) =>
            drawingObject.HasRotation ||
            drawingObject.TextOrientation == ExcelDrawingTextOrientation.Vertical270;

        private static double ResolveDrawingTextRotationDegrees(ExcelVisualDrawingObject drawingObject) {
            double rotation = drawingObject.RotationDegrees;
            if (drawingObject.TextOrientation == ExcelDrawingTextOrientation.Vertical270) {
                rotation += 270D;
            }

            return NormalizeRotationDegrees(rotation);
        }

        private static double NormalizeRotationDegrees(double rotationDegrees) {
            double normalized = rotationDegrees % 360D;
            return normalized < 0D
                ? normalized + 360D
                : normalized;
        }

        private static OfficeGlow? ScaleGlow(OfficeGlow? glow, double scale) =>
            glow == null
                ? null
                : new OfficeGlow(glow.Color, glow.Opacity, glow.Radius * scale);

        private static OfficeShadow? ScaleShadow(OfficeShadow? shadow, double scale) =>
            shadow == null
                ? null
                : new OfficeShadow(shadow.Color, shadow.Opacity, shadow.OffsetX * scale, shadow.OffsetY * scale, shadow.BlurRadius * scale);

        private static void ExpandEffectBounds(OfficeShape shape, ref double offsetX, ref double offsetY) {
            bool hasRotatedShapePadding = offsetX > 0D || offsetY > 0D;
            if (shape.Glow != null) {
                if (hasRotatedShapePadding) {
                    offsetX += shape.Glow.Radius;
                    offsetY += shape.Glow.Radius;
                } else {
                    offsetX = Math.Max(offsetX, shape.Glow.Radius);
                    offsetY = Math.Max(offsetY, shape.Glow.Radius);
                }
            }

            if (shape.Shadow != null) {
                double horizontal = Math.Abs(shape.Shadow.OffsetX) + shape.Shadow.BlurRadius;
                double vertical = Math.Abs(shape.Shadow.OffsetY) + shape.Shadow.BlurRadius;
                if (hasRotatedShapePadding) {
                    offsetX += horizontal;
                    offsetY += vertical;
                } else {
                    offsetX = Math.Max(offsetX, horizontal);
                    offsetY = Math.Max(offsetY, vertical);
                }
            }
        }

        private static void AddTextAutoFitUnsupportedDiagnostic(ExcelVisualDrawingObject drawingObject, List<OfficeImageExportDiagnostic>? diagnostics) {
            if (diagnostics == null || !drawingObject.TextResizeShapeToFit || string.IsNullOrWhiteSpace(drawingObject.Text)) {
                return;
            }

            diagnostics.Add(new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning,
                ExcelImageExportDiagnosticCodes.DrawingShapeTextAutoFitUnsupported,
                "Worksheet drawing object text requested resizing the shape to fit text; image export keeps the authored shape bounds and renders text inside them.",
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
