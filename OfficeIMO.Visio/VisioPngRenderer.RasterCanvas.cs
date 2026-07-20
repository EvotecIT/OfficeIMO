using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Drawing;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    internal static partial class VisioPngRenderer {

        private sealed class RasterCanvas {
            private readonly OfficeRasterRenderTarget _target;
            private readonly OfficeRasterCanvas _canvas;

            internal RasterCanvas(
                int width,
                int height,
                int supersampling,
                Color? background,
                OfficeTrueTypeFont? outlineFont,
                OfficeFontFaceCollection? fonts,
                IOfficeTextShapingProvider? textShapingProvider,
                string? textShapingLanguage,
                ICollection<OfficeImageExportDiagnostic>? diagnosticSink,
                string? diagnosticSource,
                System.Threading.CancellationToken cancellationToken) {
                Supersampling = supersampling;
                Scale = supersampling;
                _target = new OfficeRasterRenderTarget(width, height, supersampling, background);
                _canvas = new OfficeRasterCanvas(
                    _target,
                    outlineFont,
                    fonts,
                    textShapingProvider,
                    textShapingLanguage,
                    diagnosticSink,
                    diagnosticSource,
                    cancellationToken);
            }

            internal double Scale { get; set; }

            internal int Supersampling { get; }

            internal OfficeTrueTypeFont? OutlineFont => _canvas.OutlineFont;

            internal OfficeFontFaceCollection? Fonts => _canvas.Fonts;

            internal IOfficeTextShapingProvider? TextShapingProvider => _canvas.TextShapingProvider;

            internal string? TextShapingLanguage => _canvas.TextShapingLanguage;

            internal System.Threading.CancellationToken CancellationToken => _canvas.CancellationToken;

            internal void FillPolygon(IReadOnlyList<(double X, double Y)> points, Color color) {
                _canvas.FillPolygon(points, color);
            }

            internal void FillPolygonsEvenOdd(IReadOnlyList<List<(double X, double Y)>> contours, Color color) {
                _canvas.FillPolygonsEvenOdd(contours, color);
            }

            internal void StrokePolygon(IReadOnlyList<(double X, double Y)> points, Color color, double width, OfficeStrokeDashStyle dashStyle) {
                _canvas.DrawStyledPolygon(points, color, width, dashStyle, resetDashPatternForEachSegment: true);
            }

            internal void StrokePolyline(IReadOnlyList<(double X, double Y)> points, Color color, double width, OfficeStrokeDashStyle dashStyle) {
                _canvas.DrawStyledPolyline(points, color, width, dashStyle, resetDashPatternForEachSegment: true);
            }

            internal void DrawEllipse(double cx, double cy, double rx, double ry, Color fill, Color stroke, double width, OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid, double rotationRadians = 0D, double rotationCenterX = 0D, double rotationCenterY = 0D) =>
                _canvas.DrawStyledEllipse(cx, cy, rx, ry, fill, stroke, width, dashStyle, RadiansToCanvasDegrees(rotationRadians), rotationCenterX, rotationCenterY);

            internal void DrawArc(double cx, double cy, double rx, double ry, double startDegrees, double endDegrees, Color color, double width, double rotationRadians = 0D, double rotationCenterX = 0D, double rotationCenterY = 0D) {
                _canvas.DrawArc(cx, cy, rx, ry, startDegrees, endDegrees, color, width, RadiansToCanvasDegrees(rotationRadians), rotationCenterX, rotationCenterY);
            }

            internal void DrawImage(OfficeRasterImage image, double x, double y, double width, double height) =>
                DrawImage(image, x, y, width, height, 0D, x + (width / 2D), y + (height / 2D));

            internal void DrawImage(OfficeRasterImage image, double x, double y, double width, double height, double rotationRadians, double rotationCenterX, double rotationCenterY) {
                _canvas.DrawImage(image, x, y, width, height, RadiansToCanvasDegrees(rotationRadians), rotationCenterX, rotationCenterY);
            }

            internal void DrawImage(OfficeRasterImage image, OfficeImageProjection projection) {
                _canvas.DrawImage(image, projection);
            }

            internal double MeasureText(
                string? text,
                double height,
                string? fontFamily,
                OfficeFontStyle style) =>
                _canvas.MeasureText(text, height, fontFamily, style);

            internal void DrawTextLine(string text, double anchorX, double top, double height, Color color, bool bold, bool italic, OfficeTextAlignment alignment, double rotationRadians, double rotationCenterX, double rotationCenterY) {
                _canvas.DrawTextLine(text, anchorX, top, height, color, bold, italic, alignment, RadiansToCanvasDegrees(rotationRadians), rotationCenterX, rotationCenterY);
            }

            internal void DrawTextBox(
                OfficeTextBlockRenderPlan plan,
                Color color,
                bool bold,
                bool italic,
                bool underline,
                string? fontFamily,
                double rotationRadians,
                double rotationCenterX,
                double rotationCenterY,
                Color? backgroundColor,
                double backgroundPaddingX,
                double backgroundPaddingY) {
                OfficeTextBlockRenderer.DrawRasterTextBox(
                    _canvas,
                    plan,
                    color,
                    bold,
                    italic,
                    underline,
                    rotationDegrees: RadiansToCanvasDegrees(rotationRadians),
                    rotationCenterX: rotationCenterX,
                    rotationCenterY: rotationCenterY,
                    backgroundColor: backgroundColor,
                    backgroundPaddingX: backgroundPaddingX,
                    backgroundPaddingY: backgroundPaddingY,
                    centerLineInLineHeight: false,
                    underlineOffsetFactor: 0.92D,
                    fontFamily: fontFamily);
            }

            internal byte[] Resolve() {
                return _target.ResolveRgba();
            }

            internal void FillRectangle(double x, double y, double width, double height, Color color) {
                _canvas.FillRectangle(x, y, width, height, color);
            }

            private static double RadiansToCanvasDegrees(double radians) => -OfficeGeometry.RadiansToDegrees(radians);

        }
    }
}
