using System.Collections.Generic;
using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static void RenderNativeShape(INativePdfFlow pdf, WordShape shape) {
            OfficeShape? nativeShape = CreateNativeShape(shape);
            if (nativeShape == null) {
                return;
            }

            pdf.Shape(nativeShape, PdfCore.PdfAlign.Left, spacingAfter: 6);
        }

        private static OfficeShape? CreateNativeShape(WordShape shape) {
            if (shape == null || shape.Hidden == true) {
                return null;
            }

            OfficeShape? nativeShape;
            if (shape.Line != null) {
                (double x1, double y1) = ParseNativeShapePoint(shape.Line.From?.Value ?? "0pt,0pt");
                (double x2, double y2) = ParseNativeShapePoint(shape.Line.To?.Value ?? "0pt,0pt");
                nativeShape = OfficeShape.Line(x1, y1, x2, y2);
            } else if (shape._polygon != null && TryCreateNativePolygonShape(shape._polygon.Points?.Value, out nativeShape)) {
            } else {
                (double Width, double Height)? dimensions = GetNativeShapeDimensions(shape);
                if (!dimensions.HasValue) {
                    return null;
                }

                double width = dimensions.Value.Width;
                double height = dimensions.Value.Height;
                if (shape._ellipse != null) {
                    nativeShape = OfficeShape.Ellipse(width, height);
                } else if (shape._roundRectangle != null) {
                    double arcSize = shape.ArcSize ?? 0.25D;
                    double cornerRadius = Math.Min(width, height) * Math.Max(0D, Math.Min(1D, arcSize)) / 2D;
                    nativeShape = OfficeShape.RoundedRectangle(width, height, cornerRadius);
                } else if (TryGetNativeDrawingPreset(shape, out A.ShapeTypeValues preset)) {
                    nativeShape = CreateNativeDrawingPresetShape(preset, width, height);
                    if (nativeShape == null) {
                        return null;
                    }
                } else {
                    nativeShape = OfficeShape.Rectangle(width, height);
                }
            }

            if (nativeShape == null) {
                return null;
            }

            ApplyNativeShapeStyle(nativeShape, shape);
            return nativeShape;
        }

        private static (double Width, double Height)? GetNativeShapeDimensions(WordShape shape) {
            double width = shape.Width;
            double height = shape.Height;
            if (width > 0 && height > 0) {
                return (width, height);
            }

            A.Extents? extents = shape._wpsShape?
                .GetFirstChild<Wps.ShapeProperties>()?
                .GetFirstChild<A.Transform2D>()?
                .Extents;

            long? cx = extents?.Cx?.Value;
            long? cy = extents?.Cy?.Value;
            if (!cx.HasValue || !cy.HasValue || cx.Value <= 0 || cy.Value <= 0) {
                return null;
            }

            return (ConvertNativeEmusToPoints(cx.Value), ConvertNativeEmusToPoints(cy.Value));
        }

        private static bool TryGetNativeDrawingPreset(WordShape shape, out A.ShapeTypeValues preset) {
            A.PresetGeometry? geometry = shape._wpsShape?
                .GetFirstChild<Wps.ShapeProperties>()?
                .GetFirstChild<A.PresetGeometry>();
            if (geometry?.Preset?.Value is A.ShapeTypeValues value) {
                preset = value;
                return true;
            }

            preset = default;
            return false;
        }

        private static OfficeShape? CreateNativeDrawingPresetShape(A.ShapeTypeValues preset, double width, double height) {
            if (preset == A.ShapeTypeValues.Line) {
                return OfficeShape.Line(0, height / 2D, width, height / 2D);
            }

            if (preset == A.ShapeTypeValues.Ellipse) {
                return OfficeShape.Ellipse(width, height);
            }

            if (preset == A.ShapeTypeValues.RoundRectangle) {
                return OfficeShape.RoundedRectangle(width, height, Math.Min(width, height) / 6D);
            }

            if (preset == A.ShapeTypeValues.Triangle) {
                return OfficeShape.Polygon(
                    new OfficePoint(width / 2D, 0),
                    new OfficePoint(width, height),
                    new OfficePoint(0, height));
            }

            if (preset == A.ShapeTypeValues.Diamond) {
                return OfficeShape.Polygon(
                    new OfficePoint(width / 2D, 0),
                    new OfficePoint(width, height / 2D),
                    new OfficePoint(width / 2D, height),
                    new OfficePoint(0, height / 2D));
            }

            if (preset == A.ShapeTypeValues.Pentagon) {
                return CreateRegularNativePolygon(5, width, height, -90D);
            }

            if (preset == A.ShapeTypeValues.Hexagon) {
                return OfficeShape.Polygon(
                    new OfficePoint(width * 0.25D, 0),
                    new OfficePoint(width * 0.75D, 0),
                    new OfficePoint(width, height / 2D),
                    new OfficePoint(width * 0.75D, height),
                    new OfficePoint(width * 0.25D, height),
                    new OfficePoint(0, height / 2D));
            }

            if (preset == A.ShapeTypeValues.RightArrow) {
                return OfficeShape.Polygon(
                    new OfficePoint(0, height * 0.25D),
                    new OfficePoint(width * 0.6D, height * 0.25D),
                    new OfficePoint(width * 0.6D, 0),
                    new OfficePoint(width, height / 2D),
                    new OfficePoint(width * 0.6D, height),
                    new OfficePoint(width * 0.6D, height * 0.75D),
                    new OfficePoint(0, height * 0.75D));
            }

            if (preset == A.ShapeTypeValues.LeftArrow) {
                return OfficeShape.Polygon(
                    new OfficePoint(width, height * 0.25D),
                    new OfficePoint(width * 0.4D, height * 0.25D),
                    new OfficePoint(width * 0.4D, 0),
                    new OfficePoint(0, height / 2D),
                    new OfficePoint(width * 0.4D, height),
                    new OfficePoint(width * 0.4D, height * 0.75D),
                    new OfficePoint(width, height * 0.75D));
            }

            if (preset == A.ShapeTypeValues.UpArrow) {
                return OfficeShape.Polygon(
                    new OfficePoint(width * 0.25D, height),
                    new OfficePoint(width * 0.25D, height * 0.4D),
                    new OfficePoint(0, height * 0.4D),
                    new OfficePoint(width / 2D, 0),
                    new OfficePoint(width, height * 0.4D),
                    new OfficePoint(width * 0.75D, height * 0.4D),
                    new OfficePoint(width * 0.75D, height));
            }

            if (preset == A.ShapeTypeValues.DownArrow) {
                return OfficeShape.Polygon(
                    new OfficePoint(width * 0.25D, 0),
                    new OfficePoint(width * 0.25D, height * 0.6D),
                    new OfficePoint(0, height * 0.6D),
                    new OfficePoint(width / 2D, height),
                    new OfficePoint(width, height * 0.6D),
                    new OfficePoint(width * 0.75D, height * 0.6D),
                    new OfficePoint(width * 0.75D, 0));
            }

            if (preset == A.ShapeTypeValues.Star5) {
                return CreateNativeStar5(width, height);
            }

            if (preset == A.ShapeTypeValues.Heart) {
                return CreateNativeHeart(width, height);
            }

            if (preset == A.ShapeTypeValues.Cloud) {
                return CreateNativeCloud(width, height);
            }

            if (preset == A.ShapeTypeValues.Donut) {
                return CreateNativeDonut(width, height);
            }

            if (preset == A.ShapeTypeValues.Can) {
                return CreateNativeCan(width, height);
            }

            if (preset == A.ShapeTypeValues.Cube) {
                return CreateNativeCube(width, height);
            }

            if (preset == A.ShapeTypeValues.Rectangle) {
                return OfficeShape.Rectangle(width, height);
            }

            return null;
        }

        private static OfficeShape CreateRegularNativePolygon(int sides, double width, double height, double startAngleDegrees) {
            var points = new OfficePoint[sides];
            double centerX = width / 2D;
            double centerY = height / 2D;
            double radiusX = width / 2D;
            double radiusY = height / 2D;
            for (int i = 0; i < sides; i++) {
                double angle = (startAngleDegrees + (360D * i / sides)) * Math.PI / 180D;
                points[i] = new OfficePoint(centerX + radiusX * Math.Cos(angle), centerY + radiusY * Math.Sin(angle));
            }

            return OfficeShape.Polygon(points);
        }

        private static OfficeShape CreateNativeHeart(double width, double height) =>
            OfficeShape.Path(
                OfficePathCommand.MoveTo(width * 0.5D, height),
                OfficePathCommand.CubicBezierTo(width * 0.18D, height * 0.72D, 0, height * 0.52D, 0, height * 0.28D),
                OfficePathCommand.CubicBezierTo(0, height * 0.08D, width * 0.16D, 0, width * 0.31D, 0),
                OfficePathCommand.CubicBezierTo(width * 0.42D, 0, width * 0.49D, height * 0.07D, width * 0.5D, height * 0.18D),
                OfficePathCommand.CubicBezierTo(width * 0.51D, height * 0.07D, width * 0.58D, 0, width * 0.69D, 0),
                OfficePathCommand.CubicBezierTo(width * 0.84D, 0, width, height * 0.08D, width, height * 0.28D),
                OfficePathCommand.CubicBezierTo(width, height * 0.52D, width * 0.82D, height * 0.72D, width * 0.5D, height),
                OfficePathCommand.Close());

        private static OfficeShape CreateNativeCloud(double width, double height) =>
            OfficeShape.Path(
                OfficePathCommand.MoveTo(width * 0.18D, height * 0.7D),
                OfficePathCommand.CubicBezierTo(width * 0.05D, height * 0.7D, 0, height * 0.58D, width * 0.09D, height * 0.48D),
                OfficePathCommand.CubicBezierTo(width * 0.03D, height * 0.32D, width * 0.19D, height * 0.18D, width * 0.34D, height * 0.26D),
                OfficePathCommand.CubicBezierTo(width * 0.42D, height * 0.04D, width * 0.72D, height * 0.08D, width * 0.75D, height * 0.32D),
                OfficePathCommand.CubicBezierTo(width * 0.94D, height * 0.27D, width, height * 0.46D, width * 0.91D, height * 0.61D),
                OfficePathCommand.CubicBezierTo(width * 0.84D, height * 0.75D, width * 0.63D, height * 0.76D, width * 0.54D, height * 0.68D),
                OfficePathCommand.CubicBezierTo(width * 0.46D, height * 0.82D, width * 0.25D, height * 0.82D, width * 0.18D, height * 0.7D),
                OfficePathCommand.Close());

        private static OfficeShape CreateNativeDonut(double width, double height) {
            List<OfficePathCommand> commands = CreateNativeEllipsePath(width / 2D, height / 2D, width / 2D, height / 2D, clockwise: true);
            commands.AddRange(CreateNativeEllipsePath(width / 2D, height / 2D, width * 0.22D, height * 0.22D, clockwise: false));
            return OfficeShape.Path(commands);
        }

        private static OfficeShape CreateNativeCan(double width, double height) {
            double topY = height * 0.18D;
            double bottomY = height * 0.82D;
            double rx = width / 2D;
            double ry = height * 0.14D;
            double k = 0.5522847498307936D;
            return OfficeShape.Path(
                OfficePathCommand.MoveTo(0, topY),
                OfficePathCommand.CubicBezierTo(0, topY - ry * k, rx - rx * k, topY - ry, rx, topY - ry),
                OfficePathCommand.CubicBezierTo(rx + rx * k, topY - ry, width, topY - ry * k, width, topY),
                OfficePathCommand.LineTo(width, bottomY),
                OfficePathCommand.CubicBezierTo(width, bottomY + ry * k, rx + rx * k, bottomY + ry, rx, bottomY + ry),
                OfficePathCommand.CubicBezierTo(rx - rx * k, bottomY + ry, 0, bottomY + ry * k, 0, bottomY),
                OfficePathCommand.Close());
        }

        private static OfficeShape CreateNativeCube(double width, double height) =>
            OfficeShape.Polygon(
                new OfficePoint(width * 0.32D, 0),
                new OfficePoint(width, height * 0.18D),
                new OfficePoint(width, height * 0.72D),
                new OfficePoint(width * 0.62D, height),
                new OfficePoint(0, height * 0.82D),
                new OfficePoint(0, height * 0.28D));

        private static List<OfficePathCommand> CreateNativeEllipsePath(double centerX, double centerY, double radiusX, double radiusY, bool clockwise) {
            double k = 0.5522847498307936D;
            if (clockwise) {
                return new List<OfficePathCommand> {
                    OfficePathCommand.MoveTo(centerX + radiusX, centerY),
                    OfficePathCommand.CubicBezierTo(centerX + radiusX, centerY + radiusY * k, centerX + radiusX * k, centerY + radiusY, centerX, centerY + radiusY),
                    OfficePathCommand.CubicBezierTo(centerX - radiusX * k, centerY + radiusY, centerX - radiusX, centerY + radiusY * k, centerX - radiusX, centerY),
                    OfficePathCommand.CubicBezierTo(centerX - radiusX, centerY - radiusY * k, centerX - radiusX * k, centerY - radiusY, centerX, centerY - radiusY),
                    OfficePathCommand.CubicBezierTo(centerX + radiusX * k, centerY - radiusY, centerX + radiusX, centerY - radiusY * k, centerX + radiusX, centerY),
                    OfficePathCommand.Close()
                };
            }

            return new List<OfficePathCommand> {
                OfficePathCommand.MoveTo(centerX + radiusX, centerY),
                OfficePathCommand.CubicBezierTo(centerX + radiusX, centerY - radiusY * k, centerX + radiusX * k, centerY - radiusY, centerX, centerY - radiusY),
                OfficePathCommand.CubicBezierTo(centerX - radiusX * k, centerY - radiusY, centerX - radiusX, centerY - radiusY * k, centerX - radiusX, centerY),
                OfficePathCommand.CubicBezierTo(centerX - radiusX, centerY + radiusY * k, centerX - radiusX * k, centerY + radiusY, centerX, centerY + radiusY),
                OfficePathCommand.CubicBezierTo(centerX + radiusX * k, centerY + radiusY, centerX + radiusX, centerY + radiusY * k, centerX + radiusX, centerY),
                OfficePathCommand.Close()
            };
        }

        private static OfficeShape CreateNativeStar5(double width, double height) {
            var points = new OfficePoint[10];
            double centerX = width / 2D;
            double centerY = height / 2D;
            double outerX = width / 2D;
            double outerY = height / 2D;
            double innerX = outerX * 0.45D;
            double innerY = outerY * 0.45D;
            for (int i = 0; i < points.Length; i++) {
                bool outer = i % 2 == 0;
                double angle = (-90D + 36D * i) * Math.PI / 180D;
                double radiusX = outer ? outerX : innerX;
                double radiusY = outer ? outerY : innerY;
                points[i] = new OfficePoint(centerX + radiusX * Math.Cos(angle), centerY + radiusY * Math.Sin(angle));
            }

            return OfficeShape.Polygon(points);
        }

        private static bool TryCreateNativePolygonShape(string? pointsText, out OfficeShape? shape) {
            shape = null;
            if (string.IsNullOrWhiteSpace(pointsText)) {
                return false;
            }

            string text = pointsText!;
            var points = new List<OfficePoint>();
            foreach (string token in text.Split(new[] { ' ', ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                string[] parts = token.Split(',');
                if (parts.Length != 2 ||
                    !double.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double x) ||
                    !double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double y)) {
                    return false;
                }

                points.Add(new OfficePoint(x, y));
            }

            if (points.Count < 3) {
                return false;
            }

            shape = OfficeShape.Polygon(points);
            return true;
        }

        private static void ApplyNativeShapeStyle(OfficeShape nativeShape, WordShape wordShape) {
            if (nativeShape.Kind != OfficeShapeKind.Line) {
                PdfCore.PdfColor? fill = ParseNativeColor(wordShape.FillColorHex);
                if (fill.HasValue) {
                    nativeShape.FillColor = fill.Value.ToOfficeColor();
                }
            }

            bool drawStroke = nativeShape.Kind == OfficeShapeKind.Line ||
                              wordShape.Stroked == true ||
                              (wordShape._wpsShape != null && !string.IsNullOrWhiteSpace(wordShape.StrokeColorHex));
            if (!drawStroke) {
                nativeShape.StrokeColor = null;
                nativeShape.StrokeWidth = 0;
                return;
            }

            PdfCore.PdfColor? stroke = ParseNativeColor(wordShape.StrokeColorHex);
            nativeShape.StrokeColor = (stroke ?? PdfCore.PdfColor.Black).ToOfficeColor();
            nativeShape.StrokeWidth = Math.Max(0D, wordShape.StrokeWeight ?? 1D);
        }

        private static (double X, double Y) ParseNativeShapePoint(string value) {
            string[] parts = value.Split(',');
            if (parts.Length != 2) {
                return (0D, 0D);
            }

            return (ParseNativeShapePointPart(parts[0]), ParseNativeShapePointPart(parts[1]));
        }

        private static double ParseNativeShapePointPart(string value) {
            string normalized = value.Trim().Replace("pt", string.Empty);
            return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) ? result : 0D;
        }

    }
}
