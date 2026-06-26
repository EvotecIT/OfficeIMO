using System.Collections.Generic;
using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
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
                } else if (TryGetNativeDrawingPreset(shape, out string? presetName)) {
                    nativeShape = CreateNativeDrawingPresetShape(presetName, width, height);
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

            IReadOnlyDictionary<A.SchemeColorValues, OfficeColor> themeColors = GetNativeDrawingThemeColors(GetNativeShapeSourcePart(shape));
            ApplyNativeShapeStyle(nativeShape, shape, themeColors);
            return nativeShape;
        }

        private static OpenXmlPart? GetNativeShapeSourcePart(WordShape shape) {
            OpenXmlPartRootElement? root = shape.Run?.Ancestors<OpenXmlPartRootElement>().FirstOrDefault();
            return root?.OpenXmlPart;
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

        private static bool TryGetNativeDrawingPreset(WordShape shape, out string? presetName) {
            A.PresetGeometry? geometry = shape._wpsShape?
                .GetFirstChild<Wps.ShapeProperties>()?
                .GetFirstChild<A.PresetGeometry>();

            presetName = geometry?.Preset?.InnerText;
            if (!string.IsNullOrWhiteSpace(presetName)) {
                return true;
            }

            if (geometry?.Preset?.Value is A.ShapeTypeValues value) {
                presetName = value.ToString();
                return true;
            }

            return false;
        }

        private static OfficeShape? CreateNativeDrawingPresetShape(string? presetName, double width, double height) {
            if (IsNativeDrawingLinePreset(presetName)) {
                double y = height / 2D;
                return OfficeShape.Line(0D, y, width, y);
            }

            return OfficeShapePresets.TryCreate(presetName, width, height, out OfficeShape? shape)
                ? shape
                : null;
        }

        private static bool IsNativeDrawingLinePreset(string? presetName) {
            if (string.IsNullOrWhiteSpace(presetName)) {
                return false;
            }

            string normalized = presetName!.Trim();
            return string.Equals(normalized, "line", StringComparison.OrdinalIgnoreCase);
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

        private static void ApplyNativeShapeStyle(OfficeShape nativeShape, WordShape wordShape, IReadOnlyDictionary<A.SchemeColorValues, OfficeColor> themeColors) {
            Wps.ShapeProperties? shapeProperties = wordShape._wpsShape?.GetFirstChild<Wps.ShapeProperties>();
            if (nativeShape.Kind != OfficeShapeKind.Line) {
                if (TryGetNativeDrawingGradientFill(shapeProperties, out OfficeLinearGradient? drawingGradient, themeColors)) {
                    nativeShape.FillGradient = drawingGradient;
                } else if (TryGetNativeDrawingSolidFillColor(shapeProperties, out OfficeColor drawingFill, themeColors)) {
                    nativeShape.FillColor = drawingFill;
                } else {
                    PdfCore.PdfColor? fill = ParseNativeColor(wordShape.FillColorHex);
                    if (fill.HasValue) {
                        nativeShape.FillColor = fill.Value.ToOfficeColor();
                    }
                }

                if (TryGetNativeDrawingFillOpacity(shapeProperties, out double fillOpacity)) {
                    nativeShape.FillOpacity = fillOpacity;
                } else if (nativeShape.FillColor.HasValue && nativeShape.FillColor.Value.A < byte.MaxValue) {
                    nativeShape.FillOpacity = nativeShape.FillColor.Value.A / 255D;
                }
            }

            A.Outline? drawingOutline = shapeProperties?.GetFirstChild<A.Outline>();
            bool drawingOutlineNoFill = drawingOutline?.GetFirstChild<A.NoFill>() != null;
            bool hasDrawingStroke = drawingOutline != null && !drawingOutlineNoFill;
            OfficeColor drawingStroke = default;
            bool hasDrawingStrokeColor = hasDrawingStroke && TryGetNativeDrawingSolidFillColor(drawingOutline, out drawingStroke, themeColors);
            bool drawStroke = !drawingOutlineNoFill &&
                              (nativeShape.Kind == OfficeShapeKind.Line ||
                               wordShape.Stroked == true ||
                               hasDrawingStrokeColor ||
                               (wordShape._wpsShape != null && !string.IsNullOrWhiteSpace(wordShape.StrokeColorHex)));
            if (!drawStroke) {
                nativeShape.StrokeColor = null;
                nativeShape.StrokeWidth = 0;
                return;
            }

            PdfCore.PdfColor? stroke = hasDrawingStrokeColor ? PdfCore.PdfColor.FromOfficeColor(drawingStroke) : ParseNativeColor(wordShape.StrokeColorHex);
            nativeShape.StrokeColor = (stroke ?? PdfCore.PdfColor.Black).ToOfficeColor();
            nativeShape.StrokeWidth = Math.Max(0D, wordShape.StrokeWeight ?? 1D);
            nativeShape.StrokeDashStyle = MapNativeDrawingPresetDash(drawingOutline);
            if (TryGetNativeDrawingFillOpacity(drawingOutline, out double strokeOpacity)) {
                nativeShape.StrokeOpacity = strokeOpacity;
            } else if (hasDrawingStrokeColor && drawingStroke.A < byte.MaxValue) {
                nativeShape.StrokeOpacity = drawingStroke.A / 255D;
            }
        }

        private static OfficeStrokeDashStyle MapNativeDrawingPresetDash(A.Outline? outline) {
            A.PresetDash? presetDash = outline?.GetFirstChild<A.PresetDash>();
            string? value = presetDash?.GetAttribute("val", string.Empty).Value;
            if (string.IsNullOrWhiteSpace(value)) {
                return OfficeStrokeDashStyle.Solid;
            }

            if (value!.IndexOf("Dot", StringComparison.OrdinalIgnoreCase) >= 0 &&
                value.IndexOf("Dash", StringComparison.OrdinalIgnoreCase) >= 0) {
                return OfficeStrokeDashStyle.DashDot;
            }

            if (value.IndexOf("Dot", StringComparison.OrdinalIgnoreCase) >= 0) {
                return OfficeStrokeDashStyle.Dot;
            }

            return value.IndexOf("Dash", StringComparison.OrdinalIgnoreCase) >= 0
                ? OfficeStrokeDashStyle.Dash
                : OfficeStrokeDashStyle.Solid;
        }

        private static (double X, double Y) ParseNativeShapePoint(string value) {
            string[] parts = value.Split(',');
            if (parts.Length != 2) {
                return (0D, 0D);
            }

            return (ParseNativeShapePointPart(parts[0]), ParseNativeShapePointPart(parts[1]));
        }

        private static double ParseNativeShapePointPart(string value) {
            string normalized = value.Trim();
            if (string.IsNullOrWhiteSpace(normalized)) {
                return 0D;
            }

            double? resolved = ResolveNativeVmlLength(normalized, 0D, 0D);
            if (resolved.HasValue) {
                return resolved.Value;
            }

            return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) ? result : 0D;
        }

    }
}
