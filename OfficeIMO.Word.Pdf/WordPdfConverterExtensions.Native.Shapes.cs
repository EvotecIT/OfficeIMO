using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;
using static OfficeIMO.Word.OpenXml.WordXml;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool RenderNativeShape(INativePdfFlow pdf, WordShape shape, double spacingAfter = 6D, NativeObjectParagraphSpacing? paragraphSpacing = null) {
            if (shape == null) {
                return false;
            }
            bool alignToLineTop = shape.DrawingNode().IsNone ||
                (shape.TryGetLayoutSnapshot(out WordDrawingLayoutInfo layout) && layout.Placement == WordDrawingPlacementKind.Anchored);
            // Hidden flow shapes retain their occupied line height while omitting paint.
            if (shape.Hidden == true && GetNativeShapeDimensions(shape) is { Height: > 0D } dimensions) {
                RenderNativeFlowObject(pdf, paragraphSpacing,
                    flow => flow.Spacer(dimensions.Height + spacingAfter), alignToLineTop);
                return true;
            }
            OfficeShape? nativeShape = CreateNativeShape(shape);
            if (nativeShape == null) {
                return false;
            }

            RenderNativeFlowObject(pdf, paragraphSpacing,
                flow => flow.Shape(nativeShape, global::OfficeIMO.OfficeAlignment.Left, spacingAfter: spacingAfter), alignToLineTop);
            return true;
        }

        private static OfficeShape? CreateNativeShape(WordShape shape) {
            if (shape == null || shape.Hidden == true) {
                return null;
            }

            OfficeShape? nativeShape;
            OfficeXmlNode line = shape.VmlLineNode();
            OfficeXmlNode polygon = shape.VmlPolygonNode();
            if (line.Exists) {
                (double x1, double y1) = ParseNativeShapePoint(line.Attr("from") ?? "0pt,0pt");
                (double x2, double y2) = ParseNativeShapePoint(line.Attr("to") ?? "0pt,0pt");
                nativeShape = OfficeShape.Line(x1, y1, x2, y2);
            } else if (polygon.Exists && TryCreateNativePolygonShape(polygon.Attr("points"), out nativeShape)) {
            } else {
                (double Width, double Height)? dimensions = GetNativeShapeDimensions(shape);
                if (!dimensions.HasValue) {
                    return null;
                }

                double width = dimensions.Value.Width;
                double height = dimensions.Value.Height;
                if (shape.VmlOvalNode().Exists) {
                    nativeShape = OfficeShape.Ellipse(width, height);
                } else if (shape.VmlRoundRectangleNode().Exists) {
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

            IReadOnlyDictionary<string, OfficeColor> themeColors = GetNativeDrawingThemeColors(GetNativeShapeSourcePart(shape));
            ApplyNativeShapeStyle(nativeShape, shape, themeColors);
            return nativeShape;
        }

        /// <summary>The root element of the part that holds the shape's run, or <see cref="OfficeXmlNode.None"/>.</summary>
        private static OfficeXmlNode GetNativeShapeSourcePart(WordShape shape) {
            OfficeXmlNode run = shape.RunNode();
            return run.IsNone ? default : run.PartRoot();
        }

        private static (double Width, double Height)? GetNativeShapeDimensions(WordShape shape) {
            double width = shape.WidthPoints;
            double height = shape.HeightPoints;
            if (width > 0 && height > 0) {
                return (width, height);
            }

            OfficeXmlNode shapeProperties = GetNativeWpsShapeProperties(shape.WpsShapeNode());
            OfficeXmlNode transform = shapeProperties.IsNone ? default : shapeProperties.Child(A, "xfrm");
            OfficeXmlNode extents = transform.IsNone ? default : transform.Child(A, "ext");

            long? cx = extents.Int64Attr("cx");
            long? cy = extents.Int64Attr("cy");
            if (!cx.HasValue || !cy.HasValue || cx.Value <= 0 || cy.Value <= 0) {
                return null;
            }

            return (ConvertNativeEmusToPoints(cx.Value), ConvertNativeEmusToPoints(cy.Value));
        }

        private static bool TryGetNativeDrawingPreset(WordShape shape, out string? presetName) {
            OfficeXmlNode shapeProperties = GetNativeWpsShapeProperties(shape.WpsShapeNode());
            OfficeXmlNode geometry = shapeProperties.IsNone ? default : shapeProperties.Child(A, "prstGeom");

            presetName = geometry.Attr("prst");
            return !string.IsNullOrWhiteSpace(presetName);
        }

        /// <summary>The <c>wps:spPr</c> of a <c>wps:wsp</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
        private static OfficeXmlNode GetNativeWpsShapeProperties(OfficeXmlNode wpsShape) =>
            wpsShape.IsNone ? default : wpsShape.Elements(Wps, "spPr").FirstOrDefault();

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
                    !OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out double x) ||
                    !OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double y)) {
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

        private static void ApplyNativeShapeStyle(OfficeShape nativeShape, WordShape wordShape, IReadOnlyDictionary<string, OfficeColor> themeColors) {
            OfficeXmlNode wpsShape = wordShape.WpsShapeNode();
            OfficeXmlNode shapeProperties = GetNativeWpsShapeProperties(wpsShape);
            if (nativeShape.Kind != OfficeShapeKind.Line) {
                if (HasNativeDrawingNoFill(shapeProperties)) {
                    nativeShape.FillColor = null;
                } else if (TryGetNativeDrawingGradientFill(shapeProperties, out OfficeLinearGradient? drawingGradient, themeColors)) {
                    nativeShape.FillGradient = drawingGradient;
                } else if (TryGetNativeDrawingSolidFillColor(shapeProperties, out OfficeColor drawingFill, themeColors)) {
                    nativeShape.FillColor = drawingFill;
                } else {
                    global::OfficeIMO.OfficeColor? fill = ParseNativeColor(wordShape.FillColorValue);
                    if (fill.HasValue) {
                        nativeShape.FillColor = fill.Value;
                    }
                }

                // DrawingML fill alpha is already carried by the solid color or
                // gradient stops. A separate FillOpacity would apply it twice.
            }

            OfficeXmlNode drawingOutline = shapeProperties.IsNone ? default : shapeProperties.Child(A, "ln");
            bool drawingOutlineNoFill = HasNativeDrawingNoFill(drawingOutline);
            bool hasDrawingStroke = drawingOutline.Exists && !drawingOutlineNoFill;
            OfficeColor drawingStroke = default;
            bool hasDrawingStrokeColor = hasDrawingStroke && TryGetNativeDrawingSolidFillColor(drawingOutline, out drawingStroke, themeColors);
            bool drawStroke = !drawingOutlineNoFill &&
                              (nativeShape.Kind == OfficeShapeKind.Line ||
                               wordShape.Stroked == true ||
                               hasDrawingStrokeColor ||
                               (wpsShape.Exists && !string.IsNullOrWhiteSpace(wordShape.StrokeColorValue)));
            if (!drawStroke) {
                nativeShape.StrokeColor = null;
                nativeShape.StrokeWidth = 0;
                return;
            }

            global::OfficeIMO.OfficeColor? stroke = hasDrawingStrokeColor ? drawingStroke : ParseNativeColor(wordShape.StrokeColorValue);
            OfficeColor strokeColor = stroke ?? OfficeColor.Black;
            // This native shape carries stroke alpha separately for running-canvas admission.
            var strokeRgb = strokeColor.ToNormalizedRgb();
            nativeShape.StrokeColor = OfficeColor.FromNormalizedRgb(strokeRgb.R, strokeRgb.G, strokeRgb.B);
            // Transfer the shared resolver's completed alpha once alongside RGB.
            // Running text-box admission also relies on this separate-opacity contract.
            if (hasDrawingStrokeColor) nativeShape.StrokeOpacity = drawingStroke.A / 255D;
            nativeShape.StrokeWidth = Math.Max(0D, wordShape.StrokeWeightPoints ?? 1D);
            nativeShape.StrokeDashStyle = MapNativeDrawingPresetDash(drawingOutline);
        }

        private static OfficeStrokeDashStyle MapNativeDrawingPresetDash(OfficeXmlNode outline) {
            OfficeXmlNode presetDash = outline.IsNone ? default : outline.Child(A, "prstDash");
            string? value = presetDash.Attr("val");
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

            return OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out double result) ? result : 0D;
        }

    }
}
