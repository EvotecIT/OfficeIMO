using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static bool AddShape(WordShape shape, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics) {
            if (shape._drawing == null) {
                return AddVmlShape(shape, context, diagnostics);
            }

            if (shape._wpsShape == null) {
                AddDiagnostic(diagnostics, "unsupported-word-shape", "Skipped a Word shape because its DrawingML wordprocessing shape could not be resolved.", DescribeShape(shape));
                return false;
            }

            DW.Anchor? anchor = shape._drawing.GetFirstChild<DW.Anchor>();
            if (anchor != null) {
                return AddAnchoredShape(shape, anchor, context, diagnostics);
            }

            if (!TryGetShapeSize(shape, anchor: null, out double width, out double height)) {
                AddDiagnostic(diagnostics, "unsupported-word-shape", "Skipped a Word shape because its inline DrawingML size could not be resolved.", DescribeShape(shape));
                return false;
            }

            width = Math.Min(width, context.ContentWidth);
            if (!EnsureVerticalSpace(context, height, diagnostics)) {
                return false;
            }

            string? presetName = GetShapePresetName(shape);
            if (!OfficeShapePresets.TryCreate(presetName, width, height, out OfficeShape? drawingShape) || drawingShape == null) {
                AddDiagnostic(diagnostics, "unsupported-word-shape", "Skipped a Word shape because its DrawingML preset is not yet projected through OfficeIMO.Drawing.", DescribeShape(shape));
                return false;
            }

            ApplyShapeStyle(drawingShape, shape);
            WordShapeFrameTransform transform = GetShapeFrameTransform(shape);
            if (transform.HasTransform) {
                drawingShape.Transform = CreateLocalShapeFrameTransform(width, height, transform);
            }

            context.Drawing.AddShape(drawingShape, context.Left, context.Y);
            context.Y += height + ParagraphGapPoints;
            return true;
        }

        private static bool AddVmlShape(WordShape shape, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics) {
            if (!TryCreateVmlShape(shape, out OfficeShape? drawingShape, out double leftOffset, out double topOffset) || drawingShape == null) {
                AddDiagnostic(diagnostics, "unsupported-word-shape", "Skipped a Word VML shape because its geometry or size could not be projected through OfficeIMO.Drawing.", DescribeShape(shape));
                return false;
            }

            if (drawingShape.Width <= 0D || drawingShape.Height <= 0D || !IsFinite(drawingShape.Width) || !IsFinite(drawingShape.Height)) {
                AddDiagnostic(diagnostics, "unsupported-word-shape", "Skipped a Word VML shape because its projected size is invalid.", DescribeShape(shape));
                return false;
            }

            if (drawingShape.Width + leftOffset > context.ContentWidth) {
                AddDiagnostic(diagnostics, "unsupported-word-shape", "Skipped a Word VML shape because it projects outside the current content width.", DescribeShape(shape));
                return false;
            }

            if (!EnsureVerticalSpace(context, drawingShape.Height + topOffset, diagnostics)) {
                return false;
            }

            ApplyShapeStyle(drawingShape, shape);
            WordShapeFrameTransform transform = GetShapeFrameTransform(shape);
            if (transform.HasTransform) {
                drawingShape.Transform = CreateLocalShapeFrameTransform(drawingShape.Width, drawingShape.Height, transform);
            }

            context.Drawing.AddShape(drawingShape, context.Left + leftOffset, context.Y + topOffset);
            context.Y += topOffset + drawingShape.Height + ParagraphGapPoints;
            return true;
        }

        private static bool AddAnchoredShape(WordShape shape, DW.Anchor anchor, WordImageFlowContext context, List<OfficeImageExportDiagnostic> diagnostics) {
            if (anchor.GetFirstChild<DW.WrapSquare>() == null) {
                AddDiagnostic(diagnostics, "unsupported-word-shape", "Skipped a Word floating shape because only square-wrapped DrawingML shapes are currently projected through OfficeIMO.Drawing.", DescribeShape(shape));
                return false;
            }

            if (!TryGetShapeSize(shape, anchor, out double width, out double height)) {
                AddDiagnostic(diagnostics, "unsupported-word-shape", "Skipped a Word floating shape because its DrawingML size could not be resolved.", DescribeShape(shape));
                return false;
            }

            string? presetName = GetShapePresetName(shape);
            if (!OfficeShapePresets.TryCreate(presetName, width, height, out OfficeShape? drawingShape) || drawingShape == null) {
                AddDiagnostic(diagnostics, "unsupported-word-shape", "Skipped a Word floating shape because its DrawingML preset is not yet projected through OfficeIMO.Drawing.", DescribeShape(shape));
                return false;
            }

            double left = ResolveHorizontalAnchorPosition(anchor.HorizontalPosition, context, width);
            double top = ResolveVerticalAnchorPosition(anchor.VerticalPosition, context, height);
            if (!IsFinite(left) || !IsFinite(top)) {
                AddDiagnostic(diagnostics, "unsupported-word-shape", "Skipped a Word floating shape because its anchor position could not be resolved.", DescribeShape(shape));
                return false;
            }

            double right = left + width;
            double bottom = top + height;
            if (left < 0D || top < 0D || right > context.Drawing.Width || bottom > context.Drawing.Height) {
                AddDiagnostic(diagnostics, "unsupported-word-shape", "Skipped a Word floating shape because its anchor projects outside the current page preview.", DescribeShape(shape));
                return false;
            }

            ApplyShapeStyle(drawingShape, shape);
            WordShapeFrameTransform transform = GetShapeFrameTransform(shape);
            if (transform.HasTransform) {
                drawingShape.Transform = CreateLocalShapeFrameTransform(width, height, transform);
            }

            context.Drawing.AddShape(drawingShape, left, top);
            context.AddTextExclusion(
                Math.Max(context.Left, left - GetAnchorDistancePoints(anchor.DistanceFromLeft)),
                Math.Max(0D, top - GetAnchorDistancePoints(anchor.DistanceFromTop)),
                Math.Min(context.Left + context.ContentWidth, right + GetAnchorDistancePoints(anchor.DistanceFromRight)),
                Math.Min(context.ContentBottom, bottom + GetAnchorDistancePoints(anchor.DistanceFromBottom)),
                GetShapeWrapSide(anchor));
            return true;
        }

        private static bool TryGetShapeSize(WordShape shape, DW.Anchor? anchor, out double width, out double height) {
            width = 0D;
            height = 0D;

            DW.Extent? extent = anchor?.Extent ?? shape._drawing?.GetFirstChild<DW.Inline>()?.Extent;
            if (extent?.Cx?.Value is long cx && extent.Cy?.Value is long cy && cx > 0L && cy > 0L) {
                width = Helpers.ConvertEmusToPoints(cx);
                height = Helpers.ConvertEmusToPoints(cy);
                return width > 0D && height > 0D;
            }

            A.Extents? transformExtents = shape._wpsShape?
                .GetFirstChild<Wps.ShapeProperties>()?
                .GetFirstChild<A.Transform2D>()?
                .GetFirstChild<A.Extents>();
            if (transformExtents?.Cx?.Value is long transformCx && transformExtents.Cy?.Value is long transformCy && transformCx > 0L && transformCy > 0L) {
                width = Helpers.ConvertEmusToPoints(transformCx);
                height = Helpers.ConvertEmusToPoints(transformCy);
                return width > 0D && height > 0D;
            }

            return false;
        }

        private static bool TryCreateVmlShape(WordShape shape, out OfficeShape? drawingShape, out double leftOffset, out double topOffset) {
            drawingShape = null;
            leftOffset = 0D;
            topOffset = 0D;

            if (shape._rectangle != null) {
                if (!TryGetVmlBoxSize(shape, out double width, out double height)) return false;
                drawingShape = OfficeShape.Rectangle(width, height);
                return true;
            }

            if (shape._ellipse != null) {
                if (!TryGetVmlBoxSize(shape, out double width, out double height)) return false;
                drawingShape = OfficeShape.Ellipse(width, height);
                return true;
            }

            if (shape._roundRectangle != null) {
                if (!TryGetVmlBoxSize(shape, out double width, out double height)) return false;
                double arcFraction = shape.ArcSize ?? 0.25D;
                arcFraction = Math.Max(0D, Math.Min(0.5D, arcFraction));
                drawingShape = OfficeShape.RoundedRectangle(width, height, Math.Min(width, height) * arcFraction);
                return true;
            }

            if (shape._line != null) {
                if (!TryParseVmlPointPair(shape._line.From?.Value, out OfficePoint start) ||
                    !TryParseVmlPointPair(shape._line.To?.Value, out OfficePoint end) ||
                    start == end) {
                    return false;
                }

                double minX = Math.Min(start.X, end.X);
                double minY = Math.Min(start.Y, end.Y);
                drawingShape = OfficeShape.Line(start, end);
                leftOffset = minX;
                topOffset = minY;
                return true;
            }

            if (shape._polygon?.Points?.Value is string pointsValue && TryParseVmlPoints(pointsValue, out List<OfficePoint> points)) {
                double minX = points.Min(point => point.X);
                double minY = points.Min(point => point.Y);
                drawingShape = OfficeShape.Polygon(points);
                leftOffset = minX;
                topOffset = minY;
                return true;
            }

            return false;
        }

        private static bool TryGetVmlBoxSize(WordShape shape, out double width, out double height) {
            width = shape.Width;
            height = shape.Height;
            return width > 0D && height > 0D && IsFinite(width) && IsFinite(height);
        }

        private static bool TryParseVmlPoints(string value, out List<OfficePoint> points) {
            points = new List<OfficePoint>();
            string[] tokens = value.Split(new[] { ' ', ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string token in tokens) {
                if (!TryParseVmlPointPair(token, out OfficePoint point)) {
                    return false;
                }

                points.Add(point);
            }

            return points.Count >= 3;
        }

        private static bool TryParseVmlPointPair(string? value, out OfficePoint point) {
            point = default;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }

            string text = value!;
            string[] parts = text.Split(',');
            if (parts.Length != 2 ||
                !TryParseVmlCoordinate(parts[0], out double x) ||
                !TryParseVmlCoordinate(parts[1], out double y)) {
                return false;
            }

            point = new OfficePoint(x, y);
            return true;
        }

        private static bool TryParseVmlCoordinate(string value, out double coordinate) {
            coordinate = 0D;
            string text = value.Trim();
            if (text.EndsWith("pt", StringComparison.OrdinalIgnoreCase)) {
                text = text.Substring(0, text.Length - 2);
            }

            if (!double.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out coordinate)) {
                return false;
            }

            return IsFinite(coordinate);
        }

        private static WordTextWrapSide GetShapeWrapSide(DW.Anchor anchor) {
            DW.WrapTextValues? wrapValue = anchor.Elements<DW.WrapSquare>().FirstOrDefault()?.WrapText?.Value;
            if (wrapValue == DW.WrapTextValues.Left) {
                return WordTextWrapSide.Left;
            }

            if (wrapValue == DW.WrapTextValues.Right) {
                return WordTextWrapSide.Right;
            }

            return WordTextWrapSide.Largest;
        }

        private static string? GetShapePresetName(WordShape shape) =>
            shape._wpsShape?
                .GetFirstChild<Wps.ShapeProperties>()?
                .GetFirstChild<A.PresetGeometry>()?
                .Preset?
                .InnerText;

        private static void ApplyShapeStyle(OfficeShape drawingShape, WordShape source) {
            OfficeColor fill = source.FillColor;
            drawingShape.FillColor = fill.A == 0 ? null : fill;

            OfficeColor stroke = source.StrokeColor;
            if (source.Stroked == false || stroke.A == 0) {
                drawingShape.StrokeColor = null;
                drawingShape.StrokeWidth = 0D;
            } else {
                drawingShape.StrokeColor = stroke;
                drawingShape.StrokeWidth = Math.Max(0D, source.StrokeWeight ?? 1D);
            }
        }

        private static WordShapeFrameTransform GetShapeFrameTransform(WordShape shape) {
            A.Transform2D? drawingTransform = shape._wpsShape?
                .GetFirstChild<Wps.ShapeProperties>()?
                .GetFirstChild<A.Transform2D>();
            double rotation = drawingTransform?.Rotation?.Value is int rotationValue ? rotationValue / 60000D : 0D;
            bool flipHorizontal = drawingTransform?.HorizontalFlip?.Value == true;
            bool flipVertical = drawingTransform?.VerticalFlip?.Value == true;

            string? vmlStyle = GetVmlShapeStyle(shape);
            if (!string.IsNullOrWhiteSpace(vmlStyle)) {
                if (TryGetVmlStyleText(vmlStyle, "rotation", out string vmlRotation) &&
                    double.TryParse(vmlRotation, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsedRotation) &&
                    IsFinite(parsedRotation)) {
                    rotation = parsedRotation;
                }

                if (TryGetVmlStyleText(vmlStyle, "flip", out string vmlFlip)) {
                    flipHorizontal = vmlFlip.IndexOf('x') >= 0 || vmlFlip.IndexOf('X') >= 0;
                    flipVertical = vmlFlip.IndexOf('y') >= 0 || vmlFlip.IndexOf('Y') >= 0;
                }
            }

            return new WordShapeFrameTransform(rotation, flipHorizontal, flipVertical);
        }

        private static string? GetVmlShapeStyle(WordShape shape) =>
            shape._shape?.Style?.Value ??
            shape._rectangle?.Style?.Value ??
            shape._roundRectangle?.Style?.Value ??
            shape._ellipse?.Style?.Value ??
            shape._line?.Style?.Value ??
            shape._polygon?.Style?.Value;

        private static OfficeTransform CreateLocalShapeFrameTransform(double width, double height, WordShapeFrameTransform transform) {
            double centerX = width / 2D;
            double centerY = height / 2D;
            return OfficeTransform.Translate(-centerX, -centerY)
                .Then(OfficeTransform.Scale(transform.FlipHorizontal ? -1D : 1D, transform.FlipVertical ? -1D : 1D))
                .Then(OfficeTransform.RotateDegrees(transform.RotationDegrees))
                .Then(OfficeTransform.Translate(centerX, centerY));
        }

        private static string DescribeShape(WordShape shape) {
            string? presetName = GetShapePresetName(shape);
            if (!string.IsNullOrWhiteSpace(presetName)) {
                return "Word shape " + presetName;
            }

            if (shape._rectangle != null) return "Word rectangle";
            if (shape._roundRectangle != null) return "Word rounded rectangle";
            if (shape._ellipse != null) return "Word ellipse";
            if (shape._line != null) return "Word line";
            if (shape._polygon != null) return "Word polygon";
            return "Word shape";
        }

        private readonly struct WordShapeFrameTransform {
            internal WordShapeFrameTransform(double rotationDegrees, bool flipHorizontal, bool flipVertical) {
                RotationDegrees = rotationDegrees;
                FlipHorizontal = flipHorizontal;
                FlipVertical = flipVertical;
            }

            internal double RotationDegrees { get; }

            internal bool FlipHorizontal { get; }

            internal bool FlipVertical { get; }

            internal bool HasTransform => Math.Abs(RotationDegrees) > 0.000001D || FlipHorizontal || FlipVertical;
        }
    }
}
