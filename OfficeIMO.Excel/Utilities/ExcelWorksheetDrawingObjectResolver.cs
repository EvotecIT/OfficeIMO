using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel.Utilities {
    internal static class ExcelWorksheetDrawingObjectResolver {
        private const double EmusPerPixel = 9525D;

        internal static IReadOnlyList<ExcelWorksheetDrawingObjectInfo> FindDrawingObjects(WorksheetPart worksheetPart) {
            if (worksheetPart == null) {
                throw new ArgumentNullException(nameof(worksheetPart));
            }

            Xdr.WorksheetDrawing? worksheetDrawing = worksheetPart.DrawingsPart?.WorksheetDrawing;
            if (worksheetDrawing == null) {
                return Array.Empty<ExcelWorksheetDrawingObjectInfo>();
            }

            var objects = new List<ExcelWorksheetDrawingObjectInfo>();
            for (int order = 0; order < worksheetDrawing.ChildElements.Count; order++) {
                OpenXmlElement anchor = worksheetDrawing.ChildElements[order];
                AnchorPosition position = GetAnchorPosition(anchor);
                foreach (OpenXmlElement element in anchor.ChildElements) {
                    if (element is Xdr.Shape shape) {
                        objects.Add(CreateShapeInfo(shape, position, order));
                    } else if (IsUnsupportedDrawingElement(element)) {
                        objects.Add(CreateUnsupportedInfo(element, position, order, null));
                    }
                }
            }

            return objects;
        }

        internal static IReadOnlyList<ExcelWorksheetDrawingObjectInfo> FindUnsupportedDrawingObjects(WorksheetPart worksheetPart) =>
            FindDrawingObjects(worksheetPart)
                .Where(drawing => !drawing.IsRenderable)
                .ToList();

        private static ExcelWorksheetDrawingObjectInfo CreateShapeInfo(Xdr.Shape shape, AnchorPosition position, int order) {
            string name = GetDrawingName(shape, "shape");
            A.Transform2D? transform = shape.ShapeProperties?.GetFirstChild<A.Transform2D>();
            if (TryGetRotationDegrees(transform, out double rotationDegrees) && Math.Abs(rotationDegrees) > 0.001D) {
                return CreateUnsupportedInfo(shape, position, order, "rotated shapes are not rendered yet");
            }

            if (!TryGetShapeKind(shape, out OfficeShapeKind shapeKind, out string? unsupportedReason)) {
                return CreateUnsupportedInfo(shape, position, order, unsupportedReason);
            }

            if (!TryGetFillColor(shape.ShapeProperties, out string? fillColorArgb, out unsupportedReason)) {
                return CreateUnsupportedInfo(shape, position, order, unsupportedReason);
            }

            if (!TryGetStroke(shape.ShapeProperties, out string? strokeColorArgb, out double strokeWidth, out unsupportedReason)) {
                return CreateUnsupportedInfo(shape, position, order, unsupportedReason);
            }

            string text = string.Join(Environment.NewLine, shape.TextBody?
                .Elements<A.Paragraph>()
                .Select(paragraph => string.Concat(paragraph.Descendants<A.Text>().Select(item => item.Text)))
                .Where(line => !string.IsNullOrEmpty(line)) ?? Enumerable.Empty<string>());

            return new ExcelWorksheetDrawingObjectInfo(
                name,
                "shape",
                order,
                position.Row,
                position.Column,
                position.OffsetXPixels,
                position.OffsetYPixels,
                position.WidthPixels,
                position.HeightPixels,
                position.ToColumn,
                position.ToRow,
                position.ToOffsetXPixels,
                position.ToOffsetYPixels,
                shapeKind,
                fillColorArgb,
                strokeColorArgb,
                strokeWidth,
                text,
                unsupportedReason: null);
        }

        private static ExcelWorksheetDrawingObjectInfo CreateUnsupportedInfo(OpenXmlElement element, AnchorPosition position, int order, string? unsupportedReason) {
            string kind = GetDrawingElementDisplayName(element);
            string name = GetDrawingName(element, "unnamed " + kind);
            return new ExcelWorksheetDrawingObjectInfo(
                name,
                kind,
                order,
                position.Row,
                position.Column,
                position.OffsetXPixels,
                position.OffsetYPixels,
                position.WidthPixels,
                position.HeightPixels,
                position.ToColumn,
                position.ToRow,
                position.ToOffsetXPixels,
                position.ToOffsetYPixels,
                shapeKind: null,
                fillColorArgb: null,
                strokeColorArgb: null,
                strokeWidth: 0D,
                text: string.Empty,
                unsupportedReason: unsupportedReason);
        }

        private static AnchorPosition GetAnchorPosition(OpenXmlElement anchor) {
            Xdr.MarkerType? fromMarker = anchor switch {
                Xdr.OneCellAnchor oneCellAnchor => oneCellAnchor.FromMarker,
                Xdr.TwoCellAnchor twoCellAnchor => twoCellAnchor.FromMarker,
                _ => null,
            };
            Xdr.MarkerType? toMarker = anchor is Xdr.TwoCellAnchor twoCell ? twoCell.ToMarker : null;
            Xdr.Extent? extent = anchor is Xdr.OneCellAnchor oneCell ? oneCell.Extent : null;

            int row = ParseOneBasedMarker(fromMarker?.RowId?.Text);
            int column = ParseOneBasedMarker(fromMarker?.ColumnId?.Text);
            int offsetX = ParseEmuPixels(fromMarker?.ColumnOffset?.Text);
            int offsetY = ParseEmuPixels(fromMarker?.RowOffset?.Text);
            int? toRow = ParseOneBasedMarkerOrNull(toMarker?.RowId?.Text);
            int? toColumn = ParseOneBasedMarkerOrNull(toMarker?.ColumnId?.Text);
            int toOffsetX = ParseEmuPixels(toMarker?.ColumnOffset?.Text);
            int toOffsetY = ParseEmuPixels(toMarker?.RowOffset?.Text);
            int width = ParseEmuPixels(extent?.Cx?.Value);
            int height = ParseEmuPixels(extent?.Cy?.Value);

            return new AnchorPosition(row, column, offsetX, offsetY, width, height, toColumn, toRow, toOffsetX, toOffsetY);
        }

        private static bool TryGetShapeKind(Xdr.Shape shape, out OfficeShapeKind shapeKind, out string? unsupportedReason) {
            string preset = shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText ?? string.Empty;
            if (string.Equals(preset, "rect", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(preset, "rectangle", StringComparison.OrdinalIgnoreCase)) {
                shapeKind = OfficeShapeKind.Rectangle;
                unsupportedReason = null;
                return true;
            }

            if (string.Equals(preset, "roundRect", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(preset, "roundRectangle", StringComparison.OrdinalIgnoreCase)) {
                shapeKind = OfficeShapeKind.RoundedRectangle;
                unsupportedReason = null;
                return true;
            }

            shapeKind = OfficeShapeKind.Rectangle;
            unsupportedReason = string.IsNullOrWhiteSpace(preset)
                ? "shape geometry is missing"
                : "shape geometry '" + preset + "' is not rendered yet";
            return false;
        }

        private static bool TryGetFillColor(OpenXmlCompositeElement? properties, out string? fillColorArgb, out string? unsupportedReason) {
            fillColorArgb = null;
            unsupportedReason = null;
            if (properties == null || properties.GetFirstChild<A.NoFill>() != null) {
                return true;
            }

            A.SolidFill? solidFill = properties.GetFirstChild<A.SolidFill>();
            if (solidFill == null) {
                unsupportedReason = "shape fill is not a supported solid RGB fill";
                return false;
            }

            fillColorArgb = NormalizeRgb(solidFill.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
            if (fillColorArgb != null) {
                return true;
            }

            unsupportedReason = "shape fill uses a theme, system, or transformed color that is not rendered yet";
            return false;
        }

        private static bool TryGetStroke(OpenXmlCompositeElement? properties, out string? strokeColorArgb, out double strokeWidth, out string? unsupportedReason) {
            strokeColorArgb = null;
            strokeWidth = 1D;
            unsupportedReason = null;
            A.Outline? outline = properties?.GetFirstChild<A.Outline>();
            if (outline == null || outline.GetFirstChild<A.NoFill>() != null) {
                strokeWidth = 0D;
                return true;
            }

            A.SolidFill? solidFill = outline.GetFirstChild<A.SolidFill>();
            if (solidFill == null) {
                unsupportedReason = "shape outline is not a supported solid RGB line";
                return false;
            }

            strokeColorArgb = NormalizeRgb(solidFill.GetFirstChild<A.RgbColorModelHex>()?.Val?.Value);
            if (strokeColorArgb == null) {
                unsupportedReason = "shape outline uses a theme, system, or transformed color that is not rendered yet";
                return false;
            }

            if (outline.Width != null && outline.Width.Value > 0) {
                strokeWidth = Math.Max(1D, outline.Width.Value / EmusPerPixel);
            }

            return true;
        }

        private static bool TryGetRotationDegrees(A.Transform2D? transform, out double rotationDegrees) {
            rotationDegrees = 0D;
            if (transform == null) {
                return false;
            }

            string? raw = transform.GetAttribute("rot", string.Empty).Value;
            if (!long.TryParse(raw, out long rotation)) {
                return false;
            }

            rotationDegrees = rotation / 60000D;
            return true;
        }

        private static bool IsUnsupportedDrawingElement(OpenXmlElement element) {
            if (!string.Equals(element.NamespaceUri, "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", StringComparison.Ordinal)) {
                return false;
            }

            switch (element.LocalName) {
                case "cxnSp":
                case "grpSp":
                    return true;
                case "graphicFrame":
                    return !element.Descendants<C.ChartReference>().Any();
                default:
                    return false;
            }
        }

        private static string GetDrawingName(OpenXmlElement element, string fallback) =>
            element.Descendants<Xdr.NonVisualDrawingProperties>()
                .FirstOrDefault()?.Name?.Value
            ?? fallback;

        private static string GetDrawingElementDisplayName(OpenXmlElement element) {
            switch (element.LocalName) {
                case "sp":
                    return "shape";
                case "cxnSp":
                    return "connector";
                case "grpSp":
                    return "group shape";
                case "graphicFrame":
                    return "graphic frame";
                default:
                    return element.LocalName;
            }
        }

        private static string? NormalizeRgb(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            string normalized = value!.Trim().TrimStart('#');
            if (normalized.Length == 6) {
                return "FF" + normalized.ToUpperInvariant();
            }

            return normalized.Length == 8 ? normalized.ToUpperInvariant() : null;
        }

        private static int ParseOneBasedMarker(string? value) =>
            int.TryParse(value, out int zeroBased) && zeroBased >= 0 ? zeroBased + 1 : 0;

        private static int? ParseOneBasedMarkerOrNull(string? value) {
            int parsed = ParseOneBasedMarker(value);
            return parsed > 0 ? parsed : null;
        }

        private static int ParseEmuPixels(string? value) =>
            long.TryParse(value, out long emus) ? Math.Max(0, (int)Math.Round(emus / EmusPerPixel)) : 0;

        private static int ParseEmuPixels(long? value) =>
            value.HasValue ? Math.Max(0, (int)Math.Round(value.Value / EmusPerPixel)) : 0;

        private readonly struct AnchorPosition {
            internal AnchorPosition(
                int row,
                int column,
                int offsetXPixels,
                int offsetYPixels,
                int widthPixels,
                int heightPixels,
                int? toColumn,
                int? toRow,
                int toOffsetXPixels,
                int toOffsetYPixels) {
                Row = row;
                Column = column;
                OffsetXPixels = offsetXPixels;
                OffsetYPixels = offsetYPixels;
                WidthPixels = widthPixels;
                HeightPixels = heightPixels;
                ToColumn = toColumn;
                ToRow = toRow;
                ToOffsetXPixels = toOffsetXPixels;
                ToOffsetYPixels = toOffsetYPixels;
            }

            internal int Row { get; }

            internal int Column { get; }

            internal int OffsetXPixels { get; }

            internal int OffsetYPixels { get; }

            internal int WidthPixels { get; }

            internal int HeightPixels { get; }

            internal int? ToColumn { get; }

            internal int? ToRow { get; }

            internal int ToOffsetXPixels { get; }

            internal int ToOffsetYPixels { get; }
        }
    }

    internal sealed class ExcelWorksheetDrawingObjectInfo {
        internal ExcelWorksheetDrawingObjectInfo(
            string name,
            string kind,
            int order,
            int row,
            int column,
            int offsetXPixels,
            int offsetYPixels,
            int widthPixels,
            int heightPixels,
            int? toColumn,
            int? toRow,
            int toOffsetXPixels,
            int toOffsetYPixels,
            OfficeShapeKind? shapeKind,
            string? fillColorArgb,
            string? strokeColorArgb,
            double strokeWidth,
            string text,
            string? unsupportedReason) {
            Name = name ?? string.Empty;
            Kind = kind ?? string.Empty;
            Order = order;
            Row = row;
            Column = column;
            OffsetXPixels = offsetXPixels;
            OffsetYPixels = offsetYPixels;
            WidthPixels = widthPixels;
            HeightPixels = heightPixels;
            ToColumn = toColumn;
            ToRow = toRow;
            ToOffsetXPixels = toOffsetXPixels;
            ToOffsetYPixels = toOffsetYPixels;
            ShapeKind = shapeKind;
            FillColorArgb = fillColorArgb;
            StrokeColorArgb = strokeColorArgb;
            StrokeWidth = strokeWidth;
            Text = text ?? string.Empty;
            UnsupportedReason = unsupportedReason;
        }

        internal string Name { get; }

        internal string Kind { get; }

        internal int Order { get; }

        internal int Row { get; }

        internal int Column { get; }

        internal int OffsetXPixels { get; }

        internal int OffsetYPixels { get; }

        internal int WidthPixels { get; }

        internal int HeightPixels { get; }

        internal int? ToColumn { get; }

        internal int? ToRow { get; }

        internal int ToOffsetXPixels { get; }

        internal int ToOffsetYPixels { get; }

        internal OfficeShapeKind? ShapeKind { get; }

        internal string? FillColorArgb { get; }

        internal string? StrokeColorArgb { get; }

        internal double StrokeWidth { get; }

        internal string Text { get; }

        internal string? UnsupportedReason { get; }

        internal bool IsRenderable => ShapeKind.HasValue && string.IsNullOrEmpty(UnsupportedReason);

        internal string? CellReference => Row > 0 && Column > 0 ? A1.CellReference(Row, Column) : null;
    }
}
