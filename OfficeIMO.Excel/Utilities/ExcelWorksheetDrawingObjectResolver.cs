using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel.Utilities {
    internal static class ExcelWorksheetDrawingObjectResolver {
        private const double EmusPerPixel = 9525D;
        private const long DefaultLeftRightTextInsetEmu = 91440L;
        private const long DefaultTopBottomTextInsetEmu = 45720L;

        internal static IReadOnlyList<ExcelWorksheetDrawingObjectInfo> FindDrawingObjects(WorksheetPart worksheetPart) {
            if (worksheetPart == null) {
                throw new ArgumentNullException(nameof(worksheetPart));
            }

            Xdr.WorksheetDrawing? worksheetDrawing = worksheetPart.DrawingsPart?.WorksheetDrawing;
            if (worksheetDrawing == null) {
                return Array.Empty<ExcelWorksheetDrawingObjectInfo>();
            }

            WorkbookPart? workbookPart = worksheetPart.GetParentParts().OfType<WorkbookPart>().FirstOrDefault();
            ExcelWorksheetGeometryIndex geometry = ExcelWorksheetGeometryIndex.Create(worksheetPart);
            var objects = new List<ExcelWorksheetDrawingObjectInfo>();
            for (int order = 0; order < worksheetDrawing.ChildElements.Count; order++) {
                OpenXmlElement anchor = worksheetDrawing.ChildElements[order];
                AnchorPosition position = GetAnchorPosition(anchor, geometry);
                foreach (OpenXmlElement element in anchor.ChildElements) {
                    if (element is Xdr.Shape shape) {
                        objects.Add(CreateShapeInfo(shape, position, order, workbookPart));
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

        private static ExcelWorksheetDrawingObjectInfo CreateShapeInfo(Xdr.Shape shape, AnchorPosition position, int order, WorkbookPart? workbookPart) {
            string name = GetDrawingName(shape, "shape");
            A.Transform2D? transform = shape.ShapeProperties?.GetFirstChild<A.Transform2D>();
            TryGetRotationDegrees(transform, out double rotationDegrees);

            if (!TryGetShapePreset(shape, out string shapePresetName, out OfficeShapeKind shapeKind, out string? unsupportedReason)) {
                return CreateUnsupportedInfo(shape, position, order, unsupportedReason);
            }

            if (!TryGetFillColor(shape.ShapeProperties, workbookPart, out string? fillColorArgb, out unsupportedReason)) {
                return CreateUnsupportedInfo(shape, position, order, unsupportedReason);
            }

            if (!TryGetStroke(
                shape.ShapeProperties,
                workbookPart,
                out string? strokeColorArgb,
                out double strokeWidth,
                out OfficeStrokeDashStyle strokeDashStyle,
                out OfficeStrokeLineCap? strokeLineCap,
                out OfficeStrokeLineJoin? strokeLineJoin,
                out unsupportedReason)) {
                return CreateUnsupportedInfo(shape, position, order, unsupportedReason);
            }

            string text = string.Join(Environment.NewLine, shape.TextBody?
                .Elements<A.Paragraph>()
                .Select(GetDrawingParagraphText)
                .Where(line => !string.IsNullOrEmpty(line)) ?? Enumerable.Empty<string>());
            OfficeTextAlignment textAlignment = ResolveTextAlignment(shape.TextBody);
            A.BodyProperties? bodyProperties = shape.TextBody?.GetFirstChild<A.BodyProperties>();
            OfficeTextVerticalAlignment textVerticalAlignment = ResolveTextVerticalAlignment(bodyProperties);
            bool textWrap = ResolveTextWrap(bodyProperties);
            bool textShrinkToFit = ResolveTextShrinkToFit(bodyProperties);
            bool textResizeShapeToFit = ResolveTextResizeShapeToFit(bodyProperties);
            ExcelDrawingTextOrientation textOrientation = ResolveTextOrientation(bodyProperties);
            DrawingTextInsets textInsets = ResolveTextInsets(bodyProperties);
            DrawingTextStyle textStyle = ResolveTextStyle(shape.TextBody, workbookPart);
            A.EffectList? effects = shape.ShapeProperties?.GetFirstChild<A.EffectList>();
            TryCreateGlow(effects?.GetFirstChild<A.Glow>(), workbookPart, out OfficeGlow? glow);
            TryCreateShadow(effects?.GetFirstChild<A.OuterShadow>(), workbookPart, out OfficeShadow? shadow);

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
                shapePresetName,
                shapeKind,
                ReadBooleanAttribute(transform, "flipH"),
                ReadBooleanAttribute(transform, "flipV"),
                rotationDegrees,
                fillColorArgb,
                strokeColorArgb,
                strokeWidth,
                strokeDashStyle,
                strokeLineCap,
                strokeLineJoin,
                text,
                textAlignment,
                textVerticalAlignment,
                textStyle.ColorArgb,
                textStyle.FontFamily,
                textStyle.FontSize,
                textStyle.FontStyle,
                textWrap,
                textShrinkToFit,
                textResizeShapeToFit,
                textOrientation,
                textInsets.Left,
                textInsets.Top,
                textInsets.Right,
                textInsets.Bottom,
                glow,
                shadow,
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
                shapePresetName: string.Empty,
                shapeKind: null,
                horizontalFlip: false,
                verticalFlip: false,
                rotationDegrees: 0D,
                fillColorArgb: null,
                strokeColorArgb: null,
                strokeWidth: 0D,
                strokeDashStyle: OfficeStrokeDashStyle.Solid,
                strokeLineCap: null,
                strokeLineJoin: null,
                text: string.Empty,
                textAlignment: OfficeTextAlignment.Center,
                textVerticalAlignment: OfficeTextVerticalAlignment.Center,
                textColorArgb: null,
                textFontFamily: null,
                textFontSize: null,
                textFontStyle: OfficeFontStyle.Regular,
                textWrap: false,
                textShrinkToFit: false,
                textResizeShapeToFit: false,
                textOrientation: ExcelDrawingTextOrientation.Horizontal,
                textInsetLeft: 0D,
                textInsetTop: 0D,
                textInsetRight: 0D,
                textInsetBottom: 0D,
                glow: null,
                shadow: null,
                unsupportedReason: unsupportedReason);
        }

        private static OfficeTextAlignment ResolveTextAlignment(Xdr.TextBody? textBody) {
            string? alignment = textBody?
                .Elements<A.Paragraph>()
                .Select(paragraph => GetRawAttribute(paragraph.GetFirstChild<A.ParagraphProperties>(), "algn"))
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            if (string.Equals(alignment, "r", StringComparison.OrdinalIgnoreCase)) {
                return OfficeTextAlignment.Right;
            }

            if (string.Equals(alignment, "l", StringComparison.OrdinalIgnoreCase)) {
                return OfficeTextAlignment.Left;
            }

            return OfficeTextAlignment.Center;
        }

        private static OfficeTextVerticalAlignment ResolveTextVerticalAlignment(A.BodyProperties? bodyProperties) {
            string? anchor = GetRawAttribute(bodyProperties, "anchor");
            if (string.Equals(anchor, "t", StringComparison.OrdinalIgnoreCase)) {
                return OfficeTextVerticalAlignment.Top;
            }

            if (string.Equals(anchor, "b", StringComparison.OrdinalIgnoreCase)) {
                return OfficeTextVerticalAlignment.Bottom;
            }

            return OfficeTextVerticalAlignment.Center;
        }

        private static bool ResolveTextWrap(A.BodyProperties? bodyProperties) {
            if (bodyProperties == null) {
                return false;
            }

            string? wrap = GetRawAttribute(bodyProperties, "wrap");
            return !string.Equals(wrap, "none", StringComparison.OrdinalIgnoreCase);
        }

        private static bool ResolveTextShrinkToFit(A.BodyProperties? bodyProperties) =>
            bodyProperties?.GetFirstChild<A.NormalAutoFit>() != null;

        private static bool ResolveTextResizeShapeToFit(A.BodyProperties? bodyProperties) =>
            bodyProperties?.GetFirstChild<A.ShapeAutoFit>() != null;

        private static ExcelDrawingTextOrientation ResolveTextOrientation(A.BodyProperties? bodyProperties) {
            string? vertical = GetRawAttribute(bodyProperties, "vert");
            if (string.IsNullOrWhiteSpace(vertical) || string.Equals(vertical, "horz", StringComparison.OrdinalIgnoreCase)) {
                return ExcelDrawingTextOrientation.Horizontal;
            }

            if (string.Equals(vertical, "vert", StringComparison.OrdinalIgnoreCase)) {
                return ExcelDrawingTextOrientation.Vertical;
            }

            if (string.Equals(vertical, "vert270", StringComparison.OrdinalIgnoreCase)) {
                return ExcelDrawingTextOrientation.Vertical270;
            }

            if (string.Equals(vertical, "eaVert", StringComparison.OrdinalIgnoreCase)) {
                return ExcelDrawingTextOrientation.EastAsianVertical;
            }

            if (string.Equals(vertical, "mongolianVert", StringComparison.OrdinalIgnoreCase)) {
                return ExcelDrawingTextOrientation.MongolianVertical;
            }

            if (string.Equals(vertical, "wordArtVert", StringComparison.OrdinalIgnoreCase)) {
                return ExcelDrawingTextOrientation.WordArtVertical;
            }

            if (string.Equals(vertical, "wordArtVertRtl", StringComparison.OrdinalIgnoreCase)) {
                return ExcelDrawingTextOrientation.WordArtLeftToRight;
            }

            return ExcelDrawingTextOrientation.Unknown;
        }

        private static DrawingTextInsets ResolveTextInsets(A.BodyProperties? bodyProperties) {
            if (bodyProperties == null) {
                return DrawingTextInsets.None;
            }

            return new DrawingTextInsets(
                ParseEmuPixels(GetRawAttribute(bodyProperties, "lIns"), DefaultLeftRightTextInsetEmu),
                ParseEmuPixels(GetRawAttribute(bodyProperties, "tIns"), DefaultTopBottomTextInsetEmu),
                ParseEmuPixels(GetRawAttribute(bodyProperties, "rIns"), DefaultLeftRightTextInsetEmu),
                ParseEmuPixels(GetRawAttribute(bodyProperties, "bIns"), DefaultTopBottomTextInsetEmu));
        }

        private static DrawingTextStyle ResolveTextStyle(Xdr.TextBody? textBody, WorkbookPart? workbookPart) {
            A.RunProperties? runProperties = textBody?
                .Descendants<A.RunProperties>()
                .FirstOrDefault();
            if (runProperties == null) {
                return DrawingTextStyle.Default;
            }

            string? colorArgb = ExcelThemeColorResolver.Resolve(runProperties.GetFirstChild<A.SolidFill>(), workbookPart);
            string? fontFamily = NormalizeFontFamily(GetRawAttribute(runProperties.GetFirstChild<A.LatinFont>(), "typeface"));
            double? fontSize = int.TryParse(GetRawAttribute(runProperties, "sz"), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int rawFontSize) && rawFontSize > 0
                ? rawFontSize / 100D
                : null;
            OfficeFontStyle fontStyle = OfficeFontStyle.Regular;
            if (ReadBooleanAttribute(runProperties, "b")) {
                fontStyle |= OfficeFontStyle.Bold;
            }

            if (ReadBooleanAttribute(runProperties, "i")) {
                fontStyle |= OfficeFontStyle.Italic;
            }

            string? underline = GetRawAttribute(runProperties, "u");
            if (!string.IsNullOrWhiteSpace(underline) && !string.Equals(underline, "none", StringComparison.OrdinalIgnoreCase)) {
                fontStyle |= OfficeFontStyle.Underline;
            }

            string? strike = GetRawAttribute(runProperties, "strike");
            if (!string.IsNullOrWhiteSpace(strike) && !string.Equals(strike, "noStrike", StringComparison.OrdinalIgnoreCase)) {
                fontStyle |= OfficeFontStyle.Strikethrough;
            }

            return new DrawingTextStyle(colorArgb, fontFamily, fontSize, fontStyle);
        }

        private static AnchorPosition GetAnchorPosition(OpenXmlElement anchor, ExcelWorksheetGeometryIndex geometry) {
            if (anchor is Xdr.AbsoluteAnchor absoluteAnchor) {
                int absoluteX = ExcelImageExportLimits.ClampAbsoluteOffsetPixels(ParseEmuPixels(absoluteAnchor.Position?.X?.Value));
                int absoluteY = ExcelImageExportLimits.ClampAbsoluteOffsetPixels(ParseEmuPixels(absoluteAnchor.Position?.Y?.Value));
                ResolveAbsoluteColumn(geometry, absoluteX, out int absoluteColumn, out int columnOffset);
                ResolveAbsoluteRow(geometry, absoluteY, out int absoluteRow, out int rowOffset);
                return new AnchorPosition(
                    absoluteRow,
                    absoluteColumn,
                    columnOffset,
                    rowOffset,
                    ExcelImageExportLimits.ClampExtentPixels(ParseEmuPixels(absoluteAnchor.Extent?.Cx?.Value)),
                    ExcelImageExportLimits.ClampExtentPixels(ParseEmuPixels(absoluteAnchor.Extent?.Cy?.Value)),
                    toColumn: null,
                    toRow: null,
                    toOffsetXPixels: 0,
                    toOffsetYPixels: 0);
            }

            Xdr.MarkerType? fromMarker = anchor switch {
                Xdr.OneCellAnchor oneCellAnchor => oneCellAnchor.FromMarker,
                Xdr.TwoCellAnchor twoCellAnchor => twoCellAnchor.FromMarker,
                _ => null,
            };
            Xdr.MarkerType? toMarker = anchor is Xdr.TwoCellAnchor twoCell ? twoCell.ToMarker : null;
            Xdr.Extent? extent = anchor is Xdr.OneCellAnchor oneCell ? oneCell.Extent : null;

            int row = ParseOneBasedMarker(fromMarker?.RowId?.Text, 1048576);
            int column = ParseOneBasedMarker(fromMarker?.ColumnId?.Text, 16384);
            int offsetX = ExcelImageExportLimits.ClampOffsetPixels(ParseEmuPixels(fromMarker?.ColumnOffset?.Text));
            int offsetY = ExcelImageExportLimits.ClampOffsetPixels(ParseEmuPixels(fromMarker?.RowOffset?.Text));
            int? toRow = ParseOneBasedMarkerOrNull(toMarker?.RowId?.Text, 1048576);
            int? toColumn = ParseOneBasedMarkerOrNull(toMarker?.ColumnId?.Text, 16384);
            int toOffsetX = ExcelImageExportLimits.ClampOffsetPixels(ParseEmuPixels(toMarker?.ColumnOffset?.Text));
            int toOffsetY = ExcelImageExportLimits.ClampOffsetPixels(ParseEmuPixels(toMarker?.RowOffset?.Text));
            int width = ExcelImageExportLimits.ClampExtentPixels(ParseEmuPixels(extent?.Cx?.Value));
            int height = ExcelImageExportLimits.ClampExtentPixels(ParseEmuPixels(extent?.Cy?.Value));
            if (anchor is Xdr.TwoCellAnchor && toColumn.HasValue && toRow.HasValue) {
                width = ResolveTwoCellWidthPixels(geometry, column, offsetX, toColumn.Value, toOffsetX);
                height = ResolveTwoCellHeightPixels(geometry, row, offsetY, toRow.Value, toOffsetY);
            }

            return new AnchorPosition(row, column, offsetX, offsetY, width, height, toColumn, toRow, toOffsetX, toOffsetY);
        }

        private static void ResolveAbsoluteColumn(ExcelWorksheetGeometryIndex geometry, int absolutePixels, out int column, out int offsetPixels) {
            absolutePixels = ExcelImageExportLimits.ClampAbsoluteOffsetPixels(absolutePixels);
            int cursor = 0;
            int maximumColumn = Math.Min(16384, ExcelImageExportLimits.MaximumAnchorSpanCells);
            for (column = 1; column <= maximumColumn; column++) {
                int width = geometry.GetSimpleColumnWidthPixels(column);
                if (cursor + width >= absolutePixels) {
                    offsetPixels = Math.Max(0, absolutePixels - cursor);
                    return;
                }

                cursor += width;
            }

            column = maximumColumn;
            offsetPixels = 0;
        }

        private static void ResolveAbsoluteRow(ExcelWorksheetGeometryIndex geometry, int absolutePixels, out int row, out int offsetPixels) {
            absolutePixels = ExcelImageExportLimits.ClampAbsoluteOffsetPixels(absolutePixels);
            int cursor = 0;
            int maximumRow = Math.Min(1048576, ExcelImageExportLimits.MaximumAnchorSpanCells);
            for (row = 1; row <= maximumRow; row++) {
                int height = geometry.GetRowHeightPixels(row);
                if (cursor + height >= absolutePixels) {
                    offsetPixels = Math.Max(0, absolutePixels - cursor);
                    return;
                }

                cursor += height;
            }

            row = maximumRow;
            offsetPixels = 0;
        }

        private static int ResolveTwoCellWidthPixels(ExcelWorksheetGeometryIndex geometry, int fromColumn, int fromOffsetPixels, int toColumn, int toOffsetPixels) {
            if (!ExcelImageExportLimits.IsValidMarkerSpan(fromColumn - 1, toColumn - 1, 16384)) {
                return 0;
            }

            long pixels = -(long)fromOffsetPixels + toOffsetPixels;
            for (int column = fromColumn; column < toColumn; column++) {
                pixels += geometry.GetSimpleColumnWidthPixels(column);
            }

            return ExcelImageExportLimits.ClampExtentPixels((int)Math.Max(0L, Math.Min(int.MaxValue, pixels)));
        }

        private static int ResolveTwoCellHeightPixels(ExcelWorksheetGeometryIndex geometry, int fromRow, int fromOffsetPixels, int toRow, int toOffsetPixels) {
            if (!ExcelImageExportLimits.IsValidMarkerSpan(fromRow - 1, toRow - 1, 1048576)) {
                return 0;
            }

            long pixels = -(long)fromOffsetPixels + toOffsetPixels;
            for (int row = fromRow; row < toRow; row++) {
                pixels += geometry.GetRowHeightPixels(row);
            }

            return ExcelImageExportLimits.ClampExtentPixels((int)Math.Max(0L, Math.Min(int.MaxValue, pixels)));
        }

        private static bool TryGetShapePreset(Xdr.Shape shape, out string shapePresetName, out OfficeShapeKind shapeKind, out string? unsupportedReason) {
            shapePresetName = shape.ShapeProperties?.GetFirstChild<A.PresetGeometry>()?.Preset?.InnerText ?? string.Empty;
            shapeKind = OfficeShapeKind.Rectangle;
            if (!OfficeShapePresets.TryCreate(shapePresetName, width: 1D, height: 1D, out OfficeShape? sharedShape) || sharedShape == null) {
                unsupportedReason = string.IsNullOrWhiteSpace(shapePresetName)
                    ? "shape geometry is missing"
                    : "shape geometry '" + shapePresetName + "' is not rendered yet";
                return false;
            }

            shapeKind = sharedShape.Kind;
            unsupportedReason = null;
            return true;
        }

        private static bool TryGetFillColor(OpenXmlCompositeElement? properties, WorkbookPart? workbookPart, out string? fillColorArgb, out string? unsupportedReason) {
            fillColorArgb = null;
            unsupportedReason = null;
            if (properties == null || properties.GetFirstChild<A.NoFill>() != null) {
                return true;
            }

            A.SolidFill? solidFill = properties.GetFirstChild<A.SolidFill>();
            solidFill ??= GetStyleFill(properties);
            if (solidFill == null) {
                unsupportedReason = "shape fill is not a supported solid RGB fill";
                return false;
            }

            fillColorArgb = ExcelThemeColorResolver.Resolve(solidFill, workbookPart);
            if (fillColorArgb != null) {
                return true;
            }

            unsupportedReason = "shape fill color could not be resolved by the dependency-free exporter";
            return false;
        }

        private static bool TryGetStroke(
            OpenXmlCompositeElement? properties,
            WorkbookPart? workbookPart,
            out string? strokeColorArgb,
            out double strokeWidth,
            out OfficeStrokeDashStyle strokeDashStyle,
            out OfficeStrokeLineCap? strokeLineCap,
            out OfficeStrokeLineJoin? strokeLineJoin,
            out string? unsupportedReason) {
            strokeColorArgb = null;
            strokeWidth = 1D;
            strokeDashStyle = OfficeStrokeDashStyle.Solid;
            strokeLineCap = null;
            strokeLineJoin = null;
            unsupportedReason = null;
            A.Outline? outline = properties?.GetFirstChild<A.Outline>();
            if (outline == null) {
                A.SolidFill? styleLineFill = GetStyleLineFill(properties);
                if (styleLineFill == null) {
                    strokeWidth = 0D;
                    return true;
                }

                strokeColorArgb = ExcelThemeColorResolver.Resolve(styleLineFill, workbookPart);
                if (strokeColorArgb == null) {
                    unsupportedReason = "shape outline color could not be resolved by the dependency-free exporter";
                    return false;
                }

                return true;
            }

            if (outline.GetFirstChild<A.NoFill>() != null) {
                strokeWidth = 0D;
                return true;
            }

            A.SolidFill? solidFill = outline.GetFirstChild<A.SolidFill>();
            if (solidFill == null) {
                unsupportedReason = "shape outline is not a supported solid RGB line";
                return false;
            }

            strokeColorArgb = ExcelThemeColorResolver.Resolve(solidFill, workbookPart);
            if (strokeColorArgb == null) {
                unsupportedReason = "shape outline color could not be resolved by the dependency-free exporter";
                return false;
            }

            if (outline.Width != null && outline.Width.Value > 0) {
                strokeWidth = ExcelImageExportLimits.ClampStrokeWidth(Math.Max(1D, outline.Width.Value / EmusPerPixel));
            }

            strokeDashStyle = ResolveStrokeDashStyle(outline);
            strokeLineCap = ResolveStrokeLineCap(outline.CapType?.Value);
            strokeLineJoin = ResolveStrokeLineJoin(outline);
            return true;
        }

        private static OfficeStrokeDashStyle ResolveStrokeDashStyle(A.Outline outline) {
            A.PresetLineDashValues? dash = outline.GetFirstChild<A.PresetDash>()?.Val?.Value;
            if (!dash.HasValue || dash.Value == A.PresetLineDashValues.Solid) {
                return OfficeStrokeDashStyle.Solid;
            }

            if (dash.Value == A.PresetLineDashValues.Dash ||
                dash.Value == A.PresetLineDashValues.LargeDash ||
                dash.Value == A.PresetLineDashValues.SystemDash) {
                return OfficeStrokeDashStyle.Dash;
            }

            if (dash.Value == A.PresetLineDashValues.Dot ||
                dash.Value == A.PresetLineDashValues.SystemDot) {
                return OfficeStrokeDashStyle.Dot;
            }

            if (dash.Value == A.PresetLineDashValues.DashDot ||
                dash.Value == A.PresetLineDashValues.LargeDashDot ||
                dash.Value == A.PresetLineDashValues.SystemDashDot) {
                return OfficeStrokeDashStyle.DashDot;
            }

            if (dash.Value == A.PresetLineDashValues.LargeDashDotDot ||
                dash.Value == A.PresetLineDashValues.SystemDashDotDot) {
                return OfficeStrokeDashStyle.DashDotDot;
            }

            return OfficeStrokeDashStyle.Solid;
        }

        private static OfficeStrokeLineCap? ResolveStrokeLineCap(A.LineCapValues? cap) {
            if (!cap.HasValue) {
                return null;
            }

            if (cap.Value == A.LineCapValues.Round) {
                return OfficeStrokeLineCap.Round;
            }

            if (cap.Value == A.LineCapValues.Square) {
                return OfficeStrokeLineCap.Square;
            }

            if (cap.Value == A.LineCapValues.Flat) {
                return OfficeStrokeLineCap.Butt;
            }

            return null;
        }

        private static OfficeStrokeLineJoin? ResolveStrokeLineJoin(A.Outline outline) {
            if (outline.GetFirstChild<A.Round>() != null || HasOutlineChild(outline, "round")) {
                return OfficeStrokeLineJoin.Round;
            }

            if (outline.GetFirstChild<A.Bevel>() != null || HasOutlineChild(outline, "bevel")) {
                return OfficeStrokeLineJoin.Bevel;
            }

            if (outline.GetFirstChild<A.Miter>() != null || HasOutlineChild(outline, "miter")) {
                return OfficeStrokeLineJoin.Miter;
            }

            return null;
        }

        private static bool HasOutlineChild(A.Outline outline, string localName) =>
            outline.ChildElements.Any(child => string.Equals(child.LocalName, localName, StringComparison.Ordinal));

        private static bool TryCreateGlow(A.Glow? glow, WorkbookPart? workbookPart, out OfficeGlow? officeGlow) {
            officeGlow = null;
            if (glow == null) {
                return false;
            }

            double radius = ParseEmuPixels(glow.Radius?.Value);
            if (radius <= 0D || !TryResolveEffectColor(glow, workbookPart, out OfficeColor color, out double opacity)) {
                return false;
            }

            officeGlow = new OfficeGlow(color, opacity, radius);
            return true;
        }

        private static bool TryCreateShadow(A.OuterShadow? shadow, WorkbookPart? workbookPart, out OfficeShadow? officeShadow) {
            officeShadow = null;
            if (shadow == null) {
                return false;
            }

            double distance = ParseEmuPixels(shadow.Distance?.Value);
            double blurRadius = ParseEmuPixels(shadow.BlurRadius?.Value);
            if ((distance <= 0D && blurRadius <= 0D) || !TryResolveEffectColor(shadow, workbookPart, out OfficeColor color, out double opacity)) {
                return false;
            }

            double angleDegrees = (shadow.Direction?.Value ?? 0) / 60000D;
            double radians = angleDegrees * Math.PI / 180D;
            double offsetX = Math.Cos(radians) * distance;
            double offsetY = Math.Sin(radians) * distance;
            officeShadow = new OfficeShadow(color, opacity, offsetX, offsetY, blurRadius);
            return true;
        }

        private static bool TryResolveEffectColor(OpenXmlCompositeElement owner, WorkbookPart? workbookPart, out OfficeColor color, out double opacity) {
            color = default;
            opacity = 0D;
            OpenXmlElement? colorElement = owner.GetFirstChild<A.RgbColorModelHex>();
            colorElement ??= owner.GetFirstChild<A.SchemeColor>();
            colorElement ??= owner.GetFirstChild<A.SystemColor>();
            if (colorElement == null) {
                return false;
            }

            var fill = new A.SolidFill();
            fill.Append((OpenXmlElement)colorElement.CloneNode(true));
            string? argb = ExcelThemeColorResolver.Resolve(fill, workbookPart);
            if (!TryParseArgb(argb, out byte alpha, out byte red, out byte green, out byte blue) || alpha == 0) {
                return false;
            }

            color = OfficeColor.FromRgb(red, green, blue);
            opacity = alpha / 255D;
            return true;
        }

        private static bool TryParseArgb(string? argb, out byte alpha, out byte red, out byte green, out byte blue) {
            alpha = red = green = blue = 0;
            if (string.IsNullOrWhiteSpace(argb)) {
                return false;
            }

            string value = argb!.Trim().TrimStart('#');
            if (value.Length != 8) {
                return false;
            }

            return byte.TryParse(value.Substring(0, 2), System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out alpha) &&
                byte.TryParse(value.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out red) &&
                byte.TryParse(value.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out green) &&
                byte.TryParse(value.Substring(6, 2), System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out blue);
        }

        private static string GetDrawingParagraphText(A.Paragraph paragraph) {
            var builder = new System.Text.StringBuilder();
            AppendDrawingText(paragraph, builder);
            return builder.ToString();
        }

        private static void AppendDrawingText(OpenXmlElement element, System.Text.StringBuilder builder) {
            if (element is A.Text text) {
                builder.Append(text.Text);
                return;
            }

            if (element is A.Break) {
                builder.Append(Environment.NewLine);
                return;
            }

            foreach (OpenXmlElement child in element.ChildElements) {
                AppendDrawingText(child, builder);
            }
        }

        private static A.SolidFill? GetStyleFill(OpenXmlElement properties) {
            Xdr.ShapeStyle? style = properties.Parent?.GetFirstChild<Xdr.ShapeStyle>();
            A.SchemeColor? schemeColor = style?
                .GetFirstChild<A.FillReference>()?
                .GetFirstChild<A.SchemeColor>();
            if (schemeColor == null) {
                return null;
            }

            return new A.SolidFill((A.SchemeColor)schemeColor.CloneNode(true));
        }

        private static A.SolidFill? GetStyleLineFill(OpenXmlElement? properties) {
            Xdr.ShapeStyle? style = properties?.Parent?.GetFirstChild<Xdr.ShapeStyle>();
            A.SchemeColor? schemeColor = style?
                .GetFirstChild<A.LineReference>()?
                .GetFirstChild<A.SchemeColor>();
            if (schemeColor == null) {
                return null;
            }

            return new A.SolidFill((A.SchemeColor)schemeColor.CloneNode(true));
        }

        private static bool TryGetRotationDegrees(A.Transform2D? transform, out double rotationDegrees) {
            rotationDegrees = 0D;
            if (transform == null) {
                return false;
            }

            if (!int.TryParse(GetRawAttribute(transform, "rot"), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int rotation)) {
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

        private static string? NormalizeFontFamily(string? value) =>
            string.IsNullOrWhiteSpace(value) ? null : value!.Trim();

        private static string? GetRawAttribute(OpenXmlElement? element, string localName) {
            if (element == null) {
                return null;
            }

            foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
                if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal)) {
                    return attribute.Value;
                }
            }

            return null;
        }

        private static bool ReadBooleanAttribute(OpenXmlElement? element, string localName) {
            string? value = GetRawAttribute(element, localName);
            return string.Equals(value, "1", StringComparison.Ordinal) ||
                string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
        }

        private static int ParseOneBasedMarker(string? value, int maximum) =>
            int.TryParse(value, out int zeroBased) && zeroBased >= 0 && zeroBased < maximum ? zeroBased + 1 : 0;

        private static int? ParseOneBasedMarkerOrNull(string? value, int maximum) {
            int parsed = ParseOneBasedMarker(value, maximum);
            return parsed > 0 ? parsed : null;
        }

        private static int ParseEmuPixels(string? value) =>
            long.TryParse(value, out long emus) ? ConvertEmuToPixels(emus) : 0;

        private static int ParseEmuPixels(long? value) =>
            value.HasValue ? ConvertEmuToPixels(value.Value) : 0;

        private static int ParseEmuPixels(string? value, long defaultValue) =>
            long.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out long emus)
                ? ConvertEmuToPixels(emus)
                : ConvertEmuToPixels(defaultValue);

        private static int ConvertEmuToPixels(long emus) {
            double pixels = Math.Round(emus / EmusPerPixel);
            if (pixels <= 0D) {
                return 0;
            }

            return pixels >= int.MaxValue ? int.MaxValue : (int)pixels;
        }

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

        private readonly struct DrawingTextStyle {
            internal DrawingTextStyle(string? colorArgb, string? fontFamily, double? fontSize, OfficeFontStyle fontStyle) {
                ColorArgb = colorArgb;
                FontFamily = fontFamily;
                FontSize = fontSize;
                FontStyle = fontStyle;
            }

            internal static DrawingTextStyle Default => new DrawingTextStyle(null, null, null, OfficeFontStyle.Regular);

            internal string? ColorArgb { get; }

            internal string? FontFamily { get; }

            internal double? FontSize { get; }

            internal OfficeFontStyle FontStyle { get; }
        }

        private readonly struct DrawingTextInsets {
            internal static DrawingTextInsets None { get; } = new DrawingTextInsets(0D, 0D, 0D, 0D);

            internal DrawingTextInsets(double left, double top, double right, double bottom) {
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
            }

            internal double Left { get; }

            internal double Top { get; }

            internal double Right { get; }

            internal double Bottom { get; }
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
            string shapePresetName,
            OfficeShapeKind? shapeKind,
            bool horizontalFlip,
            bool verticalFlip,
            double rotationDegrees,
            string? fillColorArgb,
            string? strokeColorArgb,
            double strokeWidth,
            OfficeStrokeDashStyle strokeDashStyle,
            OfficeStrokeLineCap? strokeLineCap,
            OfficeStrokeLineJoin? strokeLineJoin,
            string text,
            OfficeTextAlignment textAlignment,
            OfficeTextVerticalAlignment textVerticalAlignment,
            string? textColorArgb,
            string? textFontFamily,
            double? textFontSize,
            OfficeFontStyle textFontStyle,
            bool textWrap,
            bool textShrinkToFit,
            bool textResizeShapeToFit,
            ExcelDrawingTextOrientation textOrientation,
            double textInsetLeft,
            double textInsetTop,
            double textInsetRight,
            double textInsetBottom,
            OfficeGlow? glow,
            OfficeShadow? shadow,
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
            ShapePresetName = shapePresetName ?? string.Empty;
            ShapeKind = shapeKind;
            HorizontalFlip = horizontalFlip;
            VerticalFlip = verticalFlip;
            RotationDegrees = rotationDegrees;
            FillColorArgb = fillColorArgb;
            StrokeColorArgb = strokeColorArgb;
            StrokeWidth = strokeWidth;
            StrokeDashStyle = strokeDashStyle;
            StrokeLineCap = strokeLineCap;
            StrokeLineJoin = strokeLineJoin;
            Text = text ?? string.Empty;
            TextAlignment = textAlignment;
            TextVerticalAlignment = textVerticalAlignment;
            TextColorArgb = textColorArgb;
            TextFontFamily = textFontFamily;
            TextFontSize = textFontSize;
            TextFontStyle = textFontStyle;
            TextWrap = textWrap;
            TextShrinkToFit = textShrinkToFit;
            TextResizeShapeToFit = textResizeShapeToFit;
            TextOrientation = textOrientation;
            TextInsetLeft = textInsetLeft;
            TextInsetTop = textInsetTop;
            TextInsetRight = textInsetRight;
            TextInsetBottom = textInsetBottom;
            Glow = glow;
            Shadow = shadow;
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

        internal string ShapePresetName { get; }

        internal OfficeShapeKind? ShapeKind { get; }

        internal bool HorizontalFlip { get; }

        internal bool VerticalFlip { get; }

        internal double RotationDegrees { get; }

        internal string? FillColorArgb { get; }

        internal string? StrokeColorArgb { get; }

        internal double StrokeWidth { get; }

        internal OfficeStrokeDashStyle StrokeDashStyle { get; }

        internal OfficeStrokeLineCap? StrokeLineCap { get; }

        internal OfficeStrokeLineJoin? StrokeLineJoin { get; }

        internal string Text { get; }

        internal OfficeTextAlignment TextAlignment { get; }

        internal OfficeTextVerticalAlignment TextVerticalAlignment { get; }

        internal string? TextColorArgb { get; }

        internal string? TextFontFamily { get; }

        internal double? TextFontSize { get; }

        internal OfficeFontStyle TextFontStyle { get; }

        internal bool TextWrap { get; }

        internal bool TextShrinkToFit { get; }

        internal bool TextResizeShapeToFit { get; }

        internal ExcelDrawingTextOrientation TextOrientation { get; }

        internal double TextInsetLeft { get; }

        internal double TextInsetTop { get; }

        internal double TextInsetRight { get; }

        internal double TextInsetBottom { get; }

        internal OfficeGlow? Glow { get; }

        internal OfficeShadow? Shadow { get; }

        internal string? UnsupportedReason { get; }

        internal bool IsRenderable => ShapeKind.HasValue && string.IsNullOrEmpty(UnsupportedReason);

        internal string? CellReference => Row > 0 && Column > 0 ? A1.CellReference(Row, Column) : null;
    }
}
