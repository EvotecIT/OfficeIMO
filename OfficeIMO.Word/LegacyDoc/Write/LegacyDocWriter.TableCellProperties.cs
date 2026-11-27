using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static IReadOnlyList<int> ReadSupportedTableCellWidths(IReadOnlyList<LegacyDocWritableTableCell> cells) {
            var widths = new int[cells.Count];
            for (int index = 0; index < cells.Count; index++) {
                widths[index] = cells[index].WidthTwips;
            }

            return widths;
        }

        private static IReadOnlyList<LegacyDocTableCellHorizontalMerge> ReadSupportedTableCellHorizontalMerges(IReadOnlyList<LegacyDocWritableTableCell> cells) {
            var merges = new LegacyDocTableCellHorizontalMerge[cells.Count];
            bool hasMerge = false;
            for (int index = 0; index < cells.Count; index++) {
                merges[index] = cells[index].HorizontalMerge;
                if (merges[index] != LegacyDocTableCellHorizontalMerge.None) {
                    hasMerge = true;
                }
            }

            return hasMerge ? merges : Array.Empty<LegacyDocTableCellHorizontalMerge>();
        }

        private static IReadOnlyList<LegacyDocTableCellVerticalMerge> ReadSupportedTableCellVerticalMerges(IReadOnlyList<LegacyDocWritableTableCell> cells) {
            var merges = new LegacyDocTableCellVerticalMerge[cells.Count];
            bool hasMerge = false;
            for (int index = 0; index < cells.Count; index++) {
                merges[index] = cells[index].VerticalMerge;
                if (merges[index] != LegacyDocTableCellVerticalMerge.None) {
                    hasMerge = true;
                }
            }

            return hasMerge ? merges : Array.Empty<LegacyDocTableCellVerticalMerge>();
        }

        private static IReadOnlyList<LegacyDocTableCellVerticalAlignment> ReadSupportedTableCellVerticalAlignments(IReadOnlyList<LegacyDocWritableTableCell> cells) {
            var alignments = new LegacyDocTableCellVerticalAlignment[cells.Count];
            bool hasNonDefaultAlignment = false;
            for (int index = 0; index < cells.Count; index++) {
                alignments[index] = cells[index].VerticalAlignment;
                if (alignments[index] != LegacyDocTableCellVerticalAlignment.Top) {
                    hasNonDefaultAlignment = true;
                }
            }

            return hasNonDefaultAlignment ? alignments : Array.Empty<LegacyDocTableCellVerticalAlignment>();
        }

        private static IReadOnlyList<LegacyDocTableCellTextDirection> ReadSupportedTableCellTextDirections(IReadOnlyList<LegacyDocWritableTableCell> cells) {
            var textDirections = new LegacyDocTableCellTextDirection[cells.Count];
            bool hasNonDefaultTextDirection = false;
            for (int index = 0; index < cells.Count; index++) {
                textDirections[index] = cells[index].TextDirection;
                if (textDirections[index] != LegacyDocTableCellTextDirection.LeftToRightTopToBottom) {
                    hasNonDefaultTextDirection = true;
                }
            }

            return hasNonDefaultTextDirection ? textDirections : Array.Empty<LegacyDocTableCellTextDirection>();
        }

        private static IReadOnlyList<bool> ReadSupportedTableCellFitTexts(IReadOnlyList<LegacyDocWritableTableCell> cells) {
            var fitTexts = new bool[cells.Count];
            bool hasFitText = false;
            for (int index = 0; index < cells.Count; index++) {
                fitTexts[index] = cells[index].FitText;
                if (fitTexts[index]) {
                    hasFitText = true;
                }
            }

            return hasFitText ? fitTexts : Array.Empty<bool>();
        }

        private static IReadOnlyList<bool> ReadSupportedTableCellNoWraps(IReadOnlyList<LegacyDocWritableTableCell> cells) {
            var noWraps = new bool[cells.Count];
            bool hasNoWrap = false;
            for (int index = 0; index < cells.Count; index++) {
                noWraps[index] = cells[index].NoWrap;
                if (noWraps[index]) {
                    hasNoWrap = true;
                }
            }

            return hasNoWrap ? noWraps : Array.Empty<bool>();
        }

        private static IReadOnlyList<bool> ReadSupportedTableCellHideMarks(IReadOnlyList<LegacyDocWritableTableCell> cells) {
            var hideMarks = new bool[cells.Count];
            bool hasHideMark = false;
            for (int index = 0; index < cells.Count; index++) {
                hideMarks[index] = cells[index].HideMark;
                if (hideMarks[index]) {
                    hasHideMark = true;
                }
            }

            return hasHideMark ? hideMarks : Array.Empty<bool>();
        }

        private static IReadOnlyList<LegacyDocTableCellMargins> ReadSupportedTableCellMargins(IReadOnlyList<LegacyDocWritableTableCell> cells) {
            var margins = new LegacyDocTableCellMargins[cells.Count];
            bool hasMargins = false;
            for (int index = 0; index < cells.Count; index++) {
                margins[index] = cells[index].Margins;
                if (margins[index].HasAny) {
                    hasMargins = true;
                }
            }

            return hasMargins ? margins : Array.Empty<LegacyDocTableCellMargins>();
        }

        private static IReadOnlyList<LegacyDocTableCellShading> ReadSupportedTableCellShadings(IReadOnlyList<LegacyDocWritableTableCell> cells) {
            var shadings = new LegacyDocTableCellShading[cells.Count];
            bool hasShading = false;
            for (int index = 0; index < cells.Count; index++) {
                shadings[index] = cells[index].Shading;
                if (shadings[index].HasAny) {
                    hasShading = true;
                }
            }

            return hasShading ? shadings : Array.Empty<LegacyDocTableCellShading>();
        }

        private static IReadOnlyList<LegacyDocTableCellBorders> ReadSupportedTableCellBorders(IReadOnlyList<LegacyDocWritableTableCell> cells) {
            var borders = new LegacyDocTableCellBorders[cells.Count];
            bool hasBorders = false;
            for (int index = 0; index < cells.Count; index++) {
                borders[index] = cells[index].Borders;
                if (borders[index].HasAny) {
                    hasBorders = true;
                }
            }

            return hasBorders ? borders : Array.Empty<LegacyDocTableCellBorders>();
        }

        private static LegacyDocTableCellHorizontalMerge ReadSupportedTableCellHorizontalMerge(LegacyDocElement cell) {
            LegacyDocElement horizontalMerge = cell.Child("tcPr").Child("hMerge");
            if (horizontalMerge.IsNone) {
                return LegacyDocTableCellHorizontalMerge.None;
            }

            string? value = horizontalMerge.Val;
            if (value == "restart") {
                return LegacyDocTableCellHorizontalMerge.Restart;
            }

            if (value == "continue") {
                return LegacyDocTableCellHorizontalMerge.Continue;
            }

            throw new NotSupportedException($"Native DOC saving does not support table cell horizontal merge value '{FormatSdkEnumValue(value, "MergedCellValues")}'.");
        }

        private static LegacyDocTableCellVerticalMerge ReadSupportedTableCellVerticalMerge(LegacyDocElement cell) {
            LegacyDocElement verticalMerge = cell.Child("tcPr").Child("vMerge");
            if (verticalMerge.IsNone) {
                return LegacyDocTableCellVerticalMerge.None;
            }

            string? value = verticalMerge.Val;
            if (value == "restart") {
                return LegacyDocTableCellVerticalMerge.Restart;
            }

            if (value == null || value == "continue") {
                return LegacyDocTableCellVerticalMerge.Continue;
            }

            throw new NotSupportedException($"Native DOC saving does not support table cell vertical merge value '{FormatSdkEnumValue(value, "MergedCellValues")}'.");
        }

        private static LegacyDocTableCellVerticalAlignment ReadSupportedTableCellVerticalAlignmentFromProperties(LegacyDocElement cellProperties) {
            LegacyDocElement verticalAlignment = cellProperties.Child("vAlign");
            return ReadSupportedTableCellVerticalAlignment(verticalAlignment) ?? LegacyDocTableCellVerticalAlignment.Top;
        }

        private static LegacyDocTableCellVerticalAlignment? ReadSupportedTableCellVerticalAlignment(LegacyDocElement verticalAlignment) {
            if (verticalAlignment.IsNone) {
                return null;
            }

            string? value = verticalAlignment.Val;
            if (value == null || value == "top") {
                return LegacyDocTableCellVerticalAlignment.Top;
            }

            if (value == "center") {
                return LegacyDocTableCellVerticalAlignment.Center;
            }

            if (value == "bottom") {
                return LegacyDocTableCellVerticalAlignment.Bottom;
            }

            throw new NotSupportedException($"Native DOC saving does not support table cell vertical alignment value '{FormatSdkEnumValue(value, "TableVerticalAlignmentValues")}'.");
        }

        private static LegacyDocTableCellTextDirection ReadSupportedTableCellTextDirectionFromProperties(LegacyDocElement cellProperties) {
            LegacyDocElement textDirection = cellProperties.Child("textDirection");
            return ReadSupportedTableCellTextDirection(textDirection) ?? LegacyDocTableCellTextDirection.LeftToRightTopToBottom;
        }

        private static LegacyDocTableCellTextDirection? ReadSupportedTableCellTextDirection(LegacyDocElement textDirection) {
            if (textDirection.IsNone) {
                return null;
            }

            string? value = textDirection.Val;
            if (value == null || value == "lrTb") {
                return LegacyDocTableCellTextDirection.LeftToRightTopToBottom;
            }

            if (value == "tbRl") {
                return LegacyDocTableCellTextDirection.TopToBottomRightToLeft;
            }

            if (value == "btLr") {
                return LegacyDocTableCellTextDirection.BottomToTopLeftToRight;
            }

            if (value == "lrTbV") {
                return LegacyDocTableCellTextDirection.LeftToRightTopToBottomRotated;
            }

            if (value == "tbRlV") {
                return LegacyDocTableCellTextDirection.TopToBottomRightToLeftRotated;
            }

            throw new NotSupportedException($"Native DOC saving does not support table cell text direction value '{FormatSdkEnumValue(value, "TextDirectionValues")}'.");
        }

        private static bool ReadSupportedTableCellFitTextFromProperties(LegacyDocElement cellProperties) {
            LegacyDocElement fitText = cellProperties.Child("tcFitText");
            return ReadSupportedTableCellFitText(fitText) == true;
        }

        private static bool? ReadSupportedTableCellFitText(LegacyDocElement fitText) {
            return fitText.IsNone ? null : ReadTableCellOnOffValue(fitText);
        }

        private static bool ReadSupportedTableCellNoWrapFromProperties(LegacyDocElement cellProperties) {
            LegacyDocElement noWrap = cellProperties.Child("noWrap");
            return ReadSupportedTableCellNoWrap(noWrap) == true;
        }

        private static bool? ReadSupportedTableCellNoWrap(LegacyDocElement noWrap) {
            return noWrap.IsNone ? null : ReadTableCellOnOffValue(noWrap);
        }

        private static bool ReadSupportedTableCellHideMarkFromProperties(LegacyDocElement cellProperties) {
            LegacyDocElement hideMark = cellProperties.Child("hideMark");
            return ReadSupportedTableCellHideMark(hideMark) == true;
        }

        private static bool? ReadSupportedTableCellHideMark(LegacyDocElement hideMark) {
            return hideMark.IsNone ? null : ReadTableCellOnOffValue(hideMark);
        }

        private static LegacyDocTableCellMargins ReadSupportedTableCellMarginsFromProperties(LegacyDocElement cellProperties) {
            LegacyDocElement margins = cellProperties.Child("tcMar");
            return ReadSupportedTableCellMargins(margins);
        }

        private static LegacyDocTableCellMargins ReadSupportedTableCellMargins(LegacyDocElement margins) {
            if (margins.IsNone) {
                return default;
            }

            return new LegacyDocTableCellMargins(
                ReadSupportedTableCellMarginWidth(margins.Child("top"), "top"),
                ReadSupportedTableCellMarginWidth(margins.Child("right"), "right"),
                ReadSupportedTableCellMarginWidth(margins.Child("bottom"), "bottom"),
                ReadSupportedTableCellMarginWidth(margins.Child("left"), "left"));
        }

        private static int? ReadSupportedTableCellMarginWidth(LegacyDocElement margin, string sideName) {
            if (margin.IsNone) {
                return null;
            }

            string? widthText = FindTableAttributeIgnoringCase(margin, "w");
            string? typeText = FindTableAttributeIgnoringCase(margin, "type");
            if (!string.IsNullOrEmpty(typeText) && !string.Equals(typeText, "dxa", StringComparison.OrdinalIgnoreCase)) {
                throw new NotSupportedException($"Native DOC saving supports table cell {sideName} margins only as DXA twip values.");
            }

            if (string.IsNullOrWhiteSpace(widthText)) {
                return null;
            }

            if (!int.TryParse(widthText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int width)
                || width < 0
                || width > 31680) {
                throw new NotSupportedException($"Native DOC saving supports table cell {sideName} margins only as nonnegative DXA twip values within the Word 97-2003 limit.");
            }

            return width;
        }

        private static bool ReadTableCellOnOffValue(LegacyDocElement element) {
            string? value = FindTableAttributeIgnoringCase(element, "val");
            return ReadTableRowOnOffValue(value);
        }

        private static int GetGridColumnWidth(IReadOnlyList<int> gridColumnWidthsTwips, int columnIndex) {
            return columnIndex < gridColumnWidthsTwips.Count ? gridColumnWidthsTwips[columnIndex] : 0;
        }

        private static int ReadSupportedTableCellWidth(LegacyDocElement cellProperties, int gridColumnWidthTwips) {
            int? explicitWidth = ReadSupportedExplicitTableCellWidth(cellProperties);
            if (explicitWidth != null) {
                return explicitWidth.Value;
            }

            return gridColumnWidthTwips > 0 ? gridColumnWidthTwips : 2400;
        }

        private static int ReadSupportedTableCellWidthForSpan(LegacyDocElement cellProperties, IReadOnlyList<int> gridColumnWidthsTwips, int logicalColumnIndex, int spanIndex, int gridSpan) {
            int gridColumnWidth = GetGridColumnWidth(gridColumnWidthsTwips, logicalColumnIndex + spanIndex);
            if (gridColumnWidth > 0) {
                return gridColumnWidth;
            }

            int? explicitWidth = ReadSupportedExplicitTableCellWidth(cellProperties);
            if (explicitWidth == null) {
                return 2400;
            }

            if (gridSpan == 1) {
                return explicitWidth.Value;
            }

            int baseWidth = explicitWidth.Value / gridSpan;
            int remainder = explicitWidth.Value % gridSpan;
            int width = baseWidth + (spanIndex < remainder ? 1 : 0);
            return width > 0 ? width : 1;
        }

        private static int? ReadSupportedExplicitTableCellWidth(LegacyDocElement cellProperties) {
            LegacyDocElement cellWidth = cellProperties.Child("tcW");
            if (cellWidth.IsNone) {
                return null;
            }

            if (cellWidth.Attr("type") == null || LegacyDocWordHelpers.NativeTableWidthType(cellWidth) != "dxa") {
                throw new NotSupportedException("Native DOC saving supports simple table cell widths only as explicit DXA twip values.");
            }

            string? widthText = cellWidth.Attr("w");
            if (string.IsNullOrWhiteSpace(widthText)
                || !int.TryParse(widthText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int width)
                || width <= 0
                || width > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports simple table cell widths only within the Word 97-2003 signed twip range.");
            }

            return width;
        }

        private static LegacyDocTableCellShading ReadSupportedTableCellShading(LegacyDocElement cellProperties) {
            LegacyDocElement shading = cellProperties.Child("shd");
            if (shading.IsNone) {
                return default;
            }

            return ReadSupportedTableCellShading(shading, "table cell shading");
        }

        private static LegacyDocTableCellShading ReadSupportedTableCellShading(LegacyDocElement shading, string featureName) {
            string? pattern = shading.Val;
            if (pattern != null && pattern != "clear") {
                throw new NotSupportedException($"Native DOC saving supports {featureName} only for clear fill patterns.");
            }

            string? fillColorHex = shading.Attr("fill");
            if (string.IsNullOrWhiteSpace(fillColorHex)
                || string.Equals(fillColorHex, "auto", StringComparison.OrdinalIgnoreCase)) {
                return new LegacyDocTableCellShading(null);
            }

            if (!LegacyDocPaletteProjection.CanResolveColor(fillColorHex)) {
                throw new NotSupportedException($"Native DOC saving supports {featureName} only for Word 97-2003 palette fill colors.");
            }

            return new LegacyDocTableCellShading(fillColorHex);
        }

        private static LegacyDocTableCellBorders ReadSupportedTableCellBorders(LegacyDocElement cellProperties) {
            LegacyDocElement borders = cellProperties.Child("tcBorders");
            if (borders.IsNone) {
                return default;
            }

            return new LegacyDocTableCellBorders(
                ReadSupportedTableCellBorder(borders.Child("top")),
                ReadSupportedTableCellBorder(borders.Child("left")),
                ReadSupportedTableCellBorder(borders.Child("bottom")),
                ReadSupportedTableCellBorder(borders.Child("right")));
        }

        private static LegacyDocTableBorders ReadSupportedTableBorders(LegacyDocElement tableProperties, string? tableStyleId, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            LegacyDocElement borders = tableProperties.Child("tblBorders");
            if (borders.Exists) {
                return ReadSupportedTableBorders(borders);
            }

            return ReadSupportedTableStyleBorders(tableStyleId, tableStyleDefinitions);
        }

        private static LegacyDocTableCellShading ReadSupportedTableShading(string? tableStyleId, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            // Direct tblPr shading paints table-spacing gaps, not cell defaults.
            // Its validation remains separate from ordinary style tcPr shading.
            return ReadSupportedTableStyleShading(tableStyleId, tableStyleDefinitions);
        }

        private static LegacyDocTableBorders ReadSupportedTableBorders(LegacyDocElement borders) {
            foreach (LegacyDocElement child in borders.Elements()) {
                switch (child.WName) {
                    case "top":
                    case "left":
                    case "bottom":
                    case "right":
                    case "insideH":
                    case "insideV":
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple table borders only. Unsupported table border: {child.LocalName}.");
                }
            }

            return new LegacyDocTableBorders(
                ReadSupportedTableCellBorder(borders.Child("top")),
                ReadSupportedTableCellBorder(borders.Child("left")),
                ReadSupportedTableCellBorder(borders.Child("bottom")),
                ReadSupportedTableCellBorder(borders.Child("right")),
                ReadSupportedTableCellBorder(borders.Child("insideH")),
                ReadSupportedTableCellBorder(borders.Child("insideV")));
        }

        private static LegacyDocTableCellBorder ReadSupportedTableCellBorder(LegacyDocElement border) {
            if (border.IsNone) {
                return default;
            }

            string? value = border.Val;
            if (value == null) {
                return default;
            }

            if (value == "none" || value == "nil") {
                return new LegacyDocTableCellBorder(LegacyDocTableCellBorderStyle.ExplicitNone, null, 0, 0);
            }

            LegacyDocTableCellBorderStyle style = MapSupportedTableCellBorderStyle(value);
            string? colorHex = border.Attr("color");
            if (string.Equals(colorHex, "auto", StringComparison.OrdinalIgnoreCase)) {
                colorHex = null;
            }

            if (!LegacyDocPaletteProjection.CanResolveColor(colorHex)) {
                throw new NotSupportedException("Native DOC saving supports table cell borders only with Word 97-2003 palette colors.");
            }

            uint? sizeValue = LegacyDocValues.UInt32(border.Attr("sz"));
            int size = sizeValue == null ? 4 : checked((int)sizeValue.Value);
            uint? spaceValue = LegacyDocValues.UInt32(border.Attr("space"));
            int space = spaceValue == null ? 0 : checked((int)spaceValue.Value);
            if (size <= 0 || size > byte.MaxValue || space < 0 || space > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table cell border size and spacing only within Word 97-2003 BRC80 byte ranges.");
            }

            return new LegacyDocTableCellBorder(style, colorHex, size, space);
        }

        private static LegacyDocTableCellBorderStyle MapSupportedTableCellBorderStyle(string value) {
            if (value == "single") {
                return LegacyDocTableCellBorderStyle.Single;
            }

            if (value == "double") {
                return LegacyDocTableCellBorderStyle.Double;
            }

            if (value == "dotted") {
                return LegacyDocTableCellBorderStyle.Dotted;
            }

            if (value == "dashed" || value == "dashSmallGap") {
                return LegacyDocTableCellBorderStyle.Dashed;
            }

            throw new NotSupportedException($"Native DOC saving does not support table cell border style '{FormatSdkEnumValue(value, "BorderValues")}'.");
        }

        private static int ReadSupportedGridSpan(LegacyDocElement cellProperties) {
            LegacyDocElement gridSpan = cellProperties.Child("gridSpan");
            if (gridSpan.IsNone) {
                return 1;
            }

            int span = LegacyDocValues.Int32(gridSpan.Val) ?? 1;
            if (span <= 0 || span > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table cell gridSpan only as a positive value within the DOC table column limit.");
            }

            return span;
        }

    }
}
