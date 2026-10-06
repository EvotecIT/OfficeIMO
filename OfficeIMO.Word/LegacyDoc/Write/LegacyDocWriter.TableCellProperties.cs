using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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

        private static LegacyDocTableCellHorizontalMerge ReadSupportedTableCellHorizontalMerge(TableCell cell) {
            HorizontalMerge? horizontalMerge = cell.TableCellProperties?.GetFirstChild<HorizontalMerge>();
            if (horizontalMerge == null) {
                return LegacyDocTableCellHorizontalMerge.None;
            }

            MergedCellValues? value = horizontalMerge.Val?.Value;
            if (value == MergedCellValues.Restart) {
                return LegacyDocTableCellHorizontalMerge.Restart;
            }

            if (value == MergedCellValues.Continue) {
                return LegacyDocTableCellHorizontalMerge.Continue;
            }

            throw new NotSupportedException($"Native DOC saving does not support table cell horizontal merge value '{value}'.");
        }

        private static LegacyDocTableCellVerticalMerge ReadSupportedTableCellVerticalMerge(TableCell cell) {
            VerticalMerge? verticalMerge = cell.TableCellProperties?.GetFirstChild<VerticalMerge>();
            if (verticalMerge == null) {
                return LegacyDocTableCellVerticalMerge.None;
            }

            MergedCellValues? value = verticalMerge.Val?.Value;
            if (value == MergedCellValues.Restart) {
                return LegacyDocTableCellVerticalMerge.Restart;
            }

            if (value == null || value == MergedCellValues.Continue) {
                return LegacyDocTableCellVerticalMerge.Continue;
            }

            throw new NotSupportedException($"Native DOC saving does not support table cell vertical merge value '{value}'.");
        }

        private static LegacyDocTableCellVerticalAlignment ReadSupportedTableCellVerticalAlignment(TableCellProperties? cellProperties) {
            TableCellVerticalAlignment? verticalAlignment = cellProperties?.GetFirstChild<TableCellVerticalAlignment>();
            return ReadSupportedTableCellVerticalAlignment(verticalAlignment) ?? LegacyDocTableCellVerticalAlignment.Top;
        }

        private static LegacyDocTableCellVerticalAlignment? ReadSupportedTableCellVerticalAlignment(TableCellVerticalAlignment? verticalAlignment) {
            if (verticalAlignment == null) {
                return null;
            }

            TableVerticalAlignmentValues? value = verticalAlignment.Val?.Value;
            if (value == null || value == TableVerticalAlignmentValues.Top) {
                return LegacyDocTableCellVerticalAlignment.Top;
            }

            if (value == TableVerticalAlignmentValues.Center) {
                return LegacyDocTableCellVerticalAlignment.Center;
            }

            if (value == TableVerticalAlignmentValues.Bottom) {
                return LegacyDocTableCellVerticalAlignment.Bottom;
            }

            throw new NotSupportedException($"Native DOC saving does not support table cell vertical alignment value '{value}'.");
        }

        private static LegacyDocTableCellTextDirection ReadSupportedTableCellTextDirection(TableCellProperties? cellProperties) {
            TextDirection? textDirection = cellProperties?.GetFirstChild<TextDirection>();
            return ReadSupportedTableCellTextDirection(textDirection) ?? LegacyDocTableCellTextDirection.LeftToRightTopToBottom;
        }

        private static LegacyDocTableCellTextDirection? ReadSupportedTableCellTextDirection(TextDirection? textDirection) {
            if (textDirection == null) {
                return null;
            }

            TextDirectionValues? value = textDirection.Val?.Value;
            if (value == null || value == TextDirectionValues.LefToRightTopToBottom) {
                return LegacyDocTableCellTextDirection.LeftToRightTopToBottom;
            }

            if (value == TextDirectionValues.TopToBottomRightToLeft) {
                return LegacyDocTableCellTextDirection.TopToBottomRightToLeft;
            }

            if (value == TextDirectionValues.BottomToTopLeftToRight) {
                return LegacyDocTableCellTextDirection.BottomToTopLeftToRight;
            }

            if (value == TextDirectionValues.LefttoRightTopToBottomRotated) {
                return LegacyDocTableCellTextDirection.LeftToRightTopToBottomRotated;
            }

            if (value == TextDirectionValues.TopToBottomRightToLeftRotated) {
                return LegacyDocTableCellTextDirection.TopToBottomRightToLeftRotated;
            }

            throw new NotSupportedException($"Native DOC saving does not support table cell text direction value '{value}'.");
        }

        private static bool ReadSupportedTableCellFitText(TableCellProperties? cellProperties) {
            TableCellFitText? fitText = cellProperties?.GetFirstChild<TableCellFitText>();
            return ReadSupportedTableCellFitText(fitText) == true;
        }

        private static bool? ReadSupportedTableCellFitText(TableCellFitText? fitText) {
            return fitText == null ? null : ReadTableCellOnOffValue(fitText);
        }

        private static bool ReadSupportedTableCellNoWrap(TableCellProperties? cellProperties) {
            NoWrap? noWrap = cellProperties?.GetFirstChild<NoWrap>();
            return ReadSupportedTableCellNoWrap(noWrap) == true;
        }

        private static bool? ReadSupportedTableCellNoWrap(NoWrap? noWrap) {
            return noWrap == null ? null : ReadTableCellOnOffValue(noWrap);
        }

        private static bool ReadSupportedTableCellHideMark(TableCellProperties? cellProperties) {
            HideMark? hideMark = cellProperties?.GetFirstChild<HideMark>();
            return ReadSupportedTableCellHideMark(hideMark) == true;
        }

        private static bool? ReadSupportedTableCellHideMark(HideMark? hideMark) {
            return hideMark == null ? null : ReadTableCellOnOffValue(hideMark);
        }

        private static LegacyDocTableCellMargins ReadSupportedTableCellMargins(TableCellProperties? cellProperties) {
            TableCellMargin? margins = cellProperties?.GetFirstChild<TableCellMargin>();
            return ReadSupportedTableCellMargins(margins);
        }

        private static LegacyDocTableCellMargins ReadSupportedTableCellMargins(TableCellMargin? margins) {
            if (margins == null) {
                return default;
            }

            return new LegacyDocTableCellMargins(
                ReadSupportedTableCellMarginWidth(margins.TopMargin, "top"),
                ReadSupportedTableCellMarginWidth(margins.RightMargin, "right"),
                ReadSupportedTableCellMarginWidth(margins.BottomMargin, "bottom"),
                ReadSupportedTableCellMarginWidth(margins.LeftMargin, "left"));
        }

        private static int? ReadSupportedTableCellMarginWidth(OpenXmlElement? margin, string sideName) {
            if (margin == null) {
                return null;
            }

            string? widthText = margin.GetAttributes()
                .FirstOrDefault(attribute => string.Equals(attribute.LocalName, "w", StringComparison.OrdinalIgnoreCase))
                .Value;
            string? typeText = margin.GetAttributes()
                .FirstOrDefault(attribute => string.Equals(attribute.LocalName, "type", StringComparison.OrdinalIgnoreCase))
                .Value;
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

        private static bool ReadTableCellOnOffValue(OpenXmlElement element) {
            string? value = element.GetAttributes()
                .FirstOrDefault(attribute => string.Equals(attribute.LocalName, "val", StringComparison.OrdinalIgnoreCase))
                .Value;
            return ReadTableRowOnOffValue(value);
        }

        private static int GetGridColumnWidth(IReadOnlyList<int> gridColumnWidthsTwips, int columnIndex) {
            return columnIndex < gridColumnWidthsTwips.Count ? gridColumnWidthsTwips[columnIndex] : 0;
        }

        private static int ReadSupportedTableCellWidth(TableCellProperties? cellProperties, int gridColumnWidthTwips) {
            int? explicitWidth = ReadSupportedExplicitTableCellWidth(cellProperties);
            if (explicitWidth != null) {
                return explicitWidth.Value;
            }

            return gridColumnWidthTwips > 0 ? gridColumnWidthTwips : 2400;
        }

        private static int ReadSupportedTableCellWidthForSpan(TableCellProperties? cellProperties, IReadOnlyList<int> gridColumnWidthsTwips, int logicalColumnIndex, int spanIndex, int gridSpan) {
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

        private static int? ReadSupportedExplicitTableCellWidth(TableCellProperties? cellProperties) {
            TableCellWidth? cellWidth = cellProperties?.GetFirstChild<TableCellWidth>();
            if (cellWidth == null) {
                return null;
            }

            if (cellWidth.Type?.Value != TableWidthUnitValues.Dxa) {
                throw new NotSupportedException("Native DOC saving supports simple table cell widths only as explicit DXA twip values.");
            }

            string? widthText = cellWidth.Width?.Value;
            if (string.IsNullOrWhiteSpace(widthText)
                || !int.TryParse(widthText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int width)
                || width <= 0
                || width > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports simple table cell widths only within the Word 97-2003 signed twip range.");
            }

            return width;
        }

        private static LegacyDocTableCellShading ReadSupportedTableCellShading(TableCellProperties? cellProperties) {
            Shading? shading = cellProperties?.GetFirstChild<Shading>();
            if (shading == null) {
                return default;
            }

            return ReadSupportedTableCellShading(shading, "table cell shading");
        }

        private static LegacyDocTableCellShading ReadSupportedTableCellShading(Shading shading, string featureName) {
            ShadingPatternValues? pattern = shading.Val?.Value;
            if (pattern != null && pattern != ShadingPatternValues.Clear) {
                throw new NotSupportedException($"Native DOC saving supports {featureName} only for clear fill patterns.");
            }

            string? fillColorHex = shading.Fill?.Value;
            if (string.IsNullOrWhiteSpace(fillColorHex)
                || string.Equals(fillColorHex, "auto", StringComparison.OrdinalIgnoreCase)) {
                return new LegacyDocTableCellShading(null);
            }

            if (!LegacyDocColorPalette.TryGetIcoForHex(fillColorHex, out _)) {
                throw new NotSupportedException($"Native DOC saving supports {featureName} only for Word 97-2003 palette fill colors.");
            }

            return new LegacyDocTableCellShading(fillColorHex);
        }

        private static LegacyDocTableCellBorders ReadSupportedTableCellBorders(TableCellProperties? cellProperties) {
            TableCellBorders? borders = cellProperties?.GetFirstChild<TableCellBorders>();
            if (borders == null) {
                return default;
            }

            return new LegacyDocTableCellBorders(
                ReadSupportedTableCellBorder(borders.TopBorder),
                ReadSupportedTableCellBorder(borders.LeftBorder),
                ReadSupportedTableCellBorder(borders.BottomBorder),
                ReadSupportedTableCellBorder(borders.RightBorder));
        }

        private static LegacyDocTableBorders ReadSupportedTableBorders(TableProperties? tableProperties, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            TableBorders? borders = tableProperties?.GetFirstChild<TableBorders>();
            if (borders != null) {
                return ReadSupportedTableBorders(borders);
            }

            return ReadSupportedTableStyleBorders(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
        }

        private static LegacyDocTableCellShading ReadSupportedTableShading(TableProperties? tableProperties, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            // Direct tblPr shading paints table-spacing gaps, not cell defaults.
            // Its validation remains separate from ordinary style tcPr shading.
            return ReadSupportedTableStyleShading(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
        }

        private static LegacyDocTableBorders ReadSupportedTableBorders(TableBorders borders) {
            foreach (OpenXmlElement child in borders.ChildElements) {
                switch (child) {
                    case TopBorder:
                    case LeftBorder:
                    case BottomBorder:
                    case RightBorder:
                    case InsideHorizontalBorder:
                    case InsideVerticalBorder:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple table borders only. Unsupported table border: {child.LocalName}.");
                }
            }

            return new LegacyDocTableBorders(
                ReadSupportedTableCellBorder(borders.TopBorder),
                ReadSupportedTableCellBorder(borders.LeftBorder),
                ReadSupportedTableCellBorder(borders.BottomBorder),
                ReadSupportedTableCellBorder(borders.RightBorder),
                ReadSupportedTableCellBorder(borders.InsideHorizontalBorder),
                ReadSupportedTableCellBorder(borders.InsideVerticalBorder));
        }

        private static LegacyDocTableCellBorder ReadSupportedTableCellBorder(BorderType? border) {
            if (border == null) {
                return default;
            }

            BorderValues? value = border.Val?.Value;
            if (value == null) {
                return default;
            }

            if (value == BorderValues.None || value == BorderValues.Nil) {
                return new LegacyDocTableCellBorder(LegacyDocTableCellBorderStyle.ExplicitNone, null, 0, 0);
            }

            LegacyDocTableCellBorderStyle style = MapSupportedTableCellBorderStyle(value.Value);
            string? colorHex = border.Color?.Value;
            if (string.Equals(colorHex, "auto", StringComparison.OrdinalIgnoreCase)) {
                colorHex = null;
            }

            if (!LegacyDocColorPalette.TryGetIcoForHex(colorHex, out _)) {
                throw new NotSupportedException("Native DOC saving supports table cell borders only with Word 97-2003 palette colors.");
            }

            int size = border.Size?.Value == null ? 4 : checked((int)border.Size.Value);
            int space = border.Space?.Value == null ? 0 : checked((int)border.Space.Value);
            if (size <= 0 || size > byte.MaxValue || space < 0 || space > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table cell border size and spacing only within Word 97-2003 BRC80 byte ranges.");
            }

            return new LegacyDocTableCellBorder(style, colorHex, size, space);
        }

        private static LegacyDocTableCellBorderStyle MapSupportedTableCellBorderStyle(BorderValues value) {
            if (value == BorderValues.Single) {
                return LegacyDocTableCellBorderStyle.Single;
            }

            if (value == BorderValues.Double) {
                return LegacyDocTableCellBorderStyle.Double;
            }

            if (value == BorderValues.Dotted) {
                return LegacyDocTableCellBorderStyle.Dotted;
            }

            if (value == BorderValues.Dashed || value == BorderValues.DashSmallGap) {
                return LegacyDocTableCellBorderStyle.Dashed;
            }

            throw new NotSupportedException($"Native DOC saving does not support table cell border style '{value}'.");
        }

        private static int ReadSupportedGridSpan(TableCellProperties? cellProperties) {
            GridSpan? gridSpan = cellProperties?.GetFirstChild<GridSpan>();
            if (gridSpan == null) {
                return 1;
            }

            int span = gridSpan.Val?.Value ?? 1;
            if (span <= 0 || span > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table cell gridSpan only as a positive value within the DOC table column limit.");
            }

            return span;
        }

    }
}
