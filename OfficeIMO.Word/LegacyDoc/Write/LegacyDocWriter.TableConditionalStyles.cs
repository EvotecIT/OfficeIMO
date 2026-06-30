using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int TableLookFirstRow = 0x0020;
        private const int TableLookLastRow = 0x0040;
        private const int TableLookFirstColumn = 0x0080;
        private const int TableLookLastColumn = 0x0100;
        private const int TableLookNoHorizontalBand = 0x0200;
        private const int TableLookNoVerticalBand = 0x0400;

        private static LegacyDocTableConditionalStyleSet ReadSupportedTableConditionalStyles(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            Style? style = ResolveSupportedTableStyle(tableStyle, tableStyleDefinitions);
            if (style == null) {
                return LegacyDocTableConditionalStyleSet.Empty;
            }

            StyleTableProperties? styleTableProperties = style.GetFirstChild<StyleTableProperties>();
            int rowBandSize = ReadSupportedTableStyleBandSize(styleTableProperties?.GetFirstChild<TableStyleRowBandSize>()?.Val, "row");
            int columnBandSize = ReadSupportedTableStyleBandSize(styleTableProperties?.GetFirstChild<TableStyleColumnBandSize>()?.Val, "column");
            var conditionalStyles = new List<LegacyDocTableConditionalStyle>();
            foreach (TableStyleProperties properties in style.Elements<TableStyleProperties>()) {
                TableStyleOverrideValues? type = properties.Type?.Value;
                if (type == null) {
                    throw new NotSupportedException($"Native DOC saving supports table style '{style.StyleId?.Value}' conditional formatting only when the conditional type is specified.");
                }

                TableStyleConditionalFormattingTableCellProperties? cellProperties = properties.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>();
                LegacyDocTableCellShading shading = ReadSupportedConditionalTableStyleShading(cellProperties);
                LegacyDocTableCellBorders borders = ReadSupportedConditionalTableStyleBorders(cellProperties);
                if (shading.HasAny || borders.HasAny) {
                    conditionalStyles.Add(new LegacyDocTableConditionalStyle(type.Value, shading, borders));
                }
            }

            return conditionalStyles.Count == 0
                ? new LegacyDocTableConditionalStyleSet(Array.Empty<LegacyDocTableConditionalStyle>(), rowBandSize, columnBandSize)
                : new LegacyDocTableConditionalStyleSet(conditionalStyles, rowBandSize, columnBandSize);
        }

        private static int ReadSupportedTableStyleBandSize(Int32Value? value, string axisName) {
            if (value == null) {
                return 1;
            }

            int bandSize = value.Value;
            if (bandSize <= 0 || bandSize > byte.MaxValue) {
                throw new NotSupportedException($"Native DOC saving supports table style {axisName} band sizes only as positive values within the DOC table column limit.");
            }

            return bandSize;
        }

        private static LegacyDocTableLook ReadSupportedTableLook(TableLook? tableLook) {
            if (tableLook == null) {
                return LegacyDocTableLook.Empty;
            }

            int mask = 0;
            string? value = tableLook.Val?.Value;
            if (!string.IsNullOrWhiteSpace(value)
                && int.TryParse(value, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out int parsed)) {
                mask = parsed;
            }

            ApplyExpandedTableLookFlag(tableLook.FirstRow, TableLookFirstRow, ref mask);
            ApplyExpandedTableLookFlag(tableLook.LastRow, TableLookLastRow, ref mask);
            ApplyExpandedTableLookFlag(tableLook.FirstColumn, TableLookFirstColumn, ref mask);
            ApplyExpandedTableLookFlag(tableLook.LastColumn, TableLookLastColumn, ref mask);
            ApplyExpandedTableLookFlag(tableLook.NoHorizontalBand, TableLookNoHorizontalBand, ref mask);
            ApplyExpandedTableLookFlag(tableLook.NoVerticalBand, TableLookNoVerticalBand, ref mask);

            return new LegacyDocTableLook(
                (mask & TableLookFirstRow) == TableLookFirstRow,
                (mask & TableLookLastRow) == TableLookLastRow,
                (mask & TableLookFirstColumn) == TableLookFirstColumn,
                (mask & TableLookLastColumn) == TableLookLastColumn,
                (mask & TableLookNoHorizontalBand) == TableLookNoHorizontalBand,
                (mask & TableLookNoVerticalBand) == TableLookNoVerticalBand);
        }

        private static void ApplyExpandedTableLookFlag(OnOffValue? value, int flag, ref int mask) {
            if (value == null) {
                return;
            }

            if (value.Value) {
                mask |= flag;
            } else {
                mask &= ~flag;
            }
        }

        private static LegacyDocTableCellShading ReadSupportedConditionalTableStyleShading(TableStyleConditionalFormattingTableCellProperties? cellProperties) {
            Shading? shading = cellProperties?.GetFirstChild<Shading>();
            return shading == null ? default : ReadSupportedTableCellShading(shading, "conditional table style shading");
        }

        private static LegacyDocTableCellBorders ReadSupportedConditionalTableStyleBorders(TableStyleConditionalFormattingTableCellProperties? cellProperties) {
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

        private static IReadOnlyList<LegacyDocWritableTableCell> ApplySupportedTableConditionalStyles(
            IReadOnlyList<LegacyDocWritableTableCell> writableCells,
            LegacyDocTableConditionalStyleSet conditionalStyles,
            LegacyDocTableLook tableLook,
            int rowIndex,
            int rowCount) {
            if (!conditionalStyles.HasAny || writableCells.Count == 0) {
                return writableCells;
            }

            var styledCells = new LegacyDocWritableTableCell[writableCells.Count];
            for (int columnIndex = 0; columnIndex < writableCells.Count; columnIndex++) {
                LegacyDocWritableTableCell cell = writableCells[columnIndex];
                foreach (LegacyDocTableConditionalStyle conditionalStyle in conditionalStyles.Styles) {
                    if (!AppliesToCell(conditionalStyle.Type, tableLook, conditionalStyles.RowBandSize, conditionalStyles.ColumnBandSize, rowIndex, rowCount, columnIndex, writableCells.Count)) {
                        continue;
                    }

                    if (!cell.Shading.HasAny && conditionalStyle.Shading.HasAny) {
                        cell = cell.WithShading(conditionalStyle.Shading);
                    }

                    if (conditionalStyle.Borders.HasAny) {
                        cell = cell.WithBorders(MergeSupportedTableCellBorders(cell.Borders, conditionalStyle.Borders));
                    }
                }

                styledCells[columnIndex] = cell;
            }

            return styledCells;
        }

        private static bool AppliesToCell(
            TableStyleOverrideValues type,
            LegacyDocTableLook tableLook,
            int rowBandSize,
            int columnBandSize,
            int rowIndex,
            int rowCount,
            int columnIndex,
            int columnCount) {
            if (type == TableStyleOverrideValues.FirstRow) {
                return tableLook.FirstRow && rowIndex == 0;
            }

            if (type == TableStyleOverrideValues.LastRow) {
                return tableLook.LastRow && rowIndex + 1 == rowCount;
            }

            if (type == TableStyleOverrideValues.FirstColumn) {
                return tableLook.FirstColumn && columnIndex == 0;
            }

            if (type == TableStyleOverrideValues.LastColumn) {
                return tableLook.LastColumn && columnIndex + 1 == columnCount;
            }

            if (type == TableStyleOverrideValues.NorthWestCell) {
                return tableLook.FirstRow && tableLook.FirstColumn && rowIndex == 0 && columnIndex == 0;
            }

            if (type == TableStyleOverrideValues.NorthEastCell) {
                return tableLook.FirstRow && tableLook.LastColumn && rowIndex == 0 && columnIndex + 1 == columnCount;
            }

            if (type == TableStyleOverrideValues.SouthWestCell) {
                return tableLook.LastRow && tableLook.FirstColumn && rowIndex + 1 == rowCount && columnIndex == 0;
            }

            if (type == TableStyleOverrideValues.SouthEastCell) {
                return tableLook.LastRow && tableLook.LastColumn && rowIndex + 1 == rowCount && columnIndex + 1 == columnCount;
            }

            if (type == TableStyleOverrideValues.Band1Horizontal || type == TableStyleOverrideValues.Band2Horizontal) {
                return !tableLook.NoHorizontalBand && TryGetBandType(rowIndex, tableLook.FirstRow, rowBandSize, type == TableStyleOverrideValues.Band1Horizontal);
            }

            if (type == TableStyleOverrideValues.Band1Vertical || type == TableStyleOverrideValues.Band2Vertical) {
                return !tableLook.NoVerticalBand && TryGetBandType(columnIndex, tableLook.FirstColumn, columnBandSize, type == TableStyleOverrideValues.Band1Vertical);
            }

            throw new NotSupportedException($"Native DOC saving does not support table style conditional type '{type}'.");
        }

        private static bool TryGetBandType(int index, bool skipFirst, int bandSize, bool expectedBand1) {
            int adjustedIndex = skipFirst ? index - 1 : index;
            if (adjustedIndex < 0) {
                return false;
            }

            int bandIndex = adjustedIndex / Math.Max(1, bandSize);
            bool band1 = bandIndex % 2 == 0;
            return band1 == expectedBand1;
        }

        private static LegacyDocTableCellBorders MergeSupportedTableCellBorders(LegacyDocTableCellBorders cellBorders, LegacyDocTableCellBorders inheritedBorders) {
            return new LegacyDocTableCellBorders(
                cellBorders.Top.HasAny ? cellBorders.Top : inheritedBorders.Top,
                cellBorders.Left.HasAny ? cellBorders.Left : inheritedBorders.Left,
                cellBorders.Bottom.HasAny ? cellBorders.Bottom : inheritedBorders.Bottom,
                cellBorders.Right.HasAny ? cellBorders.Right : inheritedBorders.Right);
        }

        private readonly struct LegacyDocTableConditionalStyleSet {
            internal LegacyDocTableConditionalStyleSet(IReadOnlyList<LegacyDocTableConditionalStyle> styles, int rowBandSize, int columnBandSize) {
                Styles = styles;
                RowBandSize = rowBandSize;
                ColumnBandSize = columnBandSize;
            }

            internal static LegacyDocTableConditionalStyleSet Empty { get; } = new LegacyDocTableConditionalStyleSet(Array.Empty<LegacyDocTableConditionalStyle>(), 1, 1);

            internal IReadOnlyList<LegacyDocTableConditionalStyle> Styles { get; }

            internal int RowBandSize { get; }

            internal int ColumnBandSize { get; }

            internal bool HasAny => Styles.Count > 0;
        }

        private readonly struct LegacyDocTableConditionalStyle {
            internal LegacyDocTableConditionalStyle(TableStyleOverrideValues type, LegacyDocTableCellShading shading, LegacyDocTableCellBorders borders) {
                Type = type;
                Shading = shading;
                Borders = borders;
            }

            internal TableStyleOverrideValues Type { get; }

            internal LegacyDocTableCellShading Shading { get; }

            internal LegacyDocTableCellBorders Borders { get; }
        }

        private readonly struct LegacyDocTableLook {
            internal LegacyDocTableLook(bool firstRow, bool lastRow, bool firstColumn, bool lastColumn, bool noHorizontalBand, bool noVerticalBand) {
                FirstRow = firstRow;
                LastRow = lastRow;
                FirstColumn = firstColumn;
                LastColumn = lastColumn;
                NoHorizontalBand = noHorizontalBand;
                NoVerticalBand = noVerticalBand;
            }

            internal static LegacyDocTableLook Empty { get; } = new LegacyDocTableLook(false, false, false, false, false, false);

            internal bool FirstRow { get; }

            internal bool LastRow { get; }

            internal bool FirstColumn { get; }

            internal bool LastColumn { get; }

            internal bool NoHorizontalBand { get; }

            internal bool NoVerticalBand { get; }
        }
    }
}
