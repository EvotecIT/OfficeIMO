using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static IEnumerable<TableStyleProperties> EnumerateApplicableTableConditionalStyleProperties(WordTable table, int rowIndex, int columnIndex, int rowCount, int columnCount) {
            foreach (TableStyleProperties properties in EnumerateApplicableTableCornerStyleProperties(table, rowIndex, columnIndex, rowCount, columnCount)) {
                yield return properties;
            }

            if (table.ConditionalFormattingFirstRow == true && rowIndex == 0) {
                foreach (TableStyleProperties properties in EnumerateTableConditionalStyleProperties(table, TableStyleOverrideValues.FirstRow)) {
                    yield return properties;
                }
            }

            if (table.ConditionalFormattingLastRow == true && rowIndex == rowCount - 1) {
                foreach (TableStyleProperties properties in EnumerateTableConditionalStyleProperties(table, TableStyleOverrideValues.LastRow)) {
                    yield return properties;
                }
            }

            if (table.ConditionalFormattingFirstColumn == true && columnIndex == 0) {
                foreach (TableStyleProperties properties in EnumerateTableConditionalStyleProperties(table, TableStyleOverrideValues.FirstColumn)) {
                    yield return properties;
                }
            }

            if (table.ConditionalFormattingLastColumn == true && columnIndex == columnCount - 1) {
                foreach (TableStyleProperties properties in EnumerateTableConditionalStyleProperties(table, TableStyleOverrideValues.LastColumn)) {
                    yield return properties;
                }
            }

            if (table.ConditionalFormattingNoHorizontalBand != true
                && TryGetBandType(rowIndex, table.ConditionalFormattingFirstRow == true, ResolveTableBandSize(table, rowBand: true), horizontal: true, out TableStyleOverrideValues horizontalBand)) {
                foreach (TableStyleProperties properties in EnumerateTableConditionalStyleProperties(table, horizontalBand)) {
                    yield return properties;
                }
            }

            if (table.ConditionalFormattingNoVerticalBand != true
                && TryGetBandType(columnIndex, table.ConditionalFormattingFirstColumn == true, ResolveTableBandSize(table, rowBand: false), horizontal: false, out TableStyleOverrideValues verticalBand)) {
                foreach (TableStyleProperties properties in EnumerateTableConditionalStyleProperties(table, verticalBand)) {
                    yield return properties;
                }
            }
        }

        private static IEnumerable<TableStyleProperties> EnumerateApplicableTableCornerStyleProperties(WordTable table, int rowIndex, int columnIndex, int rowCount, int columnCount) {
            if (TryGetTableCornerType(table, rowIndex, columnIndex, rowCount, columnCount, out TableStyleOverrideValues cornerType)) {
                foreach (TableStyleProperties properties in EnumerateTableConditionalStyleProperties(table, cornerType)) {
                    yield return properties;
                }
            }
        }

        private static bool TryGetTableCornerType(WordTable table, int rowIndex, int columnIndex, int rowCount, int columnCount, out TableStyleOverrideValues cornerType) {
            cornerType = TableStyleOverrideValues.NorthWestCell;
            bool firstRow = table.ConditionalFormattingFirstRow == true && rowIndex == 0;
            bool lastRow = table.ConditionalFormattingLastRow == true && rowIndex == rowCount - 1;
            bool firstColumn = table.ConditionalFormattingFirstColumn == true && columnIndex == 0;
            bool lastColumn = table.ConditionalFormattingLastColumn == true && columnIndex == columnCount - 1;

            if (firstRow && firstColumn) {
                cornerType = TableStyleOverrideValues.NorthWestCell;
                return true;
            }

            if (firstRow && lastColumn) {
                cornerType = TableStyleOverrideValues.NorthEastCell;
                return true;
            }

            if (lastRow && firstColumn) {
                cornerType = TableStyleOverrideValues.SouthWestCell;
                return true;
            }

            if (lastRow && lastColumn) {
                cornerType = TableStyleOverrideValues.SouthEastCell;
                return true;
            }

            return false;
        }

        private static int ResolveTableBandSize(WordTable table, bool rowBand) {
            foreach (StyleTableProperties properties in EnumerateTableStyleProperties(table)) {
                Int32Value? value = rowBand
                    ? properties.GetFirstChild<TableStyleRowBandSize>()?.Val
                    : properties.GetFirstChild<TableStyleColumnBandSize>()?.Val;
                if (value?.Value > 0) {
                    return (int)value.Value;
                }
            }

            return 1;
        }

        private static bool TryGetBandType(int index, bool skipFirst, int bandSize, bool horizontal, out TableStyleOverrideValues bandType) {
            bandType = horizontal ? TableStyleOverrideValues.Band1Horizontal : TableStyleOverrideValues.Band1Vertical;
            int adjustedIndex = skipFirst ? index - 1 : index;
            if (adjustedIndex < 0) {
                return false;
            }

            int bandIndex = adjustedIndex / Math.Max(1, bandSize);
            bool band1 = bandIndex % 2 == 0;
            bandType = horizontal
                ? (band1 ? TableStyleOverrideValues.Band1Horizontal : TableStyleOverrideValues.Band2Horizontal)
                : (band1 ? TableStyleOverrideValues.Band1Vertical : TableStyleOverrideValues.Band2Vertical);
            return true;
        }
    }
}
