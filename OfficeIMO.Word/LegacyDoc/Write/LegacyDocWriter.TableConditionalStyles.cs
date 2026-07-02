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
        private const int SupportedTableLookMask = TableLookFirstRow
            | TableLookLastRow
            | TableLookFirstColumn
            | TableLookLastColumn
            | TableLookNoHorizontalBand
            | TableLookNoVerticalBand;

        private static LegacyDocTableConditionalStyleSet ReadSupportedTableConditionalStyles(TableStyle? tableStyle, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            Style? style = ResolveSupportedTableStyle(tableStyle, tableStyleDefinitions);
            if (style == null) {
                return LegacyDocTableConditionalStyleSet.Empty;
            }

            int rowBandSize = ReadSupportedTableStyleOwnRowBandSize(style)
                ?? ReadSupportedTableStyleBaseRowBandSize(style, tableStyleDefinitions)
                ?? 1;
            int columnBandSize = ReadSupportedTableStyleOwnColumnBandSize(style)
                ?? ReadSupportedTableStyleBaseColumnBandSize(style, tableStyleDefinitions)
                ?? 1;
            var conditionalStyles = new List<LegacyDocTableConditionalStyle>();
            AppendSupportedTableConditionalStyles(style, conditionalStyles);
            AppendSupportedTableStyleBaseConditionalStyles(style, tableStyleDefinitions, conditionalStyles);

            LegacyDocTableConditionalStyle[] orderedConditionalStyles = conditionalStyles
                .Select((style, index) => new { Style = style, Index = index })
                .OrderBy(item => GetTableConditionalStylePrecedence(item.Style.Type))
                .ThenBy(item => item.Index)
                .Select(item => item.Style)
                .ToArray();

            return orderedConditionalStyles.Length == 0
                ? new LegacyDocTableConditionalStyleSet(Array.Empty<LegacyDocTableConditionalStyle>(), rowBandSize, columnBandSize)
                : new LegacyDocTableConditionalStyleSet(orderedConditionalStyles, rowBandSize, columnBandSize);
        }

        private static void AppendSupportedTableStyleBaseConditionalStyles(Style style, IReadOnlyDictionary<string, Style> tableStyleDefinitions, List<LegacyDocTableConditionalStyle> conditionalStyles) {
            AppendSupportedTableStyleBaseConditionalStyles(style, tableStyleDefinitions, conditionalStyles, new HashSet<string>(StringComparer.OrdinalIgnoreCase));
        }

        private static void AppendSupportedTableStyleBaseConditionalStyles(Style style, IReadOnlyDictionary<string, Style> tableStyleDefinitions, List<LegacyDocTableConditionalStyle> conditionalStyles, ISet<string> visitedStyleIds) {
            string? baseStyleId = style.GetFirstChild<BasedOn>()?.Val?.Value;
            if (IsNoOpTableStyle(baseStyleId) || IsTableGridStyle(baseStyleId)) {
                return;
            }

            if (string.IsNullOrWhiteSpace(baseStyleId)
                || !tableStyleDefinitions.TryGetValue(baseStyleId!, out Style? baseStyle)) {
                return;
            }

            string currentStyleId = style.StyleId?.Value ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(currentStyleId) && !visitedStyleIds.Add(currentStyleId)) {
                throw new NotSupportedException($"Native DOC saving cannot write table style '{currentStyleId}' because its basedOn chain contains a cycle.");
            }

            ThrowIfUnsupportedInheritedTableStyleBase(currentStyleId, baseStyleId!, baseStyle, tableStyleDefinitions, visitedStyleIds);
            AppendSupportedTableConditionalStyles(baseStyle, conditionalStyles);
            AppendSupportedTableStyleBaseConditionalStyles(baseStyle, tableStyleDefinitions, conditionalStyles, visitedStyleIds);
        }

        private static void AppendSupportedTableConditionalStyles(Style style, List<LegacyDocTableConditionalStyle> conditionalStyles) {
            foreach (TableStyleProperties properties in style.Elements<TableStyleProperties>()) {
                TableStyleOverrideValues? type = properties.Type?.Value;
                if (type == null) {
                    throw new NotSupportedException($"Native DOC saving supports table style '{style.StyleId?.Value}' conditional formatting only when the conditional type is specified.");
                }

                TableStyleConditionalFormattingTableProperties? tableProperties = properties.GetFirstChild<TableStyleConditionalFormattingTableProperties>();
                TableStyleConditionalFormattingTableRowProperties? rowProperties = properties.GetFirstChild<TableStyleConditionalFormattingTableRowProperties>();
                TableStyleConditionalFormattingTableCellProperties? cellProperties = properties.GetFirstChild<TableStyleConditionalFormattingTableCellProperties>();
                LegacyDocTableCellShading tableShading = ReadSupportedConditionalTableStyleShading(tableProperties?.GetFirstChild<Shading>());
                LegacyDocTableBorders tableBorders = ReadSupportedConditionalTableStyleBorders(tableProperties?.GetFirstChild<TableBorders>());
                LegacyDocWritableTableRowFormatting rowFormatting = ReadSupportedConditionalTableStyleRowProperties(rowProperties, type.Value);
                LegacyDocTableCellVerticalAlignment? cellVerticalAlignment = ReadSupportedTableCellVerticalAlignment(cellProperties?.GetFirstChild<TableCellVerticalAlignment>());
                LegacyDocTableCellTextDirection? cellTextDirection = ReadSupportedTableCellTextDirection(cellProperties?.GetFirstChild<TextDirection>());
                bool? cellFitText = ReadSupportedTableCellFitText(cellProperties?.GetFirstChild<TableCellFitText>());
                bool? cellNoWrap = ReadSupportedTableCellNoWrap(cellProperties?.GetFirstChild<NoWrap>());
                bool? cellHideMark = ReadSupportedTableCellHideMark(cellProperties?.GetFirstChild<HideMark>());
                LegacyDocTableCellMargins cellMargins = ReadSupportedTableCellMargins(cellProperties?.GetFirstChild<TableCellMargin>());
                LegacyDocTableCellShading cellShading = ReadSupportedConditionalTableStyleShading(cellProperties?.GetFirstChild<Shading>());
                LegacyDocTableCellBorders cellBorders = ReadSupportedConditionalTableStyleBorders(cellProperties);
                LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedStyleParagraphFormatting(properties.GetFirstChild<StyleParagraphProperties>());
                LegacyDocWritableFormatting runFormatting = ReadSupportedRunFormatting(GetSupportedTableStyleRunProperties(properties));
                if (tableShading.HasAny
                    || tableBorders.HasAny
                    || rowFormatting.HasFormatting
                    || cellVerticalAlignment != null
                    || cellTextDirection != null
                    || cellFitText != null
                    || cellNoWrap != null
                    || cellHideMark != null
                    || cellMargins.HasAny
                    || cellShading.HasAny
                    || cellBorders.HasAny
                    || paragraphFormatting.HasFormatting
                    || runFormatting.HasFormatting) {
                    conditionalStyles.Add(new LegacyDocTableConditionalStyle(type.Value, tableShading, tableBorders, rowFormatting, cellVerticalAlignment, cellTextDirection, cellFitText, cellNoWrap, cellHideMark, cellMargins, cellShading, cellBorders, paragraphFormatting, runFormatting));
                }
            }
        }

        private static int GetTableConditionalStylePrecedence(TableStyleOverrideValues type) {
            if (type == TableStyleOverrideValues.NorthWestCell
                || type == TableStyleOverrideValues.NorthEastCell
                || type == TableStyleOverrideValues.SouthWestCell
                || type == TableStyleOverrideValues.SouthEastCell) {
                return 0;
            }

            if (type == TableStyleOverrideValues.FirstRow
                || type == TableStyleOverrideValues.LastRow
                || type == TableStyleOverrideValues.FirstColumn
                || type == TableStyleOverrideValues.LastColumn) {
                return 1;
            }

            if (type == TableStyleOverrideValues.Band1Horizontal
                || type == TableStyleOverrideValues.Band2Horizontal
                || type == TableStyleOverrideValues.Band1Vertical
                || type == TableStyleOverrideValues.Band2Vertical) {
                return 2;
            }

            if (type == TableStyleOverrideValues.WholeTable) {
                return 3;
            }

            return 3;
        }

        private static OpenXmlCompositeElement? GetSupportedTableStyleRunProperties(TableStyleProperties properties) {
            OpenXmlCompositeElement? runProperties = properties.GetFirstChild<StyleRunProperties>();
            return runProperties ?? properties.GetFirstChild<RunPropertiesBaseStyle>();
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
            if (!string.IsNullOrWhiteSpace(value)) {
                if (!int.TryParse(value, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out int parsed)) {
                    throw new NotSupportedException($"Native DOC saving supports table look masks only as hexadecimal values. Unsupported table look value: {value}.");
                }

                mask = parsed;
            }

            ApplyExpandedTableLookFlag(tableLook.FirstRow, TableLookFirstRow, ref mask);
            ApplyExpandedTableLookFlag(tableLook.LastRow, TableLookLastRow, ref mask);
            ApplyExpandedTableLookFlag(tableLook.FirstColumn, TableLookFirstColumn, ref mask);
            ApplyExpandedTableLookFlag(tableLook.LastColumn, TableLookLastColumn, ref mask);
            ApplyExpandedTableLookFlag(tableLook.NoHorizontalBand, TableLookNoHorizontalBand, ref mask);
            ApplyExpandedTableLookFlag(tableLook.NoVerticalBand, TableLookNoVerticalBand, ref mask);

            int unsupportedMask = mask & ~SupportedTableLookMask;
            if (unsupportedMask != 0) {
                throw new NotSupportedException($"Native DOC saving supports table look masks only with first/last row, first/last column, and banding flags. Unsupported table look mask bits: 0x{unsupportedMask:X4}.");
            }

            return new LegacyDocTableLook(
                (mask & TableLookFirstRow) == TableLookFirstRow,
                (mask & TableLookLastRow) == TableLookLastRow,
                (mask & TableLookFirstColumn) == TableLookFirstColumn,
                (mask & TableLookLastColumn) == TableLookLastColumn,
                (mask & TableLookNoHorizontalBand) == TableLookNoHorizontalBand,
                (mask & TableLookNoVerticalBand) == TableLookNoVerticalBand);
        }

        private static LegacyDocWritableTableRowFormatting ReadSupportedConditionalTableStyleRowProperties(TableStyleConditionalFormattingTableRowProperties? rowProperties, TableStyleOverrideValues type) {
            if (rowProperties == null) {
                return LegacyDocWritableTableRowFormatting.Empty;
            }

            if (!IsSupportedTableStyleConditionalRowType(type)) {
                throw new NotSupportedException($"Native DOC saving supports conditional table style row formatting only for whole-table, first/last row, and horizontal band regions. Unsupported conditional row type: {type}.");
            }

            int? rowHeightTwips = null;
            bool rowHeightIsExact = false;
            bool? rowCantSplit = null;
            bool? rowIsHeader = null;
            bool hasRowHeight = false;
            bool hasCantSplit = false;
            bool hasTableHeader = false;
            foreach (OpenXmlElement property in rowProperties.ChildElements) {
                switch (property) {
                    case TableRowHeight rowHeight:
                        if (hasRowHeight) {
                            throw new NotSupportedException("Native DOC saving supports conditional table style row formatting only with one row height.");
                        }

                        ReadSupportedTableRowHeight(rowHeight, out rowHeightTwips, out rowHeightIsExact);
                        hasRowHeight = true;
                        break;
                    case CantSplit cantSplit:
                        if (hasCantSplit) {
                            throw new NotSupportedException("Native DOC saving supports conditional table style row formatting only with one row no-split flag.");
                        }

                        rowCantSplit = ReadTableRowOnOff(cantSplit);
                        hasCantSplit = true;
                        break;
                    case TableHeader tableHeader:
                        if (hasTableHeader) {
                            throw new NotSupportedException("Native DOC saving supports conditional table style row formatting only with one row header flag.");
                        }

                        rowIsHeader = ReadTableRowOnOff(tableHeader);
                        hasTableHeader = true;
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports conditional table style row formatting only with row height, no-split, and header flags. Unsupported conditional row property: {property.LocalName}.");
                }
            }

            return new LegacyDocWritableTableRowFormatting(rowHeightTwips, rowHeightIsExact, rowCantSplit, rowIsHeader);
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
            return ReadSupportedConditionalTableStyleShading(shading);
        }

        private static LegacyDocTableCellShading ReadSupportedConditionalTableStyleShading(Shading? shading) {
            return shading == null ? default : ReadSupportedTableCellShading(shading, "conditional table style shading");
        }

        private static LegacyDocTableBorders ReadSupportedConditionalTableStyleBorders(TableBorders? borders) {
            return borders == null ? default : ReadSupportedTableBorders(borders);
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

                    if (!cell.Shading.HasAny && conditionalStyle.CellShading.HasAny) {
                        cell = cell.WithShading(conditionalStyle.CellShading);
                    }

                    if (!cell.Shading.HasAny && conditionalStyle.TableShading.HasAny) {
                        cell = cell.WithShading(conditionalStyle.TableShading);
                    }

                    if (cell.VerticalAlignment == LegacyDocTableCellVerticalAlignment.Top && conditionalStyle.CellVerticalAlignment != null) {
                        cell = cell.WithVerticalAlignment(conditionalStyle.CellVerticalAlignment.Value);
                    }

                    if (cell.TextDirection == LegacyDocTableCellTextDirection.LeftToRightTopToBottom && conditionalStyle.CellTextDirection != null) {
                        cell = cell.WithTextDirection(conditionalStyle.CellTextDirection.Value);
                    }

                    if (!cell.FitText && conditionalStyle.CellFitText == true) {
                        cell = cell.WithFitText(true);
                    }

                    if (!cell.NoWrap && conditionalStyle.CellNoWrap == true) {
                        cell = cell.WithNoWrap(true);
                    }

                    if (!cell.HideMark && conditionalStyle.CellHideMark == true) {
                        cell = cell.WithHideMark(true);
                    }

                    if (conditionalStyle.CellMargins.HasAny) {
                        cell = cell.WithMargins(MergeSupportedConditionalTableCellMargins(cell.Margins, conditionalStyle.CellMargins));
                    }

                    if (conditionalStyle.CellBorders.HasAny) {
                        cell = cell.WithBorders(MergeSupportedTableCellBorders(cell.Borders, conditionalStyle.CellBorders));
                    }

                    if (conditionalStyle.TableBorders.HasAny) {
                        LegacyDocTableCellBorders regionBorders = CreateSupportedConditionalTableBorders(conditionalStyle.Type, conditionalStyle.TableBorders, tableLook, conditionalStyles.RowBandSize, conditionalStyles.ColumnBandSize, rowIndex, rowCount, columnIndex, writableCells.Count);
                        cell = cell.WithBorders(MergeSupportedTableCellBorders(cell.Borders, regionBorders));
                    }

                    if (conditionalStyle.ParagraphFormatting.HasFormatting) {
                        cell = cell.WithParagraphFormatting(conditionalStyle.ParagraphFormatting.WithInheritedParagraphFormatting(cell.ParagraphFormatting));
                    }

                    if (conditionalStyle.RunFormatting.HasFormatting) {
                        cell = cell.WithRunFormatting(conditionalStyle.RunFormatting.WithInheritedFormatting(cell.RunFormatting));
                    }
                }

                styledCells[columnIndex] = cell;
            }

            return styledCells;
        }

        private static LegacyDocWritableTableRowFormatting ApplySupportedTableConditionalRowFormatting(
            LegacyDocWritableTableRowFormatting rowFormatting,
            LegacyDocTableConditionalStyleSet conditionalStyles,
            LegacyDocTableLook tableLook,
            int rowIndex,
            int rowCount) {
            if (!conditionalStyles.HasAny) {
                return rowFormatting;
            }

            foreach (LegacyDocTableConditionalStyle conditionalStyle in conditionalStyles.Styles) {
                if (!conditionalStyle.RowFormatting.HasFormatting
                    || !AppliesToRow(conditionalStyle.Type, tableLook, conditionalStyles.RowBandSize, rowIndex, rowCount)) {
                    continue;
                }

                rowFormatting = rowFormatting.WithInheritedRowFormatting(conditionalStyle.RowFormatting);
            }

            return rowFormatting;
        }

        private static bool AppliesToRow(
            TableStyleOverrideValues type,
            LegacyDocTableLook tableLook,
            int rowBandSize,
            int rowIndex,
            int rowCount) {
            if (rowIndex < 0 || rowIndex >= rowCount) {
                return false;
            }

            if (type == TableStyleOverrideValues.WholeTable) {
                return true;
            }

            if (type == TableStyleOverrideValues.FirstRow) {
                return tableLook.FirstRow && rowIndex == 0;
            }

            if (type == TableStyleOverrideValues.LastRow) {
                return tableLook.LastRow && rowIndex + 1 == rowCount;
            }

            if (type == TableStyleOverrideValues.Band1Horizontal || type == TableStyleOverrideValues.Band2Horizontal) {
                return !tableLook.NoHorizontalBand && TryGetBandType(rowIndex, tableLook.FirstRow, rowBandSize, type == TableStyleOverrideValues.Band1Horizontal);
            }

            return false;
        }

        private static bool IsSupportedTableStyleConditionalRowType(TableStyleOverrideValues type) {
            return type == TableStyleOverrideValues.WholeTable
                || type == TableStyleOverrideValues.FirstRow
                || type == TableStyleOverrideValues.LastRow
                || type == TableStyleOverrideValues.Band1Horizontal
                || type == TableStyleOverrideValues.Band2Horizontal;
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
            if (rowIndex < 0 || rowIndex >= rowCount || columnIndex < 0 || columnIndex >= columnCount) {
                return false;
            }

            if (type == TableStyleOverrideValues.WholeTable) {
                return true;
            }

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

        private static LegacyDocTableCellBorders CreateSupportedConditionalTableBorders(
            TableStyleOverrideValues type,
            LegacyDocTableBorders tableBorders,
            LegacyDocTableLook tableLook,
            int rowBandSize,
            int columnBandSize,
            int rowIndex,
            int rowCount,
            int columnIndex,
            int columnCount) {
            LegacyDocTableCellBorder top = AppliesToCell(type, tableLook, rowBandSize, columnBandSize, rowIndex - 1, rowCount, columnIndex, columnCount)
                ? tableBorders.InsideHorizontal
                : tableBorders.Top;
            LegacyDocTableCellBorder bottom = AppliesToCell(type, tableLook, rowBandSize, columnBandSize, rowIndex + 1, rowCount, columnIndex, columnCount)
                ? tableBorders.InsideHorizontal
                : tableBorders.Bottom;
            LegacyDocTableCellBorder left = AppliesToCell(type, tableLook, rowBandSize, columnBandSize, rowIndex, rowCount, columnIndex - 1, columnCount)
                ? tableBorders.InsideVertical
                : tableBorders.Left;
            LegacyDocTableCellBorder right = AppliesToCell(type, tableLook, rowBandSize, columnBandSize, rowIndex, rowCount, columnIndex + 1, columnCount)
                ? tableBorders.InsideVertical
                : tableBorders.Right;

            return new LegacyDocTableCellBorders(top, left, bottom, right);
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

        private static LegacyDocTableCellMargins MergeSupportedConditionalTableCellMargins(LegacyDocTableCellMargins cellMargins, LegacyDocTableCellMargins inheritedMargins) {
            return new LegacyDocTableCellMargins(
                cellMargins.TopTwips ?? inheritedMargins.TopTwips,
                cellMargins.RightTwips ?? inheritedMargins.RightTwips,
                cellMargins.BottomTwips ?? inheritedMargins.BottomTwips,
                cellMargins.LeftTwips ?? inheritedMargins.LeftTwips);
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
            internal LegacyDocTableConditionalStyle(
                TableStyleOverrideValues type,
                LegacyDocTableCellShading tableShading,
                LegacyDocTableBorders tableBorders,
                LegacyDocWritableTableRowFormatting rowFormatting,
                LegacyDocTableCellVerticalAlignment? cellVerticalAlignment,
                LegacyDocTableCellTextDirection? cellTextDirection,
                bool? cellFitText,
                bool? cellNoWrap,
                bool? cellHideMark,
                LegacyDocTableCellMargins cellMargins,
                LegacyDocTableCellShading cellShading,
                LegacyDocTableCellBorders cellBorders,
                LegacyDocWritableParagraphFormatting paragraphFormatting,
                LegacyDocWritableFormatting runFormatting) {
                Type = type;
                TableShading = tableShading;
                TableBorders = tableBorders;
                RowFormatting = rowFormatting;
                CellVerticalAlignment = cellVerticalAlignment;
                CellTextDirection = cellTextDirection;
                CellFitText = cellFitText;
                CellNoWrap = cellNoWrap;
                CellHideMark = cellHideMark;
                CellMargins = cellMargins;
                CellShading = cellShading;
                CellBorders = cellBorders;
                ParagraphFormatting = paragraphFormatting;
                RunFormatting = runFormatting;
            }

            internal TableStyleOverrideValues Type { get; }

            internal LegacyDocTableCellShading TableShading { get; }

            internal LegacyDocTableBorders TableBorders { get; }

            internal LegacyDocWritableTableRowFormatting RowFormatting { get; }

            internal LegacyDocTableCellVerticalAlignment? CellVerticalAlignment { get; }

            internal LegacyDocTableCellTextDirection? CellTextDirection { get; }

            internal bool? CellFitText { get; }

            internal bool? CellNoWrap { get; }

            internal bool? CellHideMark { get; }

            internal LegacyDocTableCellMargins CellMargins { get; }

            internal LegacyDocTableCellShading CellShading { get; }

            internal LegacyDocTableCellBorders CellBorders { get; }

            internal LegacyDocWritableParagraphFormatting ParagraphFormatting { get; }

            internal LegacyDocWritableFormatting RunFormatting { get; }
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
