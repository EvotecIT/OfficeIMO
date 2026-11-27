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

        private static LegacyDocTableConditionalStyleSet ReadSupportedTableConditionalStyles(string? tableStyleId, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            LegacyDocElement style = ResolveSupportedTableStyle(tableStyleId, tableStyleDefinitions);
            if (style.IsNone) {
                return LegacyDocTableConditionalStyleSet.Empty;
            }

            int rowBandSize = ReadSupportedTableStyleOwnRowBandSize(style)
                ?? ReadSupportedTableStyleBaseRowBandSize(style, tableStyleDefinitions)
                ?? 0;
            int columnBandSize = ReadSupportedTableStyleOwnColumnBandSize(style)
                ?? ReadSupportedTableStyleBaseColumnBandSize(style, tableStyleDefinitions)
                ?? 0;
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

        private static void AppendSupportedTableStyleBaseConditionalStyles(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions, List<LegacyDocTableConditionalStyle> conditionalStyles) {
            AppendSupportedTableStyleBaseConditionalStyles(style, tableStyleDefinitions, conditionalStyles, new HashSet<string>(StringComparer.Ordinal));
        }

        private static void AppendSupportedTableStyleBaseConditionalStyles(LegacyDocElement style, LegacyDocTableStyleDefinitions tableStyleDefinitions, List<LegacyDocTableConditionalStyle> conditionalStyles, ISet<string> visitedStyleIds) {
            string? baseStyleId = style.Child("basedOn").Val;
            if (IsNoOpTableStyle(baseStyleId)) {
                return;
            }

            if (string.IsNullOrWhiteSpace(baseStyleId)
                || !TryResolveTableStyleDefinition(baseStyleId!, tableStyleDefinitions, out LegacyDocElement baseStyle)) {
                return;
            }

            string currentStyleId = LegacyDocStyleCatalog.StyleId(style) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(currentStyleId) && !visitedStyleIds.Add(currentStyleId)) {
                throw new NotSupportedException($"Native DOC saving cannot write table style '{currentStyleId}' because its basedOn chain contains a cycle.");
            }

            ThrowIfUnsupportedInheritedTableStyleBase(currentStyleId, baseStyleId!, baseStyle, tableStyleDefinitions, visitedStyleIds);
            AppendSupportedTableConditionalStyles(baseStyle, conditionalStyles);
            AppendSupportedTableStyleBaseConditionalStyles(baseStyle, tableStyleDefinitions, conditionalStyles, visitedStyleIds);
        }

        private static void AppendSupportedTableConditionalStyles(LegacyDocElement style, List<LegacyDocTableConditionalStyle> conditionalStyles) {
            foreach (LegacyDocElement properties in style.Elements("tblStylePr")) {
                string? type = properties.Attr("type");
                if (type == null) {
                    throw new NotSupportedException($"Native DOC saving supports table style '{LegacyDocStyleCatalog.StyleId(style)}' conditional formatting only when the conditional type is specified.");
                }

                // Word discards wholeTable conditional blocks; whole-table defaults belong on the style itself.
                if (type == "wholeTable") continue;
                ThrowIfUnknownTableStyleConditionalType(LegacyDocStyleCatalog.StyleId(style), type);

                LegacyDocElement tableProperties = properties.Child("tblPr");
                LegacyDocElement rowProperties = properties.Child("trPr");
                LegacyDocElement cellProperties = properties.Child("tcPr");
                LegacyDocTableCellShading tableShading = ReadSupportedConditionalTableStyleShading(tableProperties.Child("shd"));
                LegacyDocTableBorders tableBorders = ReadSupportedConditionalTableStyleBorders(tableProperties.Child("tblBorders"));
                LegacyDocWritableTableRowFormatting rowFormatting = ReadSupportedConditionalTableStyleRowProperties(rowProperties, type);
                LegacyDocTableCellVerticalAlignment? cellVerticalAlignment = ReadSupportedTableCellVerticalAlignment(cellProperties.Child("vAlign"));
                LegacyDocTableCellTextDirection? cellTextDirection = ReadSupportedTableCellTextDirection(cellProperties.Child("textDirection"));
                bool? cellFitText = ReadSupportedTableCellFitText(cellProperties.Child("tcFitText"));
                bool? cellNoWrap = ReadSupportedTableCellNoWrap(cellProperties.Child("noWrap"));
                bool? cellHideMark = ReadSupportedTableCellHideMark(cellProperties.Child("hideMark"));
                LegacyDocTableCellMargins cellMargins = ReadSupportedTableCellMargins(cellProperties.Child("tcMar"));
                LegacyDocTableCellShading cellShading = ReadSupportedConditionalTableStyleShading(cellProperties.Child("shd"));
                LegacyDocTableCellBorders cellBorders = ReadSupportedConditionalTableStyleCellBorders(cellProperties);
                LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedStyleParagraphFormatting(properties.Child("pPr"));
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
                    || cellShading.IsSpecified
                    || cellBorders.HasAny
                    || paragraphFormatting.HasFormatting
                    || runFormatting.HasFormatting) {
                    conditionalStyles.Add(new LegacyDocTableConditionalStyle(type, tableShading, tableBorders, rowFormatting, cellVerticalAlignment, cellTextDirection, cellFitText, cellNoWrap, cellHideMark, cellMargins, cellShading, cellBorders, paragraphFormatting, runFormatting));
                }
            }
        }

        private static int GetTableConditionalStylePrecedence(string type) {
            // Higher-priority regions are visited first, independently of XML or base-style order.
            if (type == "nwCell"
                || type == "neCell"
                || type == "swCell"
                || type == "seCell") {
                return 0;
            }

            if (type == "firstRow" || type == "lastRow") {
                return 1;
            }

            if (type == "firstCol" || type == "lastCol") {
                return 2;
            }

            if (type == "band1Vert" || type == "band2Vert") {
                return 3;
            }

            if (type == "band1Horz" || type == "band2Horz") {
                return 4;
            }

            return 5;
        }

        private static int ReadSupportedTableStyleBandSize(string? value, string axisName) {
            if (value == null) {
                return 0;
            }

            int bandSize = LegacyDocValues.XmlInt32(value);
            if (bandSize < 0 || bandSize > byte.MaxValue) {
                throw new NotSupportedException($"Native DOC saving supports table style {axisName} band sizes only as non-negative values within the DOC table column limit.");
            }

            return bandSize;
        }

        private static LegacyDocTableLook ReadSupportedTableLook(LegacyDocElement tableLook) {
            string? value = tableLook.Val;
            if (tableLook.Exists && !LegacyDocWordHelpers.HasExpandedTableLookFlags(tableLook) && !string.IsNullOrWhiteSpace(value)
                && !int.TryParse(value, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out _)) {
                throw new NotSupportedException($"Native DOC saving supports table look masks only as hexadecimal values. Unsupported table look value: {value}.");
            }
            int mask = LegacyDocWordHelpers.GetTableLookMask(tableLook);

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

        private static LegacyDocWritableTableRowFormatting ReadSupportedConditionalTableStyleRowProperties(LegacyDocElement rowProperties, string type) {
            if (rowProperties.IsNone) {
                return LegacyDocWritableTableRowFormatting.Empty;
            }

            if (!IsSupportedTableStyleConditionalRowType(type)) {
                throw new NotSupportedException($"Native DOC saving supports conditional table style row formatting only for whole-table, first/last row, and horizontal band regions. Unsupported conditional row type: {FormatSdkEnumValue(type, "TableStyleOverrideValues")}.");
            }

            int? rowHeightTwips = null;
            bool rowHeightIsExact = false;
            bool? rowCantSplit = null;
            bool? rowIsHeader = null;
            bool hasRowHeight = false;
            bool hasCantSplit = false;
            bool hasTableHeader = false;
            foreach (LegacyDocElement property in rowProperties.Elements()) {
                switch (property.WName) {
                    case "trHeight":
                        if (hasRowHeight) {
                            throw new NotSupportedException("Native DOC saving supports conditional table style row formatting only with one row height.");
                        }

                        ReadSupportedTableRowHeight(property, out rowHeightTwips, out rowHeightIsExact);
                        hasRowHeight = true;
                        break;
                    case "cantSplit":
                        if (hasCantSplit) {
                            throw new NotSupportedException("Native DOC saving supports conditional table style row formatting only with one row no-split flag.");
                        }

                        rowCantSplit = ReadTableRowOnOff(property);
                        hasCantSplit = true;
                        break;
                    case "tblHeader":
                        if (hasTableHeader) {
                            throw new NotSupportedException("Native DOC saving supports conditional table style row formatting only with one row header flag.");
                        }

                        rowIsHeader = ReadTableRowOnOff(property);
                        hasTableHeader = true;
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports conditional table style row formatting only with row height, no-split, and header flags. Unsupported conditional row property: {property.LocalName}.");
                }
            }

            return new LegacyDocWritableTableRowFormatting(rowHeightTwips, rowHeightIsExact, rowCantSplit, rowIsHeader);
        }

        private static LegacyDocTableCellShading ReadSupportedConditionalTableStyleCellShading(LegacyDocElement cellProperties) {
            LegacyDocElement shading = cellProperties.Child("shd");
            return ReadSupportedConditionalTableStyleShading(shading);
        }

        private static LegacyDocTableCellShading ReadSupportedConditionalTableStyleShading(LegacyDocElement shading) {
            return shading.IsNone ? default : ReadSupportedTableCellShading(shading, "conditional table style shading");
        }

        private static LegacyDocTableBorders ReadSupportedConditionalTableStyleBorders(LegacyDocElement borders) {
            return borders.IsNone ? default : ReadSupportedTableBorders(borders);
        }

        private static LegacyDocTableCellBorders ReadSupportedConditionalTableStyleCellBorders(LegacyDocElement cellProperties) {
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

        private static IReadOnlyList<LegacyDocWritableTableCell> ApplySupportedTableConditionalStyles(
            IReadOnlyList<LegacyDocWritableTableCell> writableCells,
            LegacyDocTableConditionalStyleSet conditionalStyles,
            LegacyDocTableLook tableLook,
            int rowIndex,
            int rowCount) {
            if (!conditionalStyles.HasAny || writableCells.Count == 0) {
                return writableCells;
            }

            // Conditions with equal precedence can interleave across the style
            // chain. Resolve each condition's derived/base inheritance before
            // combining applicable conditions for a cell.
            var paragraphConditions = conditionalStyles.Styles
                .Where(style => style.ParagraphFormatting.HasFormatting)
                .GroupBy(style => style.Type)
                .Select(group => new {
                    Type = group.Key,
                    Formatting = group.Select(style => style.ParagraphFormatting)
                        .Aggregate((derived, inherited) => derived.WithInheritedParagraphFormatting(inherited))
                })
                .ToArray();
            var runConditions = conditionalStyles.Styles
                .Where(style => style.RunFormatting.HasFormatting)
                .GroupBy(style => style.Type)
                .Select(group => new {
                    Type = group.Key,
                    Formatting = group.Select(style => style.RunFormatting)
                        .Aggregate((derived, inherited) => derived.WithInheritedFormatting(inherited))
                })
                .ToArray();
            var styledCells = new LegacyDocWritableTableCell[writableCells.Count];
            for (int columnIndex = 0; columnIndex < writableCells.Count; columnIndex++) {
                LegacyDocWritableTableCell cell = writableCells[columnIndex];
                LegacyDocWritableParagraphFormatting originalParagraphFormatting = cell.ParagraphFormatting;
                LegacyDocWritableParagraphFormatting conditionalParagraphFormatting = LegacyDocWritableParagraphFormatting.Plain;
                LegacyDocWritableFormatting conditionalRunFormatting = LegacyDocWritableFormatting.Plain;
                foreach (LegacyDocTableConditionalStyle conditionalStyle in conditionalStyles.Styles) {
                    if (!AppliesToCell(conditionalStyle.Type, tableLook, conditionalStyles.RowBandSize, conditionalStyles.ColumnBandSize, rowIndex, rowCount, columnIndex, writableCells.Count)) {
                        continue;
                    }

                    if (!cell.Shading.IsSpecified && conditionalStyle.CellShading.IsSpecified) {
                        cell = cell.WithShading(conditionalStyle.CellShading);
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

                }

                foreach (var condition in paragraphConditions) {
                    if (AppliesToCell(condition.Type, tableLook, conditionalStyles.RowBandSize, conditionalStyles.ColumnBandSize, rowIndex, rowCount, columnIndex, writableCells.Count)) {
                        conditionalParagraphFormatting = conditionalParagraphFormatting.WithInheritedParagraphFormatting(condition.Formatting);
                    }
                }
                foreach (var condition in runConditions) {
                    if (AppliesToCell(condition.Type, tableLook, conditionalStyles.RowBandSize, conditionalStyles.ColumnBandSize, rowIndex, rowCount, columnIndex, writableCells.Count)) {
                        conditionalRunFormatting = conditionalRunFormatting.WithInheritedFormatting(condition.Formatting);
                    }
                }
                if (conditionalRunFormatting.HasFormatting) {
                    cell = cell.WithRunFormatting(conditionalRunFormatting.WithInheritedFormatting(cell.RunFormatting));
                }
                if (conditionalParagraphFormatting.HasFormatting) {
                    cell = cell.WithParagraphFormatting(conditionalParagraphFormatting.WithInheritedParagraphFormatting(originalParagraphFormatting));
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
            string type,
            LegacyDocTableLook tableLook,
            int rowBandSize,
            int rowIndex,
            int rowCount) {
            if (rowIndex < 0 || rowIndex >= rowCount) {
                return false;
            }

            if (type == "wholeTable") {
                return true;
            }

            if (type == "firstRow") {
                return tableLook.FirstRow && rowIndex == 0;
            }

            if (type == "lastRow") {
                return tableLook.LastRow && rowIndex + 1 == rowCount;
            }

            if (type == "band1Horz" || type == "band2Horz") {
                return rowBandSize > 0 && !tableLook.NoHorizontalBand && TryGetBandType(rowIndex, tableLook.FirstRow, rowBandSize, type == "band1Horz");
            }

            return false;
        }

        private static bool IsSupportedTableStyleConditionalRowType(string type) {
            return type == "wholeTable"
                || type == "firstRow"
                || type == "lastRow"
                || type == "band1Horz"
                || type == "band2Horz";
        }

        private static bool AppliesToCell(
            string type,
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

            if (type == "wholeTable") {
                return true;
            }

            if (type == "firstRow") {
                return tableLook.FirstRow && rowIndex == 0;
            }

            if (type == "lastRow") {
                return tableLook.LastRow && rowIndex + 1 == rowCount;
            }

            if (type == "firstCol") {
                return tableLook.FirstColumn && columnIndex == 0;
            }

            if (type == "lastCol") {
                return tableLook.LastColumn && columnIndex + 1 == columnCount;
            }

            if (type == "nwCell") {
                return tableLook.FirstRow && tableLook.FirstColumn && rowIndex == 0 && columnIndex == 0;
            }

            if (type == "neCell") {
                return tableLook.FirstRow && tableLook.LastColumn && rowIndex == 0 && columnIndex + 1 == columnCount;
            }

            if (type == "swCell") {
                return tableLook.LastRow && tableLook.FirstColumn && rowIndex + 1 == rowCount && columnIndex == 0;
            }

            if (type == "seCell") {
                return tableLook.LastRow && tableLook.LastColumn && rowIndex + 1 == rowCount && columnIndex + 1 == columnCount;
            }

            if (type == "band1Horz" || type == "band2Horz") {
                return rowBandSize > 0 && !tableLook.NoHorizontalBand && TryGetBandType(rowIndex, tableLook.FirstRow, rowBandSize, type == "band1Horz");
            }

            if (type == "band1Vert" || type == "band2Vert") {
                return columnBandSize > 0 && !tableLook.NoVerticalBand && TryGetBandType(columnIndex, tableLook.FirstColumn, columnBandSize, type == "band1Vert");
            }

            throw new NotSupportedException($"Native DOC saving does not support table style conditional type '{FormatSdkEnumValue(type, "TableStyleOverrideValues")}'.");
        }

        private static LegacyDocTableCellBorders CreateSupportedConditionalTableBorders(
            string type,
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
                string type,
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

            internal string Type { get; }

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

            internal bool FirstRow { get; }

            internal bool LastRow { get; }

            internal bool FirstColumn { get; }

            internal bool LastColumn { get; }

            internal bool NoHorizontalBand { get; }

            internal bool NoVerticalBand { get; }
        }
    }
}
