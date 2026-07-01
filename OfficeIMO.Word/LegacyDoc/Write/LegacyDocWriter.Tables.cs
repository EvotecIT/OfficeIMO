using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void AppendTable(StringBuilder text, List<LegacyDocWritableRun> runs, List<LegacyDocWritableParagraph> paragraphFormats, LegacyDocWritableBookmarksBuilder bookmarks, Table table, MainDocumentPart mainPart, IReadOnlyDictionary<string, ushort> styleIndexes, IReadOnlyDictionary<string, Style> tableStyleDefinitions, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes) {
            ThrowIfUnsupportedTableShape(table, tableStyleDefinitions);

            TableRow[] rows = table.Elements<TableRow>().ToArray();
            if (rows.Length == 0) {
                throw new NotSupportedException("Native DOC saving supports simple tables only when at least one row is present.");
            }

            TableProperties? tableProperties = table.GetFirstChild<TableProperties>();
            LegacyDocTableAlignment? tableAlignment = ReadSupportedTableAlignment(tableProperties);
            tableAlignment ??= ReadSupportedTableStyleAlignment(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            int? tableLeftIndentTwips = ReadSupportedTableIndentation(tableProperties);
            tableLeftIndentTwips ??= ReadSupportedTableStyleIndentation(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            LegacyDocTablePreferredWidth? tablePreferredWidth = ReadSupportedTablePreferredWidth(tableProperties);
            tablePreferredWidth ??= ReadSupportedTableStylePreferredWidth(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            bool? tableAutofit = ReadSupportedTableAutofit(tableProperties);
            tableAutofit ??= ReadSupportedTableStyleAutofit(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            LegacyDocTableCellMargins? defaultCellMargins = ReadSupportedTableDefaultCellMargins(tableProperties);
            defaultCellMargins ??= ReadSupportedTableStyleDefaultCellMargins(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            int? defaultCellSpacingTwips = ReadSupportedTableDefaultCellSpacing(tableProperties);
            defaultCellSpacingTwips ??= ReadSupportedTableStyleDefaultCellSpacing(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            LegacyDocTableBorders tableBorders = ReadSupportedTableBorders(tableProperties, tableStyleDefinitions);
            LegacyDocTableCellShading tableShading = ReadSupportedTableShading(tableProperties, tableStyleDefinitions);
            LegacyDocWritableParagraphFormatting tableStyleParagraphFormatting = ReadSupportedTableStyleParagraphFormatting(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            LegacyDocWritableFormatting tableStyleRunFormatting = ReadSupportedTableStyleRunFormatting(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            LegacyDocTableConditionalStyleSet conditionalStyles = ReadSupportedTableConditionalStyles(tableProperties?.GetFirstChild<TableStyle>(), tableStyleDefinitions);
            LegacyDocTableLook tableLook = ReadSupportedTableLook(tableProperties?.GetFirstChild<TableLook>());
            IReadOnlyList<int> gridColumnWidthsTwips = ReadSupportedTableGridWidths(table.GetFirstChild<TableGrid>());
            for (int rowIndex = 0; rowIndex < rows.Length; rowIndex++) {
                TableRow row = rows[rowIndex];
                LegacyDocWritableTableRowFormatting rowFormatting = ReadSupportedTableRowFormatting(row, out TableCell[] cells);
                if (cells.Length == 0) {
                    throw new NotSupportedException("Native DOC saving supports simple tables only when every row contains at least one cell.");
                }

                IReadOnlyList<LegacyDocWritableTableCell> writableCells = ExpandSupportedTableCells(cells, gridColumnWidthsTwips, tableBorders, tableShading, conditionalStyles, tableLook, rowIndex, rows.Length);
                IReadOnlyList<int> cellWidthsTwips = ReadSupportedTableCellWidths(writableCells);
                IReadOnlyList<LegacyDocTableCellHorizontalMerge> cellHorizontalMerges = ReadSupportedTableCellHorizontalMerges(writableCells);
                IReadOnlyList<LegacyDocTableCellVerticalMerge> cellVerticalMerges = ReadSupportedTableCellVerticalMerges(writableCells);
                IReadOnlyList<LegacyDocTableCellVerticalAlignment> cellVerticalAlignments = ReadSupportedTableCellVerticalAlignments(writableCells);
                IReadOnlyList<LegacyDocTableCellTextDirection> cellTextDirections = ReadSupportedTableCellTextDirections(writableCells);
                IReadOnlyList<bool> cellFitTexts = ReadSupportedTableCellFitTexts(writableCells);
                IReadOnlyList<bool> cellNoWraps = ReadSupportedTableCellNoWraps(writableCells);
                IReadOnlyList<bool> cellHideMarks = ReadSupportedTableCellHideMarks(writableCells);
                IReadOnlyList<LegacyDocTableCellMargins> cellMargins = ReadSupportedTableCellMargins(writableCells);
                IReadOnlyList<LegacyDocTableCellShading> cellShadings = ReadSupportedTableCellShadings(writableCells);
                IReadOnlyList<LegacyDocTableCellBorders> cellBorders = ReadSupportedTableCellBorders(writableCells);
                foreach (LegacyDocWritableTableCell writableCell in writableCells) {
                    LegacyDocWritableParagraphFormatting cellParagraphFormatting = writableCell.ParagraphFormatting.WithInheritedParagraphFormatting(tableStyleParagraphFormatting);
                    LegacyDocWritableFormatting cellRunFormatting = writableCell.RunFormatting.WithInheritedFormatting(tableStyleRunFormatting);
                    LegacyDocWritableParagraphFormatting paragraphFormatting = AppendTableCell(text, runs, paragraphFormats, bookmarks, writableCell.SourceCell, mainPart, styleIndexes, cellParagraphFormatting, cellRunFormatting, footnotes, endnotes, out int finalParagraphStart)
                        .WithTableMarkers(isTableTerminatingParagraph: false);
                    text.Append('\a');
                    paragraphFormats.Add(new LegacyDocWritableParagraph(finalParagraphStart, text.Length - finalParagraphStart, paragraphFormatting));
                }

                int rowTerminatorStart = text.Length;
                text.Append('\a');
                paragraphFormats.Add(new LegacyDocWritableParagraph(
                    rowTerminatorStart,
                    1,
                    LegacyDocWritableParagraphFormatting.Plain.WithTableMarkers(
                        isTableTerminatingParagraph: true,
                        tableCellWidthsTwips: cellWidthsTwips,
                        tableRowHeightTwips: rowFormatting.RowHeightTwips,
                        tableRowHeightIsExact: rowFormatting.RowHeightIsExact,
                        tableRowCantSplit: rowFormatting.RowCantSplit,
                        tableRowIsHeader: rowFormatting.RowIsHeader,
                        tableAlignment: tableAlignment,
                        tableLeftIndentTwips: tableLeftIndentTwips,
                        tableCellHorizontalMerges: cellHorizontalMerges,
                        tableCellVerticalMerges: cellVerticalMerges,
                        tableCellVerticalAlignments: cellVerticalAlignments,
                        tableCellTextDirections: cellTextDirections,
                        tableCellFitTexts: cellFitTexts,
                        tableCellNoWraps: cellNoWraps,
                        tableCellHideMarks: cellHideMarks,
                        tableCellMargins: cellMargins,
                        tableCellShadings: cellShadings,
                        tableCellBorders: cellBorders,
                        defaultTableCellMargins: defaultCellMargins,
                        defaultTableCellSpacingTwips: defaultCellSpacingTwips,
                        tablePreferredWidth: tablePreferredWidth,
                        tableAutofit: tableAutofit)));
            }

            text.Append('\r');
        }

        private static void ThrowIfUnsupportedTableShape(Table table, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            foreach (OpenXmlElement child in table.ChildElements) {
                switch (child) {
                    case TableProperties tableProperties:
                        ThrowIfUnsupportedTableProperties(tableProperties, tableStyleDefinitions);
                        break;
                    case TableGrid tableGrid:
                        ThrowIfUnsupportedTableGrid(tableGrid);
                        break;
                    case TableRow:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table element: {child.LocalName}.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableProperties(TableProperties tableProperties, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            foreach (OpenXmlElement property in tableProperties.ChildElements) {
                switch (property) {
                    case TableStyle tableStyle:
                        ReadSupportedTableStyleBorders(tableStyle, tableStyleDefinitions);
                        ReadSupportedTableStyleShading(tableStyle, tableStyleDefinitions);
                        break;
                    case TableWidth tableWidth:
                        ReadSupportedTablePreferredWidth(tableWidth);
                        break;
                    case TableJustification tableJustification:
                        ReadSupportedTableAlignment(tableJustification);
                        break;
                    case TableIndentation tableIndentation:
                        ReadSupportedTableIndentation(tableIndentation);
                        break;
                    case TableLayout tableLayout:
                        ReadSupportedTableAutofit(tableLayout);
                        break;
                    case TableCellMarginDefault tableCellMarginDefault:
                        ReadSupportedTableDefaultCellMargins(tableCellMarginDefault);
                        break;
                    case TableCellSpacing tableCellSpacing:
                        ReadSupportedTableDefaultCellSpacing(tableCellSpacing);
                        break;
                    case TableBorders tableBorders:
                        ReadSupportedTableBorders(tableBorders);
                        break;
                    case Shading shading:
                        ReadSupportedTableCellShading(shading, "table shading");
                        break;
                    case TableLook:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table property: {property.LocalName}.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableGrid(TableGrid tableGrid) {
            foreach (OpenXmlElement child in tableGrid.ChildElements) {
                if (child is GridColumn) {
                    continue;
                }

                throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table grid element: {child.LocalName}.");
            }
        }

        private static LegacyDocTableAlignment? ReadSupportedTableAlignment(TableProperties? tableProperties) {
            TableJustification? tableJustification = tableProperties?.GetFirstChild<TableJustification>();
            return tableJustification == null ? null : ReadSupportedTableAlignment(tableJustification);
        }

        private static LegacyDocTableAlignment? ReadSupportedTableAlignment(TableJustification tableJustification) {
            TableRowAlignmentValues? value = tableJustification.Val?.Value;
            if (value == null) {
                return null;
            }

            if (value == TableRowAlignmentValues.Left) {
                return LegacyDocTableAlignment.Left;
            }

            if (value == TableRowAlignmentValues.Center) {
                return LegacyDocTableAlignment.Center;
            }

            if (value == TableRowAlignmentValues.Right) {
                return LegacyDocTableAlignment.Right;
            }

            throw new NotSupportedException($"Native DOC saving does not support table alignment value '{value}'.");
        }

        private static int? ReadSupportedTableIndentation(TableProperties? tableProperties) {
            TableIndentation? tableIndentation = tableProperties?.GetFirstChild<TableIndentation>();
            return tableIndentation == null ? null : ReadSupportedTableIndentation(tableIndentation);
        }

        private static int? ReadSupportedTableIndentation(TableIndentation tableIndentation) {
            if (tableIndentation.Type?.Value != TableWidthUnitValues.Dxa) {
                throw new NotSupportedException("Native DOC saving supports table indentation only as DXA twip values.");
            }

            int? width = tableIndentation.Width?.Value;
            if (width == null) {
                return null;
            }

            if (width.Value < 0 || width.Value > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table indentation only as nonnegative Word 97-2003 signed twip values.");
            }

            return width.Value == 0 ? null : width.Value;
        }

        private static LegacyDocTablePreferredWidth? ReadSupportedTablePreferredWidth(TableProperties? tableProperties) {
            TableWidth? tableWidth = tableProperties?.GetFirstChild<TableWidth>();
            return tableWidth == null ? null : ReadSupportedTablePreferredWidth(tableWidth);
        }

        private static LegacyDocTablePreferredWidth? ReadSupportedTablePreferredWidth(TableWidth tableWidth) {
            TableWidthUnitValues? type = tableWidth.Type?.Value;
            string? widthText = tableWidth.Width?.Value;
            if (type == null || type == TableWidthUnitValues.Auto) {
                if (string.IsNullOrWhiteSpace(widthText) || widthText == "0") {
                    return null;
                }

                throw new NotSupportedException("Native DOC saving supports automatic table widths only with width 0.");
            }

            if (type != TableWidthUnitValues.Dxa && type != TableWidthUnitValues.Pct) {
                throw new NotSupportedException($"Native DOC saving does not support table width type '{type}'.");
            }

            if (!int.TryParse(widthText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int width)
                || width <= 0
                || width > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table preferred width only as a positive Word 97-2003 signed width value.");
            }

            LegacyDocTablePreferredWidthUnit unit = type == TableWidthUnitValues.Dxa
                ? LegacyDocTablePreferredWidthUnit.Dxa
                : LegacyDocTablePreferredWidthUnit.Percent;
            return new LegacyDocTablePreferredWidth(unit, width);
        }

        private static bool? ReadSupportedTableAutofit(TableProperties? tableProperties) {
            TableLayout? tableLayout = tableProperties?.GetFirstChild<TableLayout>();
            return tableLayout == null ? null : ReadSupportedTableAutofit(tableLayout);
        }

        private static bool? ReadSupportedTableAutofit(TableLayout tableLayout) {
            TableLayoutValues? value = tableLayout.Type?.Value;
            if (value == null) {
                return null;
            }

            if (value == TableLayoutValues.Autofit) {
                return true;
            }

            if (value == TableLayoutValues.Fixed) {
                return false;
            }

            throw new NotSupportedException($"Native DOC saving does not support table layout value '{value}'.");
        }

        private static LegacyDocTableCellMargins? ReadSupportedTableDefaultCellMargins(TableProperties? tableProperties) {
            TableCellMarginDefault? margins = tableProperties?.GetFirstChild<TableCellMarginDefault>();
            return margins == null ? null : ReadSupportedTableDefaultCellMargins(margins);
        }

        private static LegacyDocTableCellMargins? ReadSupportedTableDefaultCellMargins(TableCellMarginDefault margins) {
            LegacyDocTableCellMargins result = new LegacyDocTableCellMargins(
                ReadSupportedTableCellMarginWidth(margins.TopMargin, "default top"),
                ReadSupportedTableCellMarginWidth(margins.TableCellRightMargin, "default right"),
                ReadSupportedTableCellMarginWidth(margins.BottomMargin, "default bottom"),
                ReadSupportedTableCellMarginWidth(margins.TableCellLeftMargin, "default left"));
            return result.HasAny ? result : null;
        }

        private static int? ReadSupportedTableDefaultCellSpacing(TableProperties? tableProperties) {
            TableCellSpacing? spacing = tableProperties?.GetFirstChild<TableCellSpacing>();
            return spacing == null ? null : ReadSupportedTableDefaultCellSpacing(spacing);
        }

        private static int? ReadSupportedTableDefaultCellSpacing(TableCellSpacing spacing) {
            string? widthText = spacing.Width?.Value;
            if (string.IsNullOrWhiteSpace(widthText)) {
                return null;
            }

            if (!int.TryParse(widthText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int width)
                || width < 0
                || width > 31680) {
                throw new NotSupportedException("Native DOC saving supports table cell spacing only as nonnegative DXA twip values within the Word 97-2003 limit.");
            }

            if (width == 0) {
                return null;
            }

            if (spacing.Type?.Value != TableWidthUnitValues.Dxa) {
                throw new NotSupportedException("Native DOC saving supports table cell spacing only as DXA twip values.");
            }

            return width;
        }

        private static IReadOnlyList<int> ReadSupportedTableGridWidths(TableGrid? tableGrid) {
            if (tableGrid == null) {
                return Array.Empty<int>();
            }

            GridColumn[] columns = tableGrid.Elements<GridColumn>().ToArray();
            var widths = new int[columns.Length];
            for (int index = 0; index < columns.Length; index++) {
                string? widthText = columns[index].Width?.Value;
                if (string.IsNullOrWhiteSpace(widthText)) {
                    widths[index] = 0;
                    continue;
                }

                if (!int.TryParse(widthText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int width)
                    || width <= 0
                    || width > short.MaxValue) {
                    throw new NotSupportedException("Native DOC saving supports table grid column widths only as positive DXA twip values within the Word 97-2003 signed twip range.");
                }

                widths[index] = width;
            }

            return widths;
        }

        private static LegacyDocWritableTableRowFormatting ReadSupportedTableRowFormatting(TableRow row, out TableCell[] cells) {
            var tableCells = new List<TableCell>();
            LegacyDocWritableTableRowFormatting rowFormatting = LegacyDocWritableTableRowFormatting.Empty;
            bool hasTableRowProperties = false;
            foreach (OpenXmlElement child in row.ChildElements) {
                switch (child) {
                    case TableRowProperties tableRowProperties:
                        if (hasTableRowProperties) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one table row property collection per row.");
                        }

                        rowFormatting = ReadSupportedTableRowProperties(tableRowProperties);
                        hasTableRowProperties = true;
                        break;
                    case TableCell tableCell:
                        tableCells.Add(tableCell);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table row element: {child.LocalName}.");
                }
            }

            cells = tableCells.ToArray();
            return rowFormatting;
        }

        private static LegacyDocWritableTableRowFormatting ReadSupportedTableRowProperties(TableRowProperties tableRowProperties) {
            int? rowHeightTwips = null;
            bool rowHeightIsExact = false;
            bool? rowCantSplit = null;
            bool? rowIsHeader = null;
            bool hasRowHeight = false;
            bool hasCantSplit = false;
            bool hasTableHeader = false;
            foreach (OpenXmlElement property in tableRowProperties.ChildElements) {
                switch (property) {
                    case TableRowHeight rowHeight:
                        if (hasRowHeight) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one table row height per row.");
                        }

                        ReadSupportedTableRowHeight(rowHeight, out rowHeightTwips, out rowHeightIsExact);
                        hasRowHeight = true;
                        break;
                    case CantSplit cantSplit:
                        if (hasCantSplit) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one row no-split flag per row.");
                        }

                        rowCantSplit = ReadTableRowOnOff(cantSplit);
                        hasCantSplit = true;
                        break;
                    case TableHeader tableHeader:
                        if (hasTableHeader) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one row header flag per row.");
                        }

                        rowIsHeader = ReadTableRowOnOff(tableHeader);
                        hasTableHeader = true;
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table row property: {property.LocalName}.");
                }
            }

            return new LegacyDocWritableTableRowFormatting(rowHeightTwips, rowHeightIsExact, rowCantSplit, rowIsHeader);
        }

        private static void ReadSupportedTableRowHeight(TableRowHeight rowHeight, out int? rowHeightTwips, out bool rowHeightIsExact) {
            rowHeightTwips = null;
            rowHeightIsExact = false;

            HeightRuleValues? heightRule = rowHeight.HeightType?.Value;
            if (heightRule == HeightRuleValues.Auto) {
                return;
            }

            uint? rawValue = rowHeight.Val?.Value;
            if (rawValue == null || rawValue.Value == 0) {
                return;
            }

            if (rawValue.Value > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table row heights only as positive twip values within the Word 97-2003 signed twip range.");
            }

            if (heightRule != null && heightRule != HeightRuleValues.Exact && heightRule != HeightRuleValues.AtLeast) {
                throw new NotSupportedException($"Native DOC saving does not support table row height rule '{heightRule}'.");
            }

            rowHeightTwips = checked((int)rawValue.Value);
            rowHeightIsExact = heightRule == HeightRuleValues.Exact;
        }

        private static bool? ReadTableRowOnOff(CantSplit cantSplit) {
            if (cantSplit.Val == null) {
                return true;
            }

            return ReadTableRowOnOffValue(cantSplit.Val.InnerText) ? true : null;
        }

        private static bool? ReadTableRowOnOff(TableHeader tableHeader) {
            if (tableHeader.Val == null) {
                return true;
            }

            return ReadTableRowOnOffValue(tableHeader.Val.InnerText) ? true : null;
        }

        private static bool ReadTableRowOnOffValue(string? value) =>
            !string.Equals(value, "0", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(value, "off", StringComparison.OrdinalIgnoreCase);

        private static IReadOnlyList<LegacyDocWritableTableCell> ExpandSupportedTableCells(
            IReadOnlyList<TableCell> cells,
            IReadOnlyList<int> gridColumnWidthsTwips,
            LegacyDocTableBorders tableBorders,
            LegacyDocTableCellShading tableShading,
            LegacyDocTableConditionalStyleSet conditionalStyles,
            LegacyDocTableLook tableLook,
            int rowIndex,
            int rowCount) {
            var writableCells = new List<LegacyDocWritableTableCell>();
            int logicalColumnIndex = 0;
            foreach (TableCell cell in cells) {
                TableCellProperties? cellProperties = cell.TableCellProperties;
                int gridSpan = ReadSupportedGridSpan(cellProperties);
                LegacyDocTableCellHorizontalMerge horizontalMerge = ReadSupportedTableCellHorizontalMerge(cell);
                LegacyDocTableCellVerticalMerge verticalMerge = ReadSupportedTableCellVerticalMerge(cell);
                LegacyDocTableCellVerticalAlignment verticalAlignment = ReadSupportedTableCellVerticalAlignment(cellProperties);
                LegacyDocTableCellTextDirection textDirection = ReadSupportedTableCellTextDirection(cellProperties);
                bool fitText = ReadSupportedTableCellFitText(cellProperties);
                bool noWrap = ReadSupportedTableCellNoWrap(cellProperties);
                bool hideMark = ReadSupportedTableCellHideMark(cellProperties);
                LegacyDocTableCellMargins margins = ReadSupportedTableCellMargins(cellProperties);
                LegacyDocTableCellShading shading = ReadSupportedTableCellShading(cellProperties);
                LegacyDocTableCellBorders borders = ReadSupportedTableCellBorders(cellProperties);
                if (gridSpan > 1 && horizontalMerge == LegacyDocTableCellHorizontalMerge.Continue) {
                    throw new NotSupportedException("Native DOC saving supports simple horizontal table cell merges only. A continued horizontal merge cannot also define gridSpan.");
                }

                for (int spanIndex = 0; spanIndex < gridSpan; spanIndex++) {
                    int width = ReadSupportedTableCellWidthForSpan(cellProperties, gridColumnWidthsTwips, logicalColumnIndex, spanIndex, gridSpan);
                    LegacyDocTableCellHorizontalMerge merge = gridSpan == 1
                        ? horizontalMerge
                        : spanIndex == 0
                            ? LegacyDocTableCellHorizontalMerge.Restart
                            : LegacyDocTableCellHorizontalMerge.Continue;
                    writableCells.Add(new LegacyDocWritableTableCell(spanIndex == 0 ? cell : null, width, merge, verticalMerge, verticalAlignment, textDirection, fitText, noWrap, hideMark, margins, shading, borders));
                }

                logicalColumnIndex += gridSpan;
            }

            IReadOnlyList<LegacyDocWritableTableCell> conditionallyStyledCells = ApplySupportedTableConditionalStyles(writableCells, conditionalStyles, tableLook, rowIndex, rowCount);
            return ApplySupportedTableBorders(ApplySupportedTableShading(conditionallyStyledCells, tableShading), tableBorders, rowIndex, rowCount);
        }

        private static IReadOnlyList<LegacyDocWritableTableCell> ApplySupportedTableShading(
            IReadOnlyList<LegacyDocWritableTableCell> writableCells,
            LegacyDocTableCellShading tableShading) {
            if (!tableShading.HasAny || writableCells.Count == 0) {
                return writableCells;
            }

            var shadedCells = new LegacyDocWritableTableCell[writableCells.Count];
            for (int columnIndex = 0; columnIndex < writableCells.Count; columnIndex++) {
                LegacyDocWritableTableCell cell = writableCells[columnIndex];
                shadedCells[columnIndex] = cell.Shading.HasAny
                    ? cell
                    : cell.WithShading(tableShading);
            }

            return shadedCells;
        }

        private static IReadOnlyList<LegacyDocWritableTableCell> ApplySupportedTableBorders(
            IReadOnlyList<LegacyDocWritableTableCell> writableCells,
            LegacyDocTableBorders tableBorders,
            int rowIndex,
            int rowCount) {
            if (!tableBorders.HasAny || writableCells.Count == 0) {
                return writableCells;
            }

            var borderedCells = new LegacyDocWritableTableCell[writableCells.Count];
            for (int columnIndex = 0; columnIndex < writableCells.Count; columnIndex++) {
                LegacyDocWritableTableCell cell = writableCells[columnIndex];
                borderedCells[columnIndex] = cell.WithBorders(MergeSupportedTableBorders(
                    cell.Borders,
                    tableBorders,
                    rowIndex,
                    rowCount,
                    columnIndex,
                    writableCells.Count));
            }

            return borderedCells;
        }

        private static LegacyDocTableCellBorders MergeSupportedTableBorders(
            LegacyDocTableCellBorders cellBorders,
            LegacyDocTableBorders tableBorders,
            int rowIndex,
            int rowCount,
            int columnIndex,
            int columnCount) {
            LegacyDocTableCellBorder top = cellBorders.Top.HasAny
                ? cellBorders.Top
                : rowIndex == 0 ? tableBorders.Top : tableBorders.InsideHorizontal;
            LegacyDocTableCellBorder left = cellBorders.Left.HasAny
                ? cellBorders.Left
                : columnIndex == 0 ? tableBorders.Left : tableBorders.InsideVertical;
            LegacyDocTableCellBorder bottom = cellBorders.Bottom.HasAny
                ? cellBorders.Bottom
                : rowIndex + 1 >= rowCount ? tableBorders.Bottom : tableBorders.InsideHorizontal;
            LegacyDocTableCellBorder right = cellBorders.Right.HasAny
                ? cellBorders.Right
                : columnIndex + 1 >= columnCount ? tableBorders.Right : tableBorders.InsideVertical;

            return new LegacyDocTableCellBorders(top, left, bottom, right);
        }

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

            if (value == MergedCellValues.Continue) {
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
                return default;
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
            Shading? shading = tableProperties?.GetFirstChild<Shading>();
            if (shading != null) {
                return ReadSupportedTableCellShading(shading, "table shading");
            }

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
            if (value == null || value == BorderValues.None || value == BorderValues.Nil) {
                return default;
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

        private static LegacyDocWritableParagraphFormatting AppendTableCell(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            List<LegacyDocWritableParagraph> paragraphFormats,
            LegacyDocWritableBookmarksBuilder bookmarks,
            TableCell? cell,
            MainDocumentPart mainPart,
            IReadOnlyDictionary<string, ushort> styleIndexes,
            LegacyDocWritableParagraphFormatting tableStyleParagraphFormatting,
            LegacyDocWritableFormatting tableStyleRunFormatting,
            LegacyDocWritableFootnotes footnotes,
            LegacyDocWritableEndnotes endnotes,
            out int finalParagraphStart) {
            finalParagraphStart = text.Length;
            if (cell == null) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            if (cell.Elements<Table>().Any()) {
                throw new NotSupportedException("Native DOC saving supports simple tables only. Nested tables are not supported yet.");
            }

            var paragraphs = new List<Paragraph>();
            foreach (OpenXmlElement child in cell.ChildElements) {
                switch (child) {
                    case TableCellProperties cellProperties:
                        ThrowIfUnsupportedTableCellProperties(cellProperties);
                        break;
                    case Paragraph cellParagraph:
                        paragraphs.Add(cellParagraph);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table cell element: {child.LocalName}.");
                }
            }

            if (paragraphs.Count == 0) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            for (int index = 0; index < paragraphs.Count - 1; index++) {
                int paragraphStart = text.Length;
                LegacyDocWritableParagraphFormatting paragraphFormatting = AppendTableCellParagraph(text, runs, bookmarks, paragraphs[index], mainPart, styleIndexes, tableStyleParagraphFormatting, tableStyleRunFormatting, footnotes, endnotes)
                    .WithTableMarkers(isTableTerminatingParagraph: false);
                text.Append('\r');
                paragraphFormats.Add(new LegacyDocWritableParagraph(paragraphStart, text.Length - paragraphStart, paragraphFormatting));
            }

            finalParagraphStart = text.Length;
            return AppendTableCellParagraph(text, runs, bookmarks, paragraphs[paragraphs.Count - 1], mainPart, styleIndexes, tableStyleParagraphFormatting, tableStyleRunFormatting, footnotes, endnotes);
        }

        private static void ThrowIfUnsupportedTableCellProperties(TableCellProperties cellProperties) {
            foreach (OpenXmlElement property in cellProperties.ChildElements) {
                switch (property) {
                    case TableCellWidth cellWidth:
                        ReadSupportedTableCellWidth(cellProperties, gridColumnWidthTwips: 0);
                        break;
                    case GridSpan:
                        ReadSupportedGridSpan(cellProperties);
                        break;
                    case HorizontalMerge:
                        break;
                    case VerticalMerge:
                        break;
                    case TableCellVerticalAlignment:
                        ReadSupportedTableCellVerticalAlignment(cellProperties);
                        break;
                    case TextDirection:
                        ReadSupportedTableCellTextDirection(cellProperties);
                        break;
                    case TableCellFitText:
                        ReadSupportedTableCellFitText(cellProperties);
                        break;
                    case NoWrap:
                        ReadSupportedTableCellNoWrap(cellProperties);
                        break;
                    case HideMark:
                        ReadSupportedTableCellHideMark(cellProperties);
                        break;
                    case TableCellMargin:
                        ReadSupportedTableCellMargins(cellProperties);
                        break;
                    case Shading:
                        ReadSupportedTableCellShading(cellProperties);
                        break;
                    case TableCellBorders:
                        ReadSupportedTableCellBorders(cellProperties);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table cell property: {property.LocalName}.");
                }
            }
        }

        private static LegacyDocWritableParagraphFormatting AppendTableCellParagraph(StringBuilder text, List<LegacyDocWritableRun> runs, LegacyDocWritableBookmarksBuilder bookmarks, Paragraph paragraph, MainDocumentPart mainPart, IReadOnlyDictionary<string, ushort> styleIndexes, LegacyDocWritableParagraphFormatting tableStyleParagraphFormatting, LegacyDocWritableFormatting tableStyleRunFormatting, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes) {
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedParagraphFormatting(paragraph.ParagraphProperties, styleIndexes)
                .WithInheritedParagraphFormatting(tableStyleParagraphFormatting);

            OpenXmlElement[] children = paragraph.ChildElements.ToArray();
            for (int index = 0; index < children.Length; index++) {
                OpenXmlElement child = children[index];
                switch (child) {
                    case ParagraphProperties:
                        break;
                    case Run run:
                        if (IsComplexFieldBeginRun(run)) {
                            AppendSupportedComplexPageNumberField(children, ref index, text, runs, tableStyleRunFormatting);
                        } else {
                            AppendSupportedRunText(text, runs, run, footnotes, endnotes, tableStyleRunFormatting);
                        }

                        break;
                    case Hyperlink hyperlink:
                        AppendSupportedHyperlinkText(text, runs, hyperlink, mainPart, footnotes, endnotes, tableStyleRunFormatting);
                        break;
                    case SimpleField simpleField:
                        AppendSupportedPageNumberFieldFromSimpleField(text, runs, simpleField, tableStyleRunFormatting);
                        break;
                    case BookmarkStart bookmarkStart:
                        bookmarks.AddStart(bookmarkStart, text.Length);
                        break;
                    case BookmarkEnd bookmarkEnd:
                        bookmarks.AddEnd(bookmarkEnd, text.Length);
                        break;
                    default:
                        if (IsIgnorableParagraphMarkup(child)) {
                            break;
                        }

                        throw new NotSupportedException($"Native DOC saving supports simple table cell paragraphs only with text runs, PAGE and NUMPAGES simple fields, bookmarks, and simple hyperlinks. Unsupported paragraph element: {child.LocalName}.");
                }
            }

            return paragraphFormatting;
        }

        private readonly struct LegacyDocWritableTableRowFormatting {
            internal static readonly LegacyDocWritableTableRowFormatting Empty = new LegacyDocWritableTableRowFormatting(null, false, null, null);

            internal LegacyDocWritableTableRowFormatting(int? rowHeightTwips, bool rowHeightIsExact, bool? rowCantSplit, bool? rowIsHeader) {
                RowHeightTwips = rowHeightTwips;
                RowHeightIsExact = rowHeightIsExact;
                RowCantSplit = rowCantSplit;
                RowIsHeader = rowIsHeader;
            }

            internal int? RowHeightTwips { get; }

            internal bool RowHeightIsExact { get; }

            internal bool? RowCantSplit { get; }

            internal bool? RowIsHeader { get; }
        }

        private readonly struct LegacyDocTableBorders {
            internal LegacyDocTableBorders(
                LegacyDocTableCellBorder top,
                LegacyDocTableCellBorder left,
                LegacyDocTableCellBorder bottom,
                LegacyDocTableCellBorder right,
                LegacyDocTableCellBorder insideHorizontal,
                LegacyDocTableCellBorder insideVertical) {
                Top = top;
                Left = left;
                Bottom = bottom;
                Right = right;
                InsideHorizontal = insideHorizontal;
                InsideVertical = insideVertical;
            }

            internal LegacyDocTableCellBorder Top { get; }

            internal LegacyDocTableCellBorder Left { get; }

            internal LegacyDocTableCellBorder Bottom { get; }

            internal LegacyDocTableCellBorder Right { get; }

            internal LegacyDocTableCellBorder InsideHorizontal { get; }

            internal LegacyDocTableCellBorder InsideVertical { get; }

            internal bool HasAny => Top.HasAny
                || Left.HasAny
                || Bottom.HasAny
                || Right.HasAny
                || InsideHorizontal.HasAny
                || InsideVertical.HasAny;
        }

        private readonly struct LegacyDocWritableTableCell {
            internal LegacyDocWritableTableCell(TableCell? sourceCell, int widthTwips, LegacyDocTableCellHorizontalMerge horizontalMerge, LegacyDocTableCellVerticalMerge verticalMerge, LegacyDocTableCellVerticalAlignment verticalAlignment, LegacyDocTableCellTextDirection textDirection, bool fitText, bool noWrap, bool hideMark, LegacyDocTableCellMargins margins, LegacyDocTableCellShading shading, LegacyDocTableCellBorders borders)
                : this(sourceCell, widthTwips, horizontalMerge, verticalMerge, verticalAlignment, textDirection, fitText, noWrap, hideMark, margins, shading, borders, LegacyDocWritableParagraphFormatting.Plain, LegacyDocWritableFormatting.Plain) {
            }

            private LegacyDocWritableTableCell(TableCell? sourceCell, int widthTwips, LegacyDocTableCellHorizontalMerge horizontalMerge, LegacyDocTableCellVerticalMerge verticalMerge, LegacyDocTableCellVerticalAlignment verticalAlignment, LegacyDocTableCellTextDirection textDirection, bool fitText, bool noWrap, bool hideMark, LegacyDocTableCellMargins margins, LegacyDocTableCellShading shading, LegacyDocTableCellBorders borders, LegacyDocWritableParagraphFormatting paragraphFormatting, LegacyDocWritableFormatting runFormatting) {
                SourceCell = sourceCell;
                WidthTwips = widthTwips;
                HorizontalMerge = horizontalMerge;
                VerticalMerge = verticalMerge;
                VerticalAlignment = verticalAlignment;
                TextDirection = textDirection;
                FitText = fitText;
                NoWrap = noWrap;
                HideMark = hideMark;
                Margins = margins;
                Shading = shading;
                Borders = borders;
                ParagraphFormatting = paragraphFormatting;
                RunFormatting = runFormatting;
            }

            internal TableCell? SourceCell { get; }

            internal int WidthTwips { get; }

            internal LegacyDocTableCellHorizontalMerge HorizontalMerge { get; }

            internal LegacyDocTableCellVerticalMerge VerticalMerge { get; }

            internal LegacyDocTableCellVerticalAlignment VerticalAlignment { get; }

            internal LegacyDocTableCellTextDirection TextDirection { get; }

            internal bool FitText { get; }

            internal bool NoWrap { get; }

            internal bool HideMark { get; }

            internal LegacyDocTableCellMargins Margins { get; }

            internal LegacyDocTableCellShading Shading { get; }

            internal LegacyDocTableCellBorders Borders { get; }

            internal LegacyDocWritableParagraphFormatting ParagraphFormatting { get; }

            internal LegacyDocWritableFormatting RunFormatting { get; }

            internal LegacyDocWritableTableCell WithBorders(LegacyDocTableCellBorders borders) =>
                new LegacyDocWritableTableCell(SourceCell, WidthTwips, HorizontalMerge, VerticalMerge, VerticalAlignment, TextDirection, FitText, NoWrap, HideMark, Margins, Shading, borders, ParagraphFormatting, RunFormatting);

            internal LegacyDocWritableTableCell WithShading(LegacyDocTableCellShading shading) =>
                new LegacyDocWritableTableCell(SourceCell, WidthTwips, HorizontalMerge, VerticalMerge, VerticalAlignment, TextDirection, FitText, NoWrap, HideMark, Margins, shading, Borders, ParagraphFormatting, RunFormatting);

            internal LegacyDocWritableTableCell WithVerticalAlignment(LegacyDocTableCellVerticalAlignment verticalAlignment) =>
                new LegacyDocWritableTableCell(SourceCell, WidthTwips, HorizontalMerge, VerticalMerge, verticalAlignment, TextDirection, FitText, NoWrap, HideMark, Margins, Shading, Borders, ParagraphFormatting, RunFormatting);

            internal LegacyDocWritableTableCell WithTextDirection(LegacyDocTableCellTextDirection textDirection) =>
                new LegacyDocWritableTableCell(SourceCell, WidthTwips, HorizontalMerge, VerticalMerge, VerticalAlignment, textDirection, FitText, NoWrap, HideMark, Margins, Shading, Borders, ParagraphFormatting, RunFormatting);

            internal LegacyDocWritableTableCell WithFitText(bool fitText) =>
                new LegacyDocWritableTableCell(SourceCell, WidthTwips, HorizontalMerge, VerticalMerge, VerticalAlignment, TextDirection, fitText, NoWrap, HideMark, Margins, Shading, Borders, ParagraphFormatting, RunFormatting);

            internal LegacyDocWritableTableCell WithNoWrap(bool noWrap) =>
                new LegacyDocWritableTableCell(SourceCell, WidthTwips, HorizontalMerge, VerticalMerge, VerticalAlignment, TextDirection, FitText, noWrap, HideMark, Margins, Shading, Borders, ParagraphFormatting, RunFormatting);

            internal LegacyDocWritableTableCell WithHideMark(bool hideMark) =>
                new LegacyDocWritableTableCell(SourceCell, WidthTwips, HorizontalMerge, VerticalMerge, VerticalAlignment, TextDirection, FitText, NoWrap, hideMark, Margins, Shading, Borders, ParagraphFormatting, RunFormatting);

            internal LegacyDocWritableTableCell WithMargins(LegacyDocTableCellMargins margins) =>
                new LegacyDocWritableTableCell(SourceCell, WidthTwips, HorizontalMerge, VerticalMerge, VerticalAlignment, TextDirection, FitText, NoWrap, HideMark, margins, Shading, Borders, ParagraphFormatting, RunFormatting);

            internal LegacyDocWritableTableCell WithParagraphFormatting(LegacyDocWritableParagraphFormatting paragraphFormatting) =>
                new LegacyDocWritableTableCell(SourceCell, WidthTwips, HorizontalMerge, VerticalMerge, VerticalAlignment, TextDirection, FitText, NoWrap, HideMark, Margins, Shading, Borders, paragraphFormatting, RunFormatting);

            internal LegacyDocWritableTableCell WithRunFormatting(LegacyDocWritableFormatting runFormatting) =>
                new LegacyDocWritableTableCell(SourceCell, WidthTwips, HorizontalMerge, VerticalMerge, VerticalAlignment, TextDirection, FitText, NoWrap, HideMark, Margins, Shading, Borders, ParagraphFormatting, runFormatting);
        }
    }
}
