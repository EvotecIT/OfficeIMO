using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
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
                    case BookmarkStart:
                    case BookmarkEnd:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table element: {child.LocalName}.");
                }
            }
        }

        private static void AppendTableRowBoundaryBookmarks(Table table, TableRow row, LegacyDocWritableBookmarksBuilder bookmarks, int characterPosition) {
            OpenXmlElement? child = row.NextSibling();
            while (child != null && child is not TableRow) {
                if (child is not BookmarkStart bookmarkStart) {
                    throw new NotSupportedException($"Native DOC saving supports row-level table bookmarks only as zero-length start/end marker pairs between table rows. Unsupported row-boundary table element: {child.LocalName}.");
                }

                OpenXmlElement? endMarker = child.NextSibling();
                if (endMarker is not BookmarkEnd bookmarkEnd || bookmarkEnd.Id?.Value != bookmarkStart.Id?.Value) {
                    throw new NotSupportedException("Native DOC saving supports row-level table bookmarks only as zero-length start/end marker pairs between table rows.");
                }

                bookmarks.AddStart(bookmarkStart, characterPosition);
                bookmarks.AddEnd(bookmarkEnd, characterPosition);
                child = endMarker.NextSibling();
            }
        }

        private static void AppendLeadingTableBoundaryBookmarks(Table table, LegacyDocWritableBookmarksBuilder bookmarks, int characterPosition) {
            foreach (OpenXmlElement child in table.ChildElements) {
                if (child is TableRow) {
                    return;
                }

                AppendSupportedTableBoundaryBookmark(bookmarks, child, characterPosition);
            }
        }

        private static void AppendTrailingTableBoundaryBookmarks(Table table, LegacyDocWritableBookmarksBuilder bookmarks, int characterPosition) {
            var trailingChildren = new List<OpenXmlElement>();
            foreach (OpenXmlElement child in table.ChildElements) {
                if (child is TableRow) {
                    trailingChildren.Clear();
                    continue;
                }

                trailingChildren.Add(child);
            }

            foreach (OpenXmlElement child in trailingChildren) {
                AppendSupportedTableBoundaryBookmark(bookmarks, child, characterPosition);
            }
        }

        private static void AppendSupportedTableBoundaryBookmark(LegacyDocWritableBookmarksBuilder bookmarks, OpenXmlElement child, int characterPosition) {
            switch (child) {
                case BookmarkStart bookmarkStart:
                    bookmarks.AddStart(bookmarkStart, characterPosition);
                    break;
                case BookmarkEnd bookmarkEnd:
                    bookmarks.AddEnd(bookmarkEnd, characterPosition);
                    break;
            }
        }

        private static void ThrowIfUnsupportedTableProperties(TableProperties tableProperties, IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            ThrowIfUnsupportedTableGapShading(tableProperties, tableStyleDefinitions);
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

            if (width.Value < short.MinValue || width.Value > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table indentation only as Word 97-2003 signed twip values.");
            }

            // An explicit zero overrides an inherited indent; only omission permits style fallback.
            return width.Value;
        }

        private static LegacyDocTablePreferredWidth? ReadSupportedTablePreferredWidth(TableProperties? tableProperties) {
            TableWidth? tableWidth = tableProperties?.GetFirstChild<TableWidth>();
            return tableWidth == null ? null : ReadSupportedTablePreferredWidth(tableWidth);
        }

        private static LegacyDocTablePreferredWidth? ReadSupportedTablePreferredWidth(TableWidth tableWidth) {
            TableWidthUnitValues? type = tableWidth.Type?.Value;
            string? widthText = tableWidth.Width?.Value;
            if (type == TableWidthUnitValues.Auto) {
                return null;
            }

            if (type == null) {
                if (string.IsNullOrWhiteSpace(widthText) || widthText == "0") {
                    return null;
                }

                throw new NotSupportedException("Native DOC saving supports table widths without an explicit type only when width is 0.");
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
                return 0;
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
                shadedCells[columnIndex] = cell.Shading.IsSpecified
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

            internal bool HasFormatting => RowHeightTwips != null || RowCantSplit != null || RowIsHeader != null;

            internal LegacyDocWritableTableRowFormatting WithInheritedRowFormatting(LegacyDocWritableTableRowFormatting inherited) {
                return new LegacyDocWritableTableRowFormatting(
                    RowHeightTwips ?? inherited.RowHeightTwips,
                    RowHeightTwips != null ? RowHeightIsExact : inherited.RowHeightIsExact,
                    RowCantSplit ?? inherited.RowCantSplit,
                    RowIsHeader ?? inherited.RowIsHeader);
            }
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
