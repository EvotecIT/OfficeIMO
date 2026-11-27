using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void ThrowIfUnsupportedTableShape(LegacyDocElement table, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            foreach (LegacyDocElement child in table.Elements()) {
                switch (child.WName) {
                    case "tblPr":
                        ThrowIfUnsupportedTableProperties(child, tableStyleDefinitions);
                        break;
                    case "tblGrid":
                        ThrowIfUnsupportedTableGrid(child);
                        break;
                    case "tr":
                        break;
                    case "bookmarkStart":
                    case "bookmarkEnd":
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table element: {child.LocalName}.");
                }
            }
        }

        private static void AppendTableRowBoundaryBookmarks(LegacyDocElement table, LegacyDocElement row, LegacyDocWritableBookmarksBuilder bookmarks, int characterPosition) {
            LegacyDocElement child = row.NextElement;
            while (child.Exists && !child.Is("tr")) {
                if (!child.Is("bookmarkStart")) {
                    throw new NotSupportedException($"Native DOC saving supports row-level table bookmarks only as zero-length start/end marker pairs between table rows. Unsupported row-boundary table element: {child.LocalName}.");
                }

                LegacyDocElement bookmarkStart = child;
                LegacyDocElement endMarker = child.NextElement;
                if (!endMarker.Is("bookmarkEnd") || endMarker.Attr("id") != bookmarkStart.Attr("id")) {
                    throw new NotSupportedException("Native DOC saving supports row-level table bookmarks only as zero-length start/end marker pairs between table rows.");
                }

                bookmarks.AddStart(bookmarkStart, characterPosition);
                bookmarks.AddEnd(endMarker, characterPosition);
                child = endMarker.NextElement;
            }
        }

        private static void AppendLeadingTableBoundaryBookmarks(LegacyDocElement table, LegacyDocWritableBookmarksBuilder bookmarks, int characterPosition) {
            foreach (LegacyDocElement child in table.Elements()) {
                if (child.Is("tr")) {
                    return;
                }

                AppendSupportedTableBoundaryBookmark(bookmarks, child, characterPosition);
            }
        }

        private static void AppendTrailingTableBoundaryBookmarks(LegacyDocElement table, LegacyDocWritableBookmarksBuilder bookmarks, int characterPosition) {
            var trailingChildren = new List<LegacyDocElement>();
            foreach (LegacyDocElement child in table.Elements()) {
                if (child.Is("tr")) {
                    trailingChildren.Clear();
                    continue;
                }

                trailingChildren.Add(child);
            }

            foreach (LegacyDocElement child in trailingChildren) {
                AppendSupportedTableBoundaryBookmark(bookmarks, child, characterPosition);
            }
        }

        private static void AppendSupportedTableBoundaryBookmark(LegacyDocWritableBookmarksBuilder bookmarks, LegacyDocElement child, int characterPosition) {
            switch (child.WName) {
                case "bookmarkStart":
                    bookmarks.AddStart(child, characterPosition);
                    break;
                case "bookmarkEnd":
                    bookmarks.AddEnd(child, characterPosition);
                    break;
            }
        }

        private static void ThrowIfUnsupportedTableProperties(LegacyDocElement tableProperties, LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            ThrowIfUnsupportedTableGapShading(tableProperties, tableStyleDefinitions);
            foreach (LegacyDocElement property in tableProperties.Elements()) {
                switch (property.WName) {
                    case "tblStyle":
                        ReadSupportedTableStyleBorders(property.Val, tableStyleDefinitions);
                        ReadSupportedTableStyleShading(property.Val, tableStyleDefinitions);
                        break;
                    case "tblW":
                        ReadSupportedTablePreferredWidth(property);
                        break;
                    case "jc":
                        ReadSupportedTableAlignment(property);
                        break;
                    case "tblInd":
                        ReadSupportedTableIndentation(property);
                        break;
                    case "tblLayout":
                        ReadSupportedTableAutofit(property);
                        break;
                    case "tblCellMar":
                        ReadSupportedTableDefaultCellMargins(property);
                        break;
                    case "tblCellSpacing":
                        ReadSupportedTableDefaultCellSpacing(property);
                        break;
                    case "tblBorders":
                        ReadSupportedTableBorders(property);
                        break;
                    case "shd":
                        ReadSupportedTableCellShading(property, "table shading");
                        break;
                    case "tblLook":
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table property: {property.LocalName}.");
                }
            }
        }

        private static void ThrowIfUnsupportedTableGrid(LegacyDocElement tableGrid) {
            foreach (LegacyDocElement child in tableGrid.Elements()) {
                if (child.Is("gridCol")) {
                    continue;
                }

                throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table grid element: {child.LocalName}.");
            }
        }

        /// <summary>
        /// The text the Open XML SDK interpolated for one of its enumeration values in a message: the value type's
        /// record form, or nothing when the value was absent.
        /// </summary>
        private static string FormatSdkEnumValue(string? token, string sdkTypeName) =>
            token == null ? string.Empty : sdkTypeName + " { }";

        /// <summary>
        /// The first attribute whose local name matches ignoring case, in the order the SDK listed a typed element's
        /// attributes: the schema attribute in the <c>w</c> namespace first, then the others in document order.
        /// </summary>
        private static string? FindTableAttributeIgnoringCase(LegacyDocElement element, string localName) {
            string? schemaValue = element.Attr(localName);
            if (schemaValue != null) {
                return schemaValue;
            }

            foreach (LegacyDocAttribute attribute in element.Attributes) {
                if (string.Equals(attribute.LocalName, localName, StringComparison.OrdinalIgnoreCase)) {
                    return attribute.Value;
                }
            }

            return null;
        }

        private static LegacyDocTableAlignment? ReadSupportedTableAlignmentFromProperties(LegacyDocElement tableProperties) {
            LegacyDocElement tableJustification = tableProperties.Child("jc");
            return tableJustification.IsNone ? null : ReadSupportedTableAlignment(tableJustification);
        }

        private static LegacyDocTableAlignment? ReadSupportedTableAlignment(LegacyDocElement tableJustification) {
            string? value = tableJustification.Val;
            if (value == null) {
                return null;
            }

            if (value == "left") {
                return LegacyDocTableAlignment.Left;
            }

            if (value == "center") {
                return LegacyDocTableAlignment.Center;
            }

            if (value == "right") {
                return LegacyDocTableAlignment.Right;
            }

            throw new NotSupportedException($"Native DOC saving does not support table alignment value '{FormatSdkEnumValue(value, "TableRowAlignmentValues")}'.");
        }

        private static int? ReadSupportedTableIndentationFromProperties(LegacyDocElement tableProperties) {
            LegacyDocElement tableIndentation = tableProperties.Child("tblInd");
            return tableIndentation.IsNone ? null : ReadSupportedTableIndentation(tableIndentation);
        }

        private static int? ReadSupportedTableIndentation(LegacyDocElement tableIndentation) {
            if (tableIndentation.Attr("type") != "dxa") {
                throw new NotSupportedException("Native DOC saving supports table indentation only as DXA twip values.");
            }

            int? width = LegacyDocValues.Int32(tableIndentation.Attr("w"));
            if (width == null) {
                return null;
            }

            if (width.Value < short.MinValue || width.Value > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table indentation only as Word 97-2003 signed twip values.");
            }

            // An explicit zero overrides an inherited indent; only omission permits style fallback.
            return width.Value;
        }

        private static LegacyDocTablePreferredWidth? ReadSupportedTablePreferredWidthFromProperties(LegacyDocElement tableProperties) {
            LegacyDocElement tableWidth = tableProperties.Child("tblW");
            return tableWidth.IsNone ? null : ReadSupportedTablePreferredWidth(tableWidth);
        }

        private static LegacyDocTablePreferredWidth? ReadSupportedTablePreferredWidth(LegacyDocElement tableWidth) {
            string? typeText = tableWidth.Attr("type");
            string? type = typeText == null ? null : LegacyDocWordHelpers.NativeTableWidthType(tableWidth);
            if (typeText != null && type == null)
                throw new NotSupportedException("Native DOC saving does not support table width type '" + typeText + "'.");
            string? widthText = tableWidth.Attr("w");
            if (type == "auto") {
                return null;
            }

            if (type == null) {
                if (string.IsNullOrWhiteSpace(widthText) || widthText == "0") {
                    return null;
                }

                throw new NotSupportedException("Native DOC saving supports table widths without an explicit type only when width is 0.");
            }

            if (type != "dxa" && type != "pct") {
                throw new NotSupportedException($"Native DOC saving does not support table width type '{FormatSdkEnumValue(type, "TableWidthUnitValues")}'.");
            }

            if (!int.TryParse(widthText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int width)
                || width <= 0
                || width > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table preferred width only as a positive Word 97-2003 signed width value.");
            }

            LegacyDocTablePreferredWidthUnit unit = type == "dxa"
                ? LegacyDocTablePreferredWidthUnit.Dxa
                : LegacyDocTablePreferredWidthUnit.Percent;
            return new LegacyDocTablePreferredWidth(unit, width);
        }

        private static bool? ReadSupportedTableAutofitFromProperties(LegacyDocElement tableProperties) {
            LegacyDocElement tableLayout = tableProperties.Child("tblLayout");
            return tableLayout.IsNone ? null : ReadSupportedTableAutofit(tableLayout);
        }

        private static bool? ReadSupportedTableAutofit(LegacyDocElement tableLayout) {
            string? value = tableLayout.Attr("type");
            if (value == null) {
                return null;
            }

            if (value == "autofit") {
                return true;
            }

            if (value == "fixed") {
                return false;
            }

            throw new NotSupportedException($"Native DOC saving does not support table layout value '{FormatSdkEnumValue(value, "TableLayoutValues")}'.");
        }

        private static LegacyDocTableCellMargins? ReadSupportedTableDefaultCellMarginsFromProperties(LegacyDocElement tableProperties) {
            LegacyDocElement margins = tableProperties.Child("tblCellMar");
            return margins.IsNone ? null : ReadSupportedTableDefaultCellMargins(margins);
        }

        private static LegacyDocTableCellMargins? ReadSupportedTableDefaultCellMargins(LegacyDocElement margins) {
            LegacyDocTableCellMargins result = new LegacyDocTableCellMargins(
                ReadSupportedTableCellMarginWidth(margins.Child("top"), "default top"),
                ReadSupportedTableCellMarginWidth(margins.Child("right"), "default right"),
                ReadSupportedTableCellMarginWidth(margins.Child("bottom"), "default bottom"),
                ReadSupportedTableCellMarginWidth(margins.Child("left"), "default left"));
            return result.HasAny ? result : null;
        }

        private static int? ReadSupportedTableDefaultCellSpacingFromProperties(LegacyDocElement tableProperties) {
            LegacyDocElement spacing = tableProperties.Child("tblCellSpacing");
            return spacing.IsNone ? null : ReadSupportedTableDefaultCellSpacing(spacing);
        }

        private static int? ReadSupportedTableDefaultCellSpacing(LegacyDocElement spacing) {
            string? widthText = spacing.Attr("w");
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

            if (spacing.Attr("type") != "dxa") {
                throw new NotSupportedException("Native DOC saving supports table cell spacing only as DXA twip values.");
            }

            return width;
        }

        private static IReadOnlyList<int> ReadSupportedTableGridWidths(LegacyDocElement tableGrid) {
            if (tableGrid.IsNone) {
                return Array.Empty<int>();
            }

            LegacyDocElement[] columns = tableGrid.Elements("gridCol").ToArray();
            var widths = new int[columns.Length];
            for (int index = 0; index < columns.Length; index++) {
                string? widthText = columns[index].Attr("w");
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

        private static LegacyDocWritableTableRowFormatting ReadSupportedTableRowFormatting(LegacyDocElement row, out LegacyDocElement[] cells) {
            var tableCells = new List<LegacyDocElement>();
            LegacyDocWritableTableRowFormatting rowFormatting = LegacyDocWritableTableRowFormatting.Empty;
            bool hasTableRowProperties = false;
            foreach (LegacyDocElement child in row.Elements()) {
                switch (child.WName) {
                    case "trPr":
                        if (hasTableRowProperties) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one table row property collection per row.");
                        }

                        rowFormatting = ReadSupportedTableRowProperties(child);
                        hasTableRowProperties = true;
                        break;
                    case "tc":
                        tableCells.Add(child);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table row element: {child.LocalName}.");
                }
            }

            cells = tableCells.ToArray();
            return rowFormatting;
        }

        private static LegacyDocWritableTableRowFormatting ReadSupportedTableRowProperties(LegacyDocElement tableRowProperties) {
            int? rowHeightTwips = null;
            bool rowHeightIsExact = false;
            bool? rowCantSplit = null;
            bool? rowIsHeader = null;
            bool hasRowHeight = false;
            bool hasCantSplit = false;
            bool hasTableHeader = false;
            foreach (LegacyDocElement property in tableRowProperties.Elements()) {
                switch (property.WName) {
                    case "trHeight":
                        if (hasRowHeight) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one table row height per row.");
                        }

                        ReadSupportedTableRowHeight(property, out rowHeightTwips, out rowHeightIsExact);
                        hasRowHeight = true;
                        break;
                    case "cantSplit":
                        if (hasCantSplit) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one row no-split flag per row.");
                        }

                        rowCantSplit = ReadTableRowOnOff(property);
                        hasCantSplit = true;
                        break;
                    case "tblHeader":
                        if (hasTableHeader) {
                            throw new NotSupportedException("Native DOC saving supports simple tables only with one row header flag per row.");
                        }

                        rowIsHeader = ReadTableRowOnOff(property);
                        hasTableHeader = true;
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving supports simple tables only. Unsupported table row property: {property.LocalName}.");
                }
            }

            return new LegacyDocWritableTableRowFormatting(rowHeightTwips, rowHeightIsExact, rowCantSplit, rowIsHeader);
        }

        private static void ReadSupportedTableRowHeight(LegacyDocElement rowHeight, out int? rowHeightTwips, out bool rowHeightIsExact) {
            rowHeightTwips = null;
            rowHeightIsExact = false;

            string? heightRule = rowHeight.Attr("hRule");
            if (heightRule == "auto") {
                return;
            }

            uint? rawValue = LegacyDocValues.UInt32(rowHeight.Val);
            if (rawValue == null || rawValue.Value == 0) {
                return;
            }

            if (rawValue.Value > short.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports table row heights only as positive twip values within the Word 97-2003 signed twip range.");
            }

            if (heightRule != null && heightRule != "exact" && heightRule != "atLeast") {
                throw new NotSupportedException($"Native DOC saving does not support table row height rule '{FormatSdkEnumValue(heightRule, "HeightRuleValues")}'.");
            }

            rowHeightTwips = checked((int)rawValue.Value);
            rowHeightIsExact = heightRule == "exact";
        }

        // Reads w:cantSplit and w:tblHeader alike: present without a value is on.
        private static bool? ReadTableRowOnOff(LegacyDocElement rowFlag) {
            string? value = rowFlag.Val;
            if (value == null) {
                return true;
            }

            return ReadTableRowOnOffValue(value) ? true : null;
        }

        private static bool ReadTableRowOnOffValue(string? value) =>
            !string.Equals(value, "0", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(value, "off", StringComparison.OrdinalIgnoreCase);

        private static IReadOnlyList<LegacyDocWritableTableCell> ExpandSupportedTableCells(
            IReadOnlyList<LegacyDocElement> cells,
            IReadOnlyList<int> gridColumnWidthsTwips,
            LegacyDocTableBorders tableBorders,
            LegacyDocTableCellShading tableShading,
            LegacyDocTableConditionalStyleSet conditionalStyles,
            LegacyDocTableLook tableLook,
            int rowIndex,
            int rowCount) {
            var writableCells = new List<LegacyDocWritableTableCell>();
            int logicalColumnIndex = 0;
            foreach (LegacyDocElement cell in cells) {
                LegacyDocElement cellProperties = cell.Child("tcPr");
                int gridSpan = ReadSupportedGridSpan(cellProperties);
                LegacyDocTableCellHorizontalMerge horizontalMerge = ReadSupportedTableCellHorizontalMerge(cell);
                LegacyDocTableCellVerticalMerge verticalMerge = ReadSupportedTableCellVerticalMerge(cell);
                LegacyDocTableCellVerticalAlignment verticalAlignment = ReadSupportedTableCellVerticalAlignmentFromProperties(cellProperties);
                LegacyDocTableCellTextDirection textDirection = ReadSupportedTableCellTextDirectionFromProperties(cellProperties);
                bool fitText = ReadSupportedTableCellFitTextFromProperties(cellProperties);
                bool noWrap = ReadSupportedTableCellNoWrapFromProperties(cellProperties);
                bool hideMark = ReadSupportedTableCellHideMarkFromProperties(cellProperties);
                LegacyDocTableCellMargins margins = ReadSupportedTableCellMarginsFromProperties(cellProperties);
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
                    writableCells.Add(new LegacyDocWritableTableCell(spanIndex == 0 ? cell : LegacyDocElement.None, width, merge, verticalMerge, verticalAlignment, textDirection, fitText, noWrap, hideMark, margins, shading, borders));
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

        private static void ThrowIfUnsupportedTableCellProperties(LegacyDocElement cellProperties) {
            foreach (LegacyDocElement property in cellProperties.Elements()) {
                switch (property.WName) {
                    case "tcW":
                        ReadSupportedTableCellWidth(cellProperties, gridColumnWidthTwips: 0);
                        break;
                    case "gridSpan":
                        ReadSupportedGridSpan(cellProperties);
                        break;
                    case "hMerge":
                        break;
                    case "vMerge":
                        break;
                    case "vAlign":
                        ReadSupportedTableCellVerticalAlignmentFromProperties(cellProperties);
                        break;
                    case "textDirection":
                        ReadSupportedTableCellTextDirectionFromProperties(cellProperties);
                        break;
                    case "tcFitText":
                        ReadSupportedTableCellFitTextFromProperties(cellProperties);
                        break;
                    case "noWrap":
                        ReadSupportedTableCellNoWrapFromProperties(cellProperties);
                        break;
                    case "hideMark":
                        ReadSupportedTableCellHideMarkFromProperties(cellProperties);
                        break;
                    case "tcMar":
                        ReadSupportedTableCellMarginsFromProperties(cellProperties);
                        break;
                    case "shd":
                        ReadSupportedTableCellShading(cellProperties);
                        break;
                    case "tcBorders":
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
            internal LegacyDocWritableTableCell(LegacyDocElement sourceCell, int widthTwips, LegacyDocTableCellHorizontalMerge horizontalMerge, LegacyDocTableCellVerticalMerge verticalMerge, LegacyDocTableCellVerticalAlignment verticalAlignment, LegacyDocTableCellTextDirection textDirection, bool fitText, bool noWrap, bool hideMark, LegacyDocTableCellMargins margins, LegacyDocTableCellShading shading, LegacyDocTableCellBorders borders)
                : this(sourceCell, widthTwips, horizontalMerge, verticalMerge, verticalAlignment, textDirection, fitText, noWrap, hideMark, margins, shading, borders, LegacyDocWritableParagraphFormatting.Plain, LegacyDocWritableFormatting.Plain) {
            }

            private LegacyDocWritableTableCell(LegacyDocElement sourceCell, int widthTwips, LegacyDocTableCellHorizontalMerge horizontalMerge, LegacyDocTableCellVerticalMerge verticalMerge, LegacyDocTableCellVerticalAlignment verticalAlignment, LegacyDocTableCellTextDirection textDirection, bool fitText, bool noWrap, bool hideMark, LegacyDocTableCellMargins margins, LegacyDocTableCellShading shading, LegacyDocTableCellBorders borders, LegacyDocWritableParagraphFormatting paragraphFormatting, LegacyDocWritableFormatting runFormatting) {
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

            internal LegacyDocElement SourceCell { get; }

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
