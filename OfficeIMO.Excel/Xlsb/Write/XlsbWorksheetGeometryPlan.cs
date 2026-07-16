using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Builds validated BIFF12 worksheet geometry records for a newly generated sheet.</summary>
    internal sealed class XlsbWorksheetGeometryPlan {
        private const int BrtColInfo = 60;
        private const int BrtBeginWsViews = 133;
        private const int BrtEndWsViews = 134;
        private const int BrtBeginWsView = 137;
        private const int BrtEndWsView = 138;
        private const int BrtPane = 151;
        private const int BrtSelection = 152;
        private const int BrtMergeCell = 176;
        private const int BrtBeginMergeCells = 177;
        private const int BrtEndMergeCells = 178;
        private const int BrtBeginColInfos = 390;
        private const int BrtEndColInfos = 391;
        private const int BrtWsFmtInfo = 485;

        private XlsbWorksheetGeometryPlan(
            byte[] dimensionPayload,
            List<XlsbGeneratedRecord> prefixRecords,
            Dictionary<int, byte[]> rowProperties,
            List<XlsbGeneratedRecord> suffixRecords) {
            DimensionPayload = dimensionPayload;
            PrefixRecords = prefixRecords.AsReadOnly();
            RowProperties = rowProperties;
            SuffixRecords = suffixRecords.AsReadOnly();
        }

        internal byte[] DimensionPayload { get; }

        internal IReadOnlyList<XlsbGeneratedRecord> PrefixRecords { get; }

        internal IReadOnlyDictionary<int, byte[]> RowProperties { get; }

        internal IReadOnlyList<XlsbGeneratedRecord> SuffixRecords { get; }

        internal static XlsbWorksheetGeometryPlan Create(
            ExcelSheet sheet,
            IReadOnlyList<XlsbWriteCell> cells,
            int cellFormatCount) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (cells == null) throw new ArgumentNullException(nameof(cells));
            if (cellFormatCount <= 0) throw new ArgumentOutOfRangeException(nameof(cellFormatCount));

            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            ValidateWorksheetChildren(worksheet, sheet.Name);

            List<XlsbCellRange> merges = ReadMergedRanges(worksheet, sheet.Name);
            byte[] dimension = CreateDimensionPayload(worksheet.GetFirstChild<SheetDimension>(), cells, merges, sheet.Name);
            var prefix = new List<XlsbGeneratedRecord>();
            AppendPaneRecord(prefix, worksheet.GetFirstChild<SheetViews>(), sheet.Name);
            AppendFormatRecord(prefix, worksheet.GetFirstChild<SheetFormatProperties>(), sheet.Name);
            AppendColumnRecords(prefix, worksheet.GetFirstChild<Columns>(), cellFormatCount, sheet.Name);
            Dictionary<int, byte[]> rows = CreateRowProperties(worksheet.GetFirstChild<SheetData>(), cellFormatCount, sheet.Name);
            List<XlsbGeneratedRecord> suffix = CreateMergeRecords(merges);
            return new XlsbWorksheetGeometryPlan(dimension, prefix, rows, suffix);
        }

        private static void ValidateWorksheetChildren(Worksheet worksheet, string sheetName) {
            var counts = new Dictionary<Type, int>();
            foreach (OpenXmlElement child in worksheet.ChildElements) {
                if (child is not SheetProperties
                    && child is not SheetDimension
                    && child is not SheetViews
                    && child is not SheetFormatProperties
                    && child is not Columns
                    && child is not SheetData
                    && child is not SheetProtection
                    && child is not AutoFilter
                    && child is not MergeCells
                    && child is not Hyperlinks
                    && child is not PrintOptions
                    && child is not PageMargins
                    && child is not PageSetup
                    && child is not HeaderFooter) {
                    throw new NotSupportedException($"Native XLSB generation does not yet support worksheet metadata '{child.LocalName}' on worksheet '{sheetName}'.");
                }

                Type type = child.GetType();
                counts.TryGetValue(type, out int count);
                if (count != 0) {
                    throw new NotSupportedException($"Native XLSB generation requires at most one '{child.LocalName}' element on worksheet '{sheetName}'.");
                }
                counts[type] = 1;
            }

            EnsureOnlyAttributes(worksheet, sheetName);
            SheetDimension? dimension = worksheet.GetFirstChild<SheetDimension>();
            if (dimension != null) EnsureOnlyAttributes(dimension, sheetName, "ref");
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData != null) EnsureOnlyAttributes(sheetData, sheetName);
        }

        private static byte[] CreateDimensionPayload(
            SheetDimension? dimension,
            IReadOnlyList<XlsbWriteCell> cells,
            IReadOnlyList<XlsbCellRange> merges,
            string sheetName) {
            int firstRow = int.MaxValue;
            int firstColumn = int.MaxValue;
            int lastRow = 0;
            int lastColumn = 0;

            if (dimension?.Reference?.Value is string reference && reference.Length != 0) {
                if (!TryParseRange(reference, out int dimensionFirstRow, out int dimensionFirstColumn, out int dimensionLastRow, out int dimensionLastColumn)) {
                    throw new NotSupportedException($"Native XLSB generation cannot encode invalid worksheet dimension '{reference}' on worksheet '{sheetName}'.");
                }
                IncludeRange(ref firstRow, ref firstColumn, ref lastRow, ref lastColumn,
                    dimensionFirstRow, dimensionFirstColumn, dimensionLastRow, dimensionLastColumn);
            }

            foreach (XlsbWriteCell cell in cells) {
                IncludeRange(ref firstRow, ref firstColumn, ref lastRow, ref lastColumn,
                    cell.Row, cell.Column, cell.Row, cell.Column);
            }
            foreach (XlsbCellRange merge in merges) {
                IncludeRange(ref firstRow, ref firstColumn, ref lastRow, ref lastColumn,
                    merge.FirstRow, merge.FirstColumn, merge.LastRow, merge.LastColumn);
            }

            if (lastRow == 0) firstRow = firstColumn = lastRow = lastColumn = 1;
            using var payload = new MemoryStream(16);
            WriteUInt32(payload, checked((uint)(firstRow - 1)));
            WriteUInt32(payload, checked((uint)(lastRow - 1)));
            WriteUInt32(payload, checked((uint)(firstColumn - 1)));
            WriteUInt32(payload, checked((uint)(lastColumn - 1)));
            return payload.ToArray();
        }

        private static void AppendFormatRecord(
            List<XlsbGeneratedRecord> records,
            SheetFormatProperties? format,
            string sheetName) {
            if (format == null) return;
            EnsureOnlyAttributes(format, sheetName,
                "baseColWidth", "defaultColWidth", "defaultRowHeight", "customHeight", "zeroHeight", "outlineLevelRow", "outlineLevelCol");
            if (format.HasChildren) ThrowUnsupportedContent(format, sheetName);

            double columnWidth = format.DefaultColumnWidth?.Value ?? format.BaseColumnWidth?.Value ?? 8D;
            double rowHeight = format.DefaultRowHeight?.Value ?? 15D;
            ValidateFiniteRange(columnWidth, 0D, 255D, "default column width", sheetName);
            ValidateFiniteRange(rowHeight, 0D, ushort.MaxValue / 20D, "default row height", sheetName);
            byte rowOutline = format.OutlineLevelRow?.Value ?? 0;
            byte columnOutline = format.OutlineLevelColumn?.Value ?? 0;
            if (rowOutline > 7 || columnOutline > 7) {
                throw new NotSupportedException($"Native XLSB generation supports worksheet outline levels from 0 through 7 on worksheet '{sheetName}'.");
            }

            using var payload = new MemoryStream(12);
            WriteUInt32(payload, checked((uint)Math.Round(columnWidth * 256D, MidpointRounding.AwayFromZero)));
            WriteUInt16(payload, checked((ushort)Math.Round(columnWidth, MidpointRounding.AwayFromZero)));
            WriteUInt16(payload, checked((ushort)Math.Round(rowHeight * 20D, MidpointRounding.AwayFromZero)));
            uint flags = (format.CustomHeight?.Value == true ? 0x01U : 0U)
                | (format.ZeroHeight?.Value == true ? 0x02U : 0U)
                | ((uint)rowOutline << 16)
                | ((uint)columnOutline << 24);
            WriteUInt32(payload, flags);
            records.Add(new XlsbGeneratedRecord(BrtWsFmtInfo, payload.ToArray()));
        }

        private static void AppendPaneRecord(List<XlsbGeneratedRecord> records, SheetViews? views, string sheetName) {
            if (views == null) return;
            EnsureOnlyAttributes(views, sheetName);
            SheetView[] sheetViews = views.Elements<SheetView>().ToArray();
            if (sheetViews.Length != 1 || views.ChildElements.Count != 1) {
                throw new NotSupportedException($"Native XLSB generation supports one sheet view on worksheet '{sheetName}'.");
            }

            SheetView view = sheetViews[0];
            EnsureOnlyAttributes(view, sheetName, "workbookViewId");
            if ((view.WorkbookViewId?.Value ?? 0U) != 0U) {
                throw new NotSupportedException($"Native XLSB generation supports workbookViewId 0 on worksheet '{sheetName}'.");
            }
            Pane[] panes = view.Elements<Pane>().ToArray();
            if (panes.Length > 1 || view.ChildElements.Any(child => child is not Pane && child is not Selection)) {
                throw new NotSupportedException($"Native XLSB generation supports one pane and standard selections on worksheet '{sheetName}'.");
            }
            foreach (Selection selection in view.Elements<Selection>()) {
                EnsureOnlyAttributes(selection, sheetName, "pane", "activeCell", "activeCellId", "sqref");
                if (selection.HasChildren) ThrowUnsupportedContent(selection, sheetName);
            }

            records.Add(new XlsbGeneratedRecord(BrtBeginWsViews, Array.Empty<byte>()));
            records.Add(new XlsbGeneratedRecord(BrtBeginWsView, CreateDefaultSheetViewPayload()));

            Pane? pane = panes.FirstOrDefault();
            if (pane != null) AppendPane(records, pane, sheetName);

            Selection[] selections = view.Elements<Selection>()
                .OrderBy(selection => selection.Pane == null ? 0 : 1)
                .ToArray();
            if (selections.Length == 0) {
                records.Add(new XlsbGeneratedRecord(BrtSelection, CreateSelectionPayload(null, sheetName)));
            } else {
                foreach (Selection selection in selections) {
                    records.Add(new XlsbGeneratedRecord(BrtSelection, CreateSelectionPayload(selection, sheetName)));
                }
            }

            records.Add(new XlsbGeneratedRecord(BrtEndWsView, Array.Empty<byte>()));
            records.Add(new XlsbGeneratedRecord(BrtEndWsViews, Array.Empty<byte>()));
        }

        private static void AppendPane(List<XlsbGeneratedRecord> records, Pane pane, string sheetName) {

            EnsureOnlyAttributes(pane, sheetName, "xSplit", "ySplit", "topLeftCell", "activePane", "state");
            if (pane.HasChildren) ThrowUnsupportedContent(pane, sheetName);
            double horizontalSplit = pane.HorizontalSplit?.Value ?? 0D;
            double verticalSplit = pane.VerticalSplit?.Value ?? 0D;
            ValidateFiniteRange(horizontalSplit, 0D, int.MaxValue, "horizontal pane split", sheetName);
            ValidateFiniteRange(verticalSplit, 0D, int.MaxValue, "vertical pane split", sheetName);

            int topRow;
            int leftColumn;
            if (!string.IsNullOrWhiteSpace(pane.TopLeftCell?.Value)) {
                if (!A1.TryParseCellReferenceFast(pane.TopLeftCell!.Value!, out topRow, out leftColumn)) {
                    throw new NotSupportedException($"Native XLSB generation cannot encode pane top-left cell '{pane.TopLeftCell.Value}' on worksheet '{sheetName}'.");
                }
            } else {
                topRow = checked((int)Math.Truncate(verticalSplit)) + 1;
                leftColumn = checked((int)Math.Truncate(horizontalSplit)) + 1;
            }

            PaneValues? paneValue = pane.ActivePane?.Value;
            uint activePane = paneValue == PaneValues.TopRight
                ? 1U
                : paneValue == PaneValues.BottomLeft
                    ? 2U
                    : paneValue == PaneValues.TopLeft ? 3U : 0U;
            PaneStateValues? stateValue = pane.State?.Value;
            byte flags = stateValue == PaneStateValues.Frozen
                ? (byte)0x01
                : stateValue == PaneStateValues.FrozenSplit ? (byte)0x03 : (byte)0x00;

            using var payload = new MemoryStream(29);
            WriteDouble(payload, horizontalSplit);
            WriteDouble(payload, verticalSplit);
            WriteUInt32(payload, checked((uint)(topRow - 1)));
            WriteUInt32(payload, checked((uint)(leftColumn - 1)));
            WriteUInt32(payload, activePane);
            payload.WriteByte(flags);
            records.Add(new XlsbGeneratedRecord(BrtPane, payload.ToArray()));
        }

        private static byte[] CreateDefaultSheetViewPayload() {
            return new byte[] {
                0xDC, 0x03, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x40, 0x00,
                0x00, 0x00, 0x64, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00,
                0x00, 0x00
            };
        }

        private static byte[] CreateSelectionPayload(Selection? selection, string sheetName) {
            uint pane = selection?.Pane?.Value == PaneValues.TopRight
                ? 1U
                : selection?.Pane?.Value == PaneValues.BottomLeft
                    ? 2U
                    : selection?.Pane?.Value == PaneValues.BottomRight ? 0U : 3U;
            string activeCellReference = selection?.ActiveCell?.Value ?? "A1";
            if (!A1.TryParseCellReferenceFast(activeCellReference, out int activeRow, out int activeColumn)) {
                throw new NotSupportedException($"Native XLSB generation cannot encode active selection cell '{activeCellReference}' on worksheet '{sheetName}'.");
            }

            string selectedReference = selection?.SequenceOfReferences?.InnerText?.Trim() ?? activeCellReference;
            if (selectedReference.IndexOf(' ') >= 0
                || !TryParseRange(selectedReference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                throw new NotSupportedException($"Native XLSB generation currently supports one contiguous selection range on worksheet '{sheetName}'.");
            }

            using var payload = new MemoryStream(36);
            WriteUInt32(payload, pane);
            WriteUInt32(payload, checked((uint)(activeRow - 1)));
            WriteUInt32(payload, checked((uint)(activeColumn - 1)));
            WriteUInt32(payload, selection?.ActiveCellId?.Value ?? 0U);
            WriteUInt32(payload, 1U);
            WriteUInt32(payload, checked((uint)(firstRow - 1)));
            WriteUInt32(payload, checked((uint)(lastRow - 1)));
            WriteUInt32(payload, checked((uint)(firstColumn - 1)));
            WriteUInt32(payload, checked((uint)(lastColumn - 1)));
            return payload.ToArray();
        }

        private static void AppendColumnRecords(
            List<XlsbGeneratedRecord> records,
            Columns? columns,
            int cellFormatCount,
            string sheetName) {
            if (columns == null) return;
            EnsureOnlyAttributes(columns, sheetName);
            Column[] definitions = columns.Elements<Column>().ToArray();
            if (definitions.Length != columns.ChildElements.Count) ThrowUnsupportedContent(columns, sheetName);
            if (definitions.Length == 0) return;

            records.Add(new XlsbGeneratedRecord(BrtBeginColInfos, Array.Empty<byte>()));
            uint previousLast = 0;
            bool hasPrevious = false;
            foreach (Column column in definitions) {
                EnsureOnlyAttributes(column, sheetName,
                    "min", "max", "width", "style", "hidden", "bestFit", "customWidth", "phonetic", "outlineLevel", "collapsed");
                if (column.HasChildren) ThrowUnsupportedContent(column, sheetName);
                uint first = column.Min?.Value ?? 0U;
                uint last = column.Max?.Value ?? 0U;
                if (first < 1U || first > last || last > A1.MaxColumns || (hasPrevious && first <= previousLast)) {
                    throw new NotSupportedException($"Native XLSB generation requires ordered, non-overlapping column ranges on worksheet '{sheetName}'.");
                }
                previousLast = last;
                hasPrevious = true;
                double width = column.Width?.Value ?? 0D;
                ValidateFiniteRange(width, 0D, 255D, "column width", sheetName);
                uint style = column.Style?.Value ?? 0U;
                ValidateStyleIndex(style, cellFormatCount, $"column {first}", sheetName);
                byte outline = column.OutlineLevel?.Value ?? 0;
                if (outline > 7) {
                    throw new NotSupportedException($"Native XLSB generation supports column outline levels from 0 through 7 on worksheet '{sheetName}'.");
                }

                ushort flags = (ushort)((column.Hidden?.Value == true ? 0x0001 : 0)
                    | (column.CustomWidth?.Value == true ? 0x0002 : 0)
                    | (column.BestFit?.Value == true ? 0x0004 : 0)
                    | (column.Phonetic?.Value == true ? 0x0008 : 0)
                    | (outline << 8)
                    | (column.Collapsed?.Value == true ? 0x1000 : 0));
                using var payload = new MemoryStream(18);
                WriteUInt32(payload, first - 1U);
                WriteUInt32(payload, last - 1U);
                WriteUInt32(payload, checked((uint)Math.Round(width * 256D, MidpointRounding.AwayFromZero)));
                WriteUInt32(payload, style);
                WriteUInt16(payload, flags);
                records.Add(new XlsbGeneratedRecord(BrtColInfo, payload.ToArray()));
            }
            records.Add(new XlsbGeneratedRecord(BrtEndColInfos, Array.Empty<byte>()));
        }

        private static Dictionary<int, byte[]> CreateRowProperties(
            SheetData? sheetData,
            int cellFormatCount,
            string sheetName) {
            var result = new Dictionary<int, byte[]>();
            if (sheetData == null) return result;
            foreach (Row row in sheetData.Elements<Row>()) {
                EnsureOnlyAttributes(row, sheetName,
                    "r", "spans", "s", "customFormat", "ht", "hidden", "customHeight", "outlineLevel", "collapsed");
                if (row.ChildElements.Any(child => child is not Cell)) ThrowUnsupportedContent(row, sheetName);
                uint index = row.RowIndex?.Value ?? 0U;
                if (index < 1U || index > A1.MaxRows || result.ContainsKey(checked((int)index))) {
                    throw new NotSupportedException($"Native XLSB generation requires unique row indexes from 1 through {A1.MaxRows} on worksheet '{sheetName}'.");
                }

                bool customFormat = row.CustomFormat?.Value == true;
                uint style = row.StyleIndex?.Value ?? 0U;
                if (!customFormat && style != 0U) {
                    throw new NotSupportedException($"Native XLSB generation requires customFormat for styled row {index} on worksheet '{sheetName}'.");
                }
                ValidateStyleIndex(style, cellFormatCount, $"row {index}", sheetName);
                bool customHeight = row.CustomHeight?.Value == true;
                double height = customHeight ? row.Height?.Value ?? 15D : 15D;
                ValidateFiniteRange(height, 0D, ushort.MaxValue / 20D, "row height", sheetName);
                byte outline = row.OutlineLevel?.Value ?? 0;
                if (outline > 7) {
                    throw new NotSupportedException($"Native XLSB generation supports row outline levels from 0 through 7 on worksheet '{sheetName}'.");
                }

                using var properties = new MemoryStream(9);
                WriteUInt32(properties, style);
                WriteUInt16(properties, checked((ushort)Math.Round(height * 20D, MidpointRounding.AwayFromZero)));
                properties.WriteByte(0);
                properties.WriteByte((byte)(outline
                    | (row.Collapsed?.Value == true ? 0x08 : 0)
                    | (row.Hidden?.Value == true ? 0x10 : 0)
                    | (customHeight ? 0x20 : 0)
                    | (customFormat ? 0x40 : 0)));
                properties.WriteByte(0);
                result.Add(checked((int)index - 1), properties.ToArray());
            }
            return result;
        }

        private static List<XlsbCellRange> ReadMergedRanges(Worksheet worksheet, string sheetName) {
            MergeCells? mergeCells = worksheet.GetFirstChild<MergeCells>();
            var result = new List<XlsbCellRange>();
            if (mergeCells == null) return result;
            EnsureOnlyAttributes(mergeCells, sheetName, "count");
            MergeCell[] definitions = mergeCells.Elements<MergeCell>().ToArray();
            if (definitions.Length != mergeCells.ChildElements.Count) ThrowUnsupportedContent(mergeCells, sheetName);
            if (mergeCells.Count?.Value is uint declared && declared != definitions.Length) {
                throw new NotSupportedException($"Native XLSB generation found a mismatched merged-cell count on worksheet '{sheetName}'.");
            }
            foreach (MergeCell merge in definitions) {
                EnsureOnlyAttributes(merge, sheetName, "ref");
                if (merge.HasChildren || !TryParseRange(merge.Reference?.Value, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                    throw new NotSupportedException($"Native XLSB generation cannot encode merged range '{merge.Reference?.Value}' on worksheet '{sheetName}'.");
                }
                result.Add(new XlsbCellRange(firstRow, firstColumn, lastRow, lastColumn));
            }
            return result;
        }

        private static List<XlsbGeneratedRecord> CreateMergeRecords(IReadOnlyList<XlsbCellRange> merges) {
            var records = new List<XlsbGeneratedRecord>();
            if (merges.Count == 0) return records;
            using var begin = new MemoryStream(4);
            WriteUInt32(begin, checked((uint)merges.Count));
            records.Add(new XlsbGeneratedRecord(BrtBeginMergeCells, begin.ToArray()));
            foreach (XlsbCellRange merge in merges) {
                using var payload = new MemoryStream(16);
                WriteUInt32(payload, checked((uint)(merge.FirstRow - 1)));
                WriteUInt32(payload, checked((uint)(merge.LastRow - 1)));
                WriteUInt32(payload, checked((uint)(merge.FirstColumn - 1)));
                WriteUInt32(payload, checked((uint)(merge.LastColumn - 1)));
                records.Add(new XlsbGeneratedRecord(BrtMergeCell, payload.ToArray()));
            }
            records.Add(new XlsbGeneratedRecord(BrtEndMergeCells, Array.Empty<byte>()));
            return records;
        }

        private static bool TryParseRange(
            string? reference,
            out int firstRow,
            out int firstColumn,
            out int lastRow,
            out int lastColumn) {
            firstRow = firstColumn = lastRow = lastColumn = 0;
            if (string.IsNullOrWhiteSpace(reference)) return false;
            if (A1.TryParseRange(reference!, out firstRow, out firstColumn, out lastRow, out lastColumn)) return true;
            if (!A1.TryParseCellReferenceFast(reference!, out firstRow, out firstColumn)) return false;
            lastRow = firstRow;
            lastColumn = firstColumn;
            return true;
        }

        private static void IncludeRange(
            ref int firstRow,
            ref int firstColumn,
            ref int lastRow,
            ref int lastColumn,
            int rangeFirstRow,
            int rangeFirstColumn,
            int rangeLastRow,
            int rangeLastColumn) {
            firstRow = Math.Min(firstRow, rangeFirstRow);
            firstColumn = Math.Min(firstColumn, rangeFirstColumn);
            lastRow = Math.Max(lastRow, rangeLastRow);
            lastColumn = Math.Max(lastColumn, rangeLastColumn);
        }

        private static void ValidateFiniteRange(double value, double minimum, double maximum, string detail, string sheetName) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < minimum || value > maximum) {
                throw new NotSupportedException($"Native XLSB generation cannot encode {detail} {value} on worksheet '{sheetName}'.");
            }
        }

        private static void ValidateStyleIndex(uint styleIndex, int cellFormatCount, string detail, string sheetName) {
            if (styleIndex >= cellFormatCount) {
                throw new NotSupportedException($"Native XLSB generation found {detail} on worksheet '{sheetName}' with missing style index {styleIndex}.");
            }
        }

        private static void EnsureOnlyAttributes(OpenXmlElement element, string sheetName, params string[] allowedNames) {
            var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
            OpenXmlAttribute? unsupported = element.GetAttributes()
                .Cast<OpenXmlAttribute?>()
                .FirstOrDefault(attribute => attribute.HasValue
                    && !string.Equals(attribute.Value.NamespaceUri, "http://www.w3.org/2000/xmlns/", StringComparison.Ordinal)
                    && !allowed.Contains(attribute.Value.LocalName));
            if (unsupported.HasValue) {
                throw new NotSupportedException($"Native XLSB generation does not yet support attribute '{unsupported.Value.LocalName}' on worksheet element '{element.LocalName}' in worksheet '{sheetName}'.");
            }
        }

        private static void ThrowUnsupportedContent(OpenXmlElement element, string sheetName) =>
            throw new NotSupportedException($"Native XLSB generation does not yet support child content in worksheet element '{element.LocalName}' on worksheet '{sheetName}'.");

        private static void WriteDouble(Stream stream, double value) {
            byte[] bytes = BitConverter.GetBytes(value);
            stream.Write(bytes, 0, bytes.Length);
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
            stream.WriteByte((byte)(value >> 16));
            stream.WriteByte((byte)(value >> 24));
        }
    }

}
