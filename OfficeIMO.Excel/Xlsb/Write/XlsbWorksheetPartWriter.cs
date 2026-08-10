using OfficeIMO.Excel.Xlsb.Biff12;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Rewrites a worksheet cell table while preserving all records outside it and unknown in-table metadata.</summary>
    internal static class XlsbWorksheetPartWriter {
        private const int BrtRowHdr = 0;
        private const int BrtCellBlank = 1;
        private const int BrtCellRk = 2;
        private const int BrtCellError = 3;
        private const int BrtCellBool = 4;
        private const int BrtCellReal = 5;
        private const int BrtCellSt = 6;
        private const int BrtCellIsst = 7;
        private const int BrtFmlaString = 8;
        private const int BrtFmlaNum = 9;
        private const int BrtFmlaBool = 10;
        private const int BrtFmlaError = 11;
        private const int BrtCellRString = 62;
        private const int BrtBeginSheetData = 145;
        private const int BrtEndSheetData = 146;
        private const int BrtWsDim = 148;
        private const int BrtEndSheet = 130;
        private const int BrtMargins = 476;
        private const int BrtPrintOptions = 477;
        private const int BrtPageSetup = 478;
        private const int BrtBeginHeaderFooter = 479;
        private const int BrtEndHeaderFooter = 480;
        private const int BrtHLink = 494;
        private const int BrtBeginRowBreaks = 392;
        private const int BrtBeginColumnBreaks = 394;
        private const int BrtDrawing = 550;
        private const int BrtLegacyDrawing = 551;
        private const int BrtLegacyDrawingHeaderFooter = 552;
        private const int BrtBeginWebPublishItems = 554;
        private const int BrtBackgroundImage = 562;
        private const int BrtBeginDataValidations = 573;
        private const int BrtBeginSmartTags = 594;
        private const int BrtBeginCellWatches = 605;
        private const int BrtBigName = 625;
        private const int BrtBeginOleObjects = 638;
        private const int BrtBeginActiveXControls = 643;
        private const int BrtBeginCellIgnoreErrors = 648;
        private const int BrtBeginTableParts = 660;

        private static readonly byte[] DefaultRowProperties = {
            0x00, 0x00, 0x00, 0x00,
            0x2C, 0x01,
            0x00, 0x00, 0x00
        };

        internal static byte[] Create(
            ExcelSheet sheet,
            IReadOnlyList<XlsbWriteCell> cells,
            int cellFormatCount,
            IReadOnlyList<XlsbGeneratedRecord> hyperlinkRecords) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (cells == null) throw new ArgumentNullException(nameof(cells));
            if (hyperlinkRecords == null) throw new ArgumentNullException(nameof(hyperlinkRecords));

            XlsbWorksheetGeometryPlan geometry = XlsbWorksheetGeometryPlan.Create(sheet, cells, cellFormatCount);

            IReadOnlyDictionary<int, IReadOnlyList<XlsbWriteCell>> cellsByRow = cells
                .GroupBy(cell => cell.Row - 1)
                .ToDictionary(group => group.Key, group => (IReadOnlyList<XlsbWriteCell>)group.OrderBy(cell => cell.Column).ToArray());
            int[] rowIndexes = cellsByRow.Keys.Concat(geometry.RowProperties.Keys).Distinct().OrderBy(row => row).ToArray();

            using var output = new MemoryStream(Math.Max(256, cells.Count * 24));
            XlsbRecordWriter.Write(output, 129); // BrtBeginSheet
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            XlsbWorksheetPropertiesWriter.Write(output, worksheet.GetFirstChild<SheetProperties>(), sheet.Name);
            XlsbRecordWriter.Write(output, BrtWsDim, geometry.DimensionPayload);
            foreach (XlsbGeneratedRecord record in geometry.PrefixRecords) {
                XlsbRecordWriter.Write(output, record.Type, record.Payload);
            }
            XlsbRecordWriter.Write(output, BrtBeginSheetData);
            foreach (int rowIndex in rowIndexes) {
                cellsByRow.TryGetValue(rowIndex, out IReadOnlyList<XlsbWriteCell>? rowCells);
                geometry.RowProperties.TryGetValue(rowIndex, out byte[]? rowProperties);
                XlsbRecordWriter.Write(output, BrtRowHdr, CreateRowHeaderPayload(
                    rowIndex,
                    sourcePayload: null,
                    rowCells ?? Array.Empty<XlsbWriteCell>(),
                    rowProperties));
                if (rowCells == null) continue;
                foreach (XlsbWriteCell cell in rowCells) {
                    WriteCell(output, cell);
                }
            }
            XlsbRecordWriter.Write(output, BrtEndSheetData);
            XlsbWorksheetProtectionWriter.Write(output, worksheet.GetFirstChild<SheetProtection>());
            XlsbWorksheetAutoFilterWriter.Write(output, worksheet.GetFirstChild<AutoFilter>(), sheet.Name);
            foreach (XlsbGeneratedRecord record in geometry.SuffixRecords) {
                XlsbRecordWriter.Write(output, record.Type, record.Payload);
            }
            foreach (XlsbGeneratedRecord record in hyperlinkRecords) {
                XlsbRecordWriter.Write(output, record.Type, record.Payload);
            }
            XlsbWorksheetPrintSettingsWriter.Write(
                output,
                worksheet.GetFirstChild<PrintOptions>(),
                worksheet.GetFirstChild<PageMargins>(),
                worksheet.GetFirstChild<PageSetup>(),
                worksheet.GetFirstChild<HeaderFooter>(),
                sheet.Name);
            XlsbRecordWriter.Write(output, 130); // BrtEndSheet
            return output.ToArray();
        }

        internal static bool TryCreateDirectTabular(
            ExcelDirectTabularSource source,
            CancellationToken cancellationToken,
            out ArraySegment<byte> worksheetPart) {
            if (source == null) throw new ArgumentNullException(nameof(source));

            IExcelSheetTabularRowSource rows = source.Rows;
            int rowOffset = source.IncludeHeaders ? 1 : 0;
            int totalRows = checked(rows.RowCount + rowOffset);
            if (totalRows > 1_048_576 || rows.ColumnCount > 16_384) {
                throw new NotSupportedException("Native XLSB saving supports 1,048,576 rows and 16,384 columns per worksheet.");
            }

            using var output = new MemoryStream(EstimateDirectWorksheetCapacity(totalRows, rows.ColumnCount));
            using var writer = new XlsbDirectRecordWriter(output);
            writer.WriteRecord(129); // BrtBeginSheet
            writer.WriteHeader(BrtWsDim, 16);
            WriteDirectDimension(writer, totalRows, rows.ColumnCount);
            writer.WriteRecord(BrtBeginSheetData);

            if (source.IncludeHeaders && rows.ColumnCount != 0) {
                WriteDirectRowHeader(writer, 0, rows.ColumnCount);
                for (int column = 0; column < rows.ColumnCount; column++) {
                    WriteDirectTextCell(writer, column, rows.GetColumnName(column));
                }
            }

            for (int row = 0; row < rows.RowCount; row++) {
                if ((row & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                int zeroBasedRow = row + rowOffset;
                if (rows.ColumnCount != 0) {
                    WriteDirectRowHeader(writer, zeroBasedRow, rows.ColumnCount);
                }
                for (int column = 0; column < rows.ColumnCount; column++) {
                    ExcelDirectTabularValue value = ExcelDirectTabularValue.Normalize(rows.GetValue(row, column));
                    switch (value.Kind) {
                        case ExcelDirectTabularValueKind.Empty:
                            break;
                        case ExcelDirectTabularValueKind.Text:
                            WriteDirectTextCell(writer, column, value.Text ?? string.Empty);
                            break;
                        case ExcelDirectTabularValueKind.Number:
                            WriteDirectNumberCell(writer, column, value.Number);
                            break;
                        case ExcelDirectTabularValueKind.Boolean:
                            WriteDirectBooleanCell(writer, column, value.Boolean);
                            break;
                        default:
                            worksheetPart = default;
                            return false;
                    }
                }
            }

            writer.WriteRecord(BrtEndSheetData);
            writer.WriteRecord(BrtEndSheet);
            worksheetPart = new ArraySegment<byte>(output.GetBuffer(), 0, checked((int)output.Length));
            return true;
        }

        private static void WriteDirectDimension(XlsbDirectRecordWriter writer, int rowCount, int columnCount) {
            int lastRow = Math.Max(1, rowCount);
            int lastColumn = Math.Max(1, columnCount);
            writer.WriteUInt32(0U);
            writer.WriteUInt32(checked((uint)(lastRow - 1)));
            writer.WriteUInt32(0U);
            writer.WriteUInt32(checked((uint)(lastColumn - 1)));
        }

        private static void WriteDirectRowHeader(
            XlsbDirectRecordWriter writer,
            int zeroBasedRow,
            int columnCount) {
            writer.WriteRowHeader(BrtRowHdr, zeroBasedRow, columnCount, DefaultRowProperties);
        }

        private static void WriteDirectTextCell(
            XlsbDirectRecordWriter writer,
            int zeroBasedColumn,
            string value) {
            CoerceValueHelper.ValidateSharedStringLength(value, nameof(value));
            writer.WriteTextCell(BrtCellSt, zeroBasedColumn, value);
        }

        private static void WriteDirectNumberCell(XlsbDirectRecordWriter writer, int zeroBasedColumn, double value) {
            writer.WriteNumberCell(BrtCellReal, zeroBasedColumn, value);
        }

        private static void WriteDirectBooleanCell(XlsbDirectRecordWriter writer, int zeroBasedColumn, bool value) {
            writer.WriteBooleanCell(BrtCellBool, zeroBasedColumn, value);
        }

        private static int EstimateDirectWorksheetCapacity(int rowCount, int columnCount) {
            const int maximumInitialCapacity = 16 * 1024 * 1024;
            int spanCount = checked((columnCount + 1023) / 1024);
            // Bound the dense-cell assumption so a very wide, sparse table does not
            // reserve a large buffer before the actual values establish its density.
            int estimatedDenseColumns = Math.Min(columnCount, 128);
            long bytesPerRow = 19L + spanCount * 8L + estimatedDenseColumns * 24L;
            long estimate = 64L + rowCount * bytesPerRow;
            return (int)Math.Max(256L, Math.Min(maximumInitialCapacity, estimate));
        }

        internal static byte[] Rewrite(byte[] originalPart, IReadOnlyList<XlsbWriteCell> cells) =>
            Rewrite(
                originalPart,
                cells,
                Array.Empty<XlsbGeneratedRecord>(),
                rewriteAutoFilter: false,
                Array.Empty<XlsbGeneratedRecord>(),
                rewriteHyperlinks: false);

        internal static byte[] Rewrite(
            byte[] originalPart,
            IReadOnlyList<XlsbWriteCell> cells,
            IReadOnlyList<XlsbGeneratedRecord> autoFilterRecords,
            bool rewriteAutoFilter,
            IReadOnlyList<XlsbGeneratedRecord> hyperlinkRecords,
            bool rewriteHyperlinks) {
            if (originalPart == null) throw new ArgumentNullException(nameof(originalPart));
            if (cells == null) throw new ArgumentNullException(nameof(cells));
            if (autoFilterRecords == null) throw new ArgumentNullException(nameof(autoFilterRecords));
            if (hyperlinkRecords == null) throw new ArgumentNullException(nameof(hyperlinkRecords));

            IReadOnlyList<XlsbRecord> records;
            using (var input = new MemoryStream(originalPart, writable: false)) {
                records = XlsbRecordReader.ReadAll(input);
            }

            int beginIndex = FindSingleRecord(records, BrtBeginSheetData, "BrtBeginSheetData");
            int endIndex = FindSingleRecord(records, BrtEndSheetData, "BrtEndSheetData");
            if (endIndex <= beginIndex) {
                throw new InvalidDataException("The XLSB worksheet has an invalid sheet-data boundary order.");
            }

            XlsbSheetDataLayout layout = ParseSheetDataLayout(records, beginIndex + 1, endIndex);
            IReadOnlyDictionary<int, IReadOnlyList<XlsbWriteCell>> cellsByRow = cells
                .GroupBy(cell => cell.Row - 1)
                .ToDictionary(group => group.Key, group => (IReadOnlyList<XlsbWriteCell>)group.OrderBy(cell => cell.Column).ToArray());
            int[] rowIndexes = layout.Rows.Keys.Concat(cellsByRow.Keys).Distinct().OrderBy(row => row).ToArray();
            byte[]? dimensionPayload = CreateExpandedDimensionPayload(records, cells);
            int hyperlinkInsertionIndex = rewriteHyperlinks
                ? FindHyperlinkInsertionIndex(records, endIndex)
                : -1;
            (int Begin, int End) autoFilterBounds = rewriteAutoFilter
                ? FindAutoFilterBounds(records, endIndex)
                : (-1, -1);

            using var output = new MemoryStream(originalPart.Length + Math.Max(256, cells.Count * 24));
            for (int index = 0; index <= beginIndex; index++) {
                if (records[index].Type == BrtWsDim && dimensionPayload != null) {
                    XlsbRecordWriter.Write(output, BrtWsDim, dimensionPayload);
                } else {
                    WriteRecord(output, records[index]);
                }
            }

            foreach (XlsbRecord metadata in layout.PrefixRecords) {
                WriteRecord(output, metadata);
            }

            foreach (int rowIndex in rowIndexes) {
                layout.Rows.TryGetValue(rowIndex, out XlsbSourceRowBlock? sourceRow);
                cellsByRow.TryGetValue(rowIndex, out IReadOnlyList<XlsbWriteCell>? rowCells);
                byte[] rowPayload = CreateRowHeaderPayload(rowIndex, sourceRow?.RowHeader.Data, rowCells ?? Array.Empty<XlsbWriteCell>(), newProperties: null);
                XlsbRecordWriter.Write(output, BrtRowHdr, rowPayload);

                if (sourceRow != null) {
                    WriteSourceRowContents(output, sourceRow, rowCells ?? Array.Empty<XlsbWriteCell>());
                } else if (rowCells != null) {
                    foreach (XlsbWriteCell cell in rowCells) WriteCell(output, cell);
                }
            }

            for (int index = endIndex; index < records.Count; index++) {
                if (rewriteAutoFilter && index == autoFilterBounds.Begin) {
                    foreach (XlsbGeneratedRecord autoFilter in autoFilterRecords) {
                        XlsbRecordWriter.Write(output, autoFilter.Type, autoFilter.Payload);
                    }
                }
                if (rewriteAutoFilter
                    && index >= autoFilterBounds.Begin
                    && index <= autoFilterBounds.End) {
                    continue;
                }
                if (rewriteHyperlinks && index == hyperlinkInsertionIndex) {
                    foreach (XlsbGeneratedRecord hyperlink in hyperlinkRecords) {
                        XlsbRecordWriter.Write(output, hyperlink.Type, hyperlink.Payload);
                    }
                }
                if (rewriteHyperlinks && records[index].Type == BrtHLink) continue;
                WriteRecord(output, records[index]);
            }

            return output.ToArray();
        }

        private static (int Begin, int End) FindAutoFilterBounds(
            IReadOnlyList<XlsbRecord> records,
            int endSheetDataIndex) {
            const int BrtBeginAFilter = 161;
            const int BrtEndAFilter = 162;
            int begin = FindSingleRecord(records, BrtBeginAFilter, "BrtBeginAFilter");
            int end = FindSingleRecord(records, BrtEndAFilter, "BrtEndAFilter");
            int endSheet = FindSingleRecord(records, BrtEndSheet, "BrtEndSheet");
            if (begin <= endSheetDataIndex || end < begin || end >= endSheet) {
                throw new InvalidDataException("The XLSB worksheet has an invalid AutoFilter record boundary order.");
            }
            return (begin, end);
        }

        private static int FindHyperlinkInsertionIndex(IReadOnlyList<XlsbRecord> records, int endSheetDataIndex) {
            int endSheetIndex = FindSingleRecord(records, BrtEndSheet, "BrtEndSheet");
            int firstHyperlinkIndex = -1;
            for (int index = 0; index < records.Count; index++) {
                if (records[index].Type != BrtHLink) continue;
                if (index <= endSheetDataIndex || index >= endSheetIndex) {
                    throw new InvalidDataException("The XLSB worksheet contains a BrtHLink record outside the supported worksheet-metadata region.");
                }
                if (firstHyperlinkIndex < 0) firstHyperlinkIndex = index;
            }
            if (firstHyperlinkIndex >= 0) return firstHyperlinkIndex;

            for (int index = endSheetDataIndex + 1; index < endSheetIndex; index++) {
                int type = records[index].Type;
                if (IsHyperlinkSuccessorRecord(type)) {
                    return index;
                }
            }
            return endSheetIndex;
        }

        private static bool IsHyperlinkSuccessorRecord(int type) =>
            type == BrtPrintOptions
            || type == BrtMargins
            || type == BrtPageSetup
            || type == BrtBeginHeaderFooter
            || type == BrtEndHeaderFooter
            || type == BrtBeginRowBreaks
            || type == BrtBeginColumnBreaks
            || type == BrtBigName
            || type == BrtBeginCellWatches
            || type == BrtBeginCellIgnoreErrors
            || type == BrtBeginSmartTags
            || type == BrtDrawing
            || type == BrtLegacyDrawing
            || type == BrtLegacyDrawingHeaderFooter
            || type == BrtBackgroundImage
            || type == BrtBeginOleObjects
            || type == BrtBeginActiveXControls
            || type == BrtBeginWebPublishItems
            || type == BrtBeginTableParts
            || type == BrtBeginDataValidations
            || type == BrtEndSheet;

        private static byte[] CreateRowHeaderPayload(
            int zeroBasedRow,
            byte[]? sourcePayload,
            IReadOnlyList<XlsbWriteCell> cells,
            byte[]? newProperties) {
            if (sourcePayload != null && sourcePayload.Length < 17) {
                throw new InvalidDataException($"The XLSB row header for row {zeroBasedRow + 1} is truncated.");
            }

            if (sourcePayload == null
                && newProperties == null
                && cells.Count > 0
                && (cells[0].Column - 1) / 1024 == (cells[cells.Count - 1].Column - 1) / 1024) {
                var compactPayload = new byte[25];
                WriteUInt32(compactPayload, 0, checked((uint)zeroBasedRow));
                Buffer.BlockCopy(DefaultRowProperties, 0, compactPayload, 4, DefaultRowProperties.Length);
                WriteUInt32(compactPayload, 13, 1);
                WriteUInt32(compactPayload, 17, checked((uint)(cells[0].Column - 1)));
                WriteUInt32(compactPayload, 21, checked((uint)(cells[cells.Count - 1].Column - 1)));
                return compactPayload;
            }

            var spans = ReadSourceSpans(sourcePayload, zeroBasedRow);
            foreach (IGrouping<int, XlsbWriteCell> group in cells.GroupBy(cell => (cell.Column - 1) / 1024)) {
                uint first = checked((uint)(group.Min(cell => cell.Column) - 1));
                uint last = checked((uint)(group.Max(cell => cell.Column) - 1));
                if (spans.TryGetValue(group.Key, out (uint First, uint Last) sourceSpan)) {
                    first = Math.Min(first, sourceSpan.First);
                    last = Math.Max(last, sourceSpan.Last);
                }
                spans[group.Key] = (first, last);
            }
            if (spans.Count > 16) {
                throw new InvalidDataException($"The XLSB row {zeroBasedRow + 1} requires {spans.Count} column spans, exceeding the BIFF12 limit of 16.");
            }

            using var payload = new MemoryStream(17 + spans.Count * 8);
            WriteUInt32(payload, checked((uint)zeroBasedRow));
            if (sourcePayload != null) {
                payload.Write(sourcePayload, 4, 9);
            } else if (newProperties != null) {
                if (newProperties.Length != 9) throw new InvalidDataException("A generated XLSB row-property payload must contain 9 bytes.");
                payload.Write(newProperties, 0, newProperties.Length);
            } else {
                payload.Write(DefaultRowProperties, 0, DefaultRowProperties.Length);
            }
            WriteUInt32(payload, checked((uint)spans.Count));
            foreach (KeyValuePair<int, (uint First, uint Last)> span in spans.OrderBy(pair => pair.Key)) {
                WriteUInt32(payload, span.Value.First);
                WriteUInt32(payload, span.Value.Last);
            }
            return payload.ToArray();
        }

        private static Dictionary<int, (uint First, uint Last)> ReadSourceSpans(byte[]? sourcePayload, int zeroBasedRow) {
            var spans = new Dictionary<int, (uint First, uint Last)>();
            if (sourcePayload == null) return spans;

            var cursor = new XlsbBinaryCursor(sourcePayload);
            cursor.Skip(13);
            uint count = cursor.ReadUInt32();
            if (count > 16 || cursor.Remaining != checked((int)count * 8)) {
                throw new InvalidDataException($"The XLSB row header for row {zeroBasedRow + 1} has an invalid column-span payload.");
            }
            for (uint index = 0; index < count; index++) {
                uint first = cursor.ReadUInt32();
                uint last = cursor.ReadUInt32();
                int segment = checked((int)(first / 1024U));
                if (first > last || last >= 16_384U || first / 1024U != last / 1024U || spans.ContainsKey(segment)) {
                    throw new InvalidDataException($"The XLSB row header for row {zeroBasedRow + 1} contains an invalid column span.");
                }
                spans.Add(segment, (first, last));
            }
            return spans;
        }

        private static byte[]? CreateExpandedDimensionPayload(
            IReadOnlyList<XlsbRecord> records,
            IReadOnlyList<XlsbWriteCell> cells) {
            XlsbRecord? dimension = null;
            foreach (XlsbRecord record in records) {
                if (record.Type != BrtWsDim) continue;
                if (dimension != null) {
                    throw new InvalidDataException("The XLSB worksheet contains more than one BrtWsDim record.");
                }
                dimension = record;
            }

            if (dimension == null) return null;
            if (dimension.Data.Length != 16) {
                throw new InvalidDataException($"The XLSB BrtWsDim record has invalid payload length {dimension.Data.Length}.");
            }
            if (cells.Count == 0) return (byte[])dimension.Data.Clone();

            var cursor = new XlsbBinaryCursor(dimension.Data);
            uint firstRow = cursor.ReadUInt32();
            uint lastRow = cursor.ReadUInt32();
            uint firstColumn = cursor.ReadUInt32();
            uint lastColumn = cursor.ReadUInt32();
            uint cellFirstRow = checked((uint)(cells.Min(cell => cell.Row) - 1));
            uint cellLastRow = checked((uint)(cells.Max(cell => cell.Row) - 1));
            uint cellFirstColumn = checked((uint)(cells.Min(cell => cell.Column) - 1));
            uint cellLastColumn = checked((uint)(cells.Max(cell => cell.Column) - 1));
            bool hasSourceCells = records.Any(record => IsCellRecord(record.Type));

            using var payload = new MemoryStream(16);
            WriteUInt32(payload, hasSourceCells ? Math.Min(firstRow, cellFirstRow) : cellFirstRow);
            WriteUInt32(payload, hasSourceCells ? Math.Max(lastRow, cellLastRow) : cellLastRow);
            WriteUInt32(payload, hasSourceCells ? Math.Min(firstColumn, cellFirstColumn) : cellFirstColumn);
            WriteUInt32(payload, hasSourceCells ? Math.Max(lastColumn, cellLastColumn) : cellLastColumn);
            return payload.ToArray();
        }

        private static XlsbSheetDataLayout ParseSheetDataLayout(IReadOnlyList<XlsbRecord> records, int start, int end) {
            var prefix = new List<XlsbRecord>();
            var rows = new Dictionary<int, XlsbSourceRowBlock>();
            XlsbSourceRowBlock? current = null;
            for (int index = start; index < end; index++) {
                XlsbRecord record = records[index];
                if (record.Type == BrtRowHdr) {
                    var cursor = new XlsbBinaryCursor(record.Data);
                    int rowIndex = cursor.ReadInt32();
                    if (rowIndex < 0 || rowIndex >= 1_048_576 || rows.ContainsKey(rowIndex)) {
                        throw new InvalidDataException($"The XLSB worksheet contains invalid or duplicate row index {rowIndex}.");
                    }

                    current = new XlsbSourceRowBlock(record);
                    rows.Add(rowIndex, current);
                } else if (current == null) {
                    if (IsCellRecord(record.Type)) {
                        throw new InvalidDataException("The XLSB worksheet contains a cell record before its row header.");
                    }

                    prefix.Add(record);
                } else {
                    if (IsCellRecord(record.Type)) {
                        current.AddCell(record, ReadCellColumn(record));
                    } else {
                        current.AddPreserved(record);
                    }
                }
            }

            return new XlsbSheetDataLayout(prefix, rows);
        }

        private static void WriteSourceRowContents(
            Stream output,
            XlsbSourceRowBlock sourceRow,
            IReadOnlyList<XlsbWriteCell> rowCells) {
            int nextCell = 0;
            for (int itemIndex = 0; itemIndex < sourceRow.Items.Count; itemIndex++) {
                XlsbSourceRowItem item = sourceRow.Items[itemIndex];
                if (sourceRow.LastCellItemIndex >= 0
                    && itemIndex > sourceRow.LastCellItemIndex
                    && nextCell < rowCells.Count) {
                    while (nextCell < rowCells.Count) WriteCell(output, rowCells[nextCell++]);
                }

                if (!item.Column.HasValue) {
                    WriteRecord(output, item.Record);
                    continue;
                }

                int sourceColumn = item.Column.Value;
                while (nextCell < rowCells.Count && rowCells[nextCell].Column < sourceColumn) {
                    WriteCell(output, rowCells[nextCell++]);
                }
                if (nextCell < rowCells.Count && rowCells[nextCell].Column == sourceColumn) {
                    WriteCell(output, rowCells[nextCell++]);
                }
            }

            while (nextCell < rowCells.Count) WriteCell(output, rowCells[nextCell++]);
        }

        private static int ReadCellColumn(XlsbRecord record) {
            if (record.Data.Length < 4) {
                throw new InvalidDataException($"The XLSB cell record {record.Type} is truncated.");
            }

            var cursor = new XlsbBinaryCursor(record.Data);
            uint zeroBasedColumn = cursor.ReadUInt32();
            if (zeroBasedColumn >= 16_384U) {
                throw new InvalidDataException($"The XLSB cell record {record.Type} has invalid column index {zeroBasedColumn}.");
            }
            return checked((int)zeroBasedColumn + 1);
        }

        private static void WriteCell(Stream output, XlsbWriteCell cell) {
            if (cell.SourceRecordType.HasValue && cell.SourceRecordData != null) {
                XlsbRecordWriter.Write(output, cell.SourceRecordType.Value, cell.SourceRecordData);
                return;
            }

            switch (cell.Kind) {
                case XlsbWriteCellKind.Blank:
                    WriteSimpleCellHeader(output, BrtCellBlank, payloadLength: 8, cell);
                    return;
                case XlsbWriteCellKind.Number:
                    WriteSimpleCellHeader(output, BrtCellReal, payloadLength: 16, cell);
                    WriteDouble(output, Convert.ToDouble(cell.Value, System.Globalization.CultureInfo.InvariantCulture));
                    return;
                case XlsbWriteCellKind.Text:
                    string text = (string?)cell.Value ?? string.Empty;
                    int textPayloadLength = checked(12 + (text.Length * 2));
                    WriteSimpleCellHeader(output, BrtCellSt, textPayloadLength, cell);
                    WriteWideString(output, text);
                    return;
                case XlsbWriteCellKind.Boolean:
                    WriteSimpleCellHeader(output, BrtCellBool, payloadLength: 9, cell);
                    output.WriteByte((bool)cell.Value! ? (byte)1 : (byte)0);
                    return;
                case XlsbWriteCellKind.Error:
                    WriteSimpleCellHeader(output, BrtCellError, payloadLength: 9, cell);
                    output.WriteByte((byte)cell.Value!);
                    return;
            }

            using var payload = new MemoryStream();
            WriteUInt32(payload, checked((uint)(cell.Column - 1)));
            WriteUInt32(payload, cell.StyleIndex & 0x00FFFFFFU);
            int recordType;
            switch (cell.Kind) {
                case XlsbWriteCellKind.FormulaNumber:
                    recordType = BrtFmlaNum;
                    WriteDouble(payload, Convert.ToDouble(cell.Value, System.Globalization.CultureInfo.InvariantCulture));
                    WriteFormula(payload, cell.FormulaPayload);
                    break;
                case XlsbWriteCellKind.FormulaText:
                    recordType = BrtFmlaString;
                    WriteWideString(payload, (string?)cell.Value ?? string.Empty);
                    WriteFormula(payload, cell.FormulaPayload);
                    break;
                case XlsbWriteCellKind.FormulaBoolean:
                    recordType = BrtFmlaBool;
                    payload.WriteByte((bool)cell.Value! ? (byte)1 : (byte)0);
                    WriteFormula(payload, cell.FormulaPayload);
                    break;
                case XlsbWriteCellKind.FormulaError:
                    recordType = BrtFmlaError;
                    payload.WriteByte((byte)cell.Value!);
                    WriteFormula(payload, cell.FormulaPayload);
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported XLSB write cell kind {cell.Kind}.");
            }

            XlsbRecordWriter.Write(output, recordType, payload.ToArray());
        }

        private static void WriteSimpleCellHeader(Stream output, int recordType, int payloadLength, XlsbWriteCell cell) {
            XlsbRecordWriter.WriteHeader(output, recordType, payloadLength);
            WriteUInt32(output, checked((uint)(cell.Column - 1)));
            WriteUInt32(output, cell.StyleIndex & 0x00FFFFFFU);
        }

        private static void WriteFormula(Stream payload, byte[]? formulaPayload) {
            byte[] bytes = formulaPayload ?? throw new InvalidOperationException("Formula cell has no preserved BIFF12 formula payload.");
            payload.Write(bytes, 0, bytes.Length);
        }

        private static void WriteWideString(Stream stream, string value) {
            WriteUInt32(stream, checked((uint)value.Length));
            for (int index = 0; index < value.Length; index++) {
                ushort character = value[index];
                stream.WriteByte((byte)character);
                stream.WriteByte((byte)(character >> 8));
            }
        }

        private static void WriteDouble(Stream stream, double value) {
            ulong bits = unchecked((ulong)BitConverter.DoubleToInt64Bits(value));
            WriteUInt32(stream, unchecked((uint)bits));
            WriteUInt32(stream, unchecked((uint)(bits >> 32)));
        }

        private static void WriteUInt32(byte[] buffer, int offset, uint value) {
            buffer[offset] = (byte)value;
            buffer[offset + 1] = (byte)(value >> 8);
            buffer[offset + 2] = (byte)(value >> 16);
            buffer[offset + 3] = (byte)(value >> 24);
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
            stream.WriteByte((byte)(value >> 16));
            stream.WriteByte((byte)(value >> 24));
        }

        private static void WriteRecord(Stream stream, XlsbRecord record) =>
            XlsbRecordWriter.Write(stream, record.Type, record.Data);

        private static int FindSingleRecord(IReadOnlyList<XlsbRecord> records, int recordType, string recordName) {
            int found = -1;
            for (int index = 0; index < records.Count; index++) {
                if (records[index].Type != recordType) continue;
                if (found >= 0) throw new InvalidDataException($"The XLSB worksheet contains more than one {recordName} record.");
                found = index;
            }

            if (found < 0) throw new InvalidDataException($"The XLSB worksheet does not contain a {recordName} record.");
            return found;
        }

        private static bool IsCellRecord(int recordType) {
            return (recordType >= BrtCellBlank && recordType <= BrtFmlaError)
                || recordType == BrtCellRString;
        }

        private sealed class XlsbSheetDataLayout {
            internal XlsbSheetDataLayout(List<XlsbRecord> prefixRecords, Dictionary<int, XlsbSourceRowBlock> rows) {
                PrefixRecords = prefixRecords;
                Rows = rows;
            }

            internal IReadOnlyList<XlsbRecord> PrefixRecords { get; }

            internal IReadOnlyDictionary<int, XlsbSourceRowBlock> Rows { get; }
        }

        private sealed class XlsbSourceRowBlock {
            internal XlsbSourceRowBlock(XlsbRecord rowHeader) {
                RowHeader = rowHeader;
            }

            internal XlsbRecord RowHeader { get; }

            internal List<XlsbSourceRowItem> Items { get; } = new List<XlsbSourceRowItem>();

            internal int LastCellItemIndex { get; private set; } = -1;

            internal void AddCell(XlsbRecord record, int column) {
                if (LastCellItemIndex >= 0) {
                    int previousColumn = Items[LastCellItemIndex].Column!.Value;
                    if (column <= previousColumn) {
                        throw new InvalidDataException("The XLSB worksheet row contains duplicate or out-of-order cell records.");
                    }
                }

                Items.Add(new XlsbSourceRowItem(record, column));
                LastCellItemIndex = Items.Count - 1;
            }

            internal void AddPreserved(XlsbRecord record) =>
                Items.Add(new XlsbSourceRowItem(record, column: null));
        }

        private sealed class XlsbSourceRowItem {
            internal XlsbSourceRowItem(XlsbRecord record, int? column) {
                Record = record;
                Column = column;
            }

            internal XlsbRecord Record { get; }

            internal int? Column { get; }
        }
    }
}
