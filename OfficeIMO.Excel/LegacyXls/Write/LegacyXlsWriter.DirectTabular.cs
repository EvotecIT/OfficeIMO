using System.Buffers;
using System.Diagnostics;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        private const ushort DirectDefaultCellStyleIndex = 15;
        private const int DirectRowsPerDbCellBlock = 32;

        private static bool TryCreateDirectTabularPlan(
            ExcelDirectTabularSource source,
            CancellationToken cancellationToken,
            out DirectTabularPlan plan) {
            IExcelSheetTabularRowSource rows = source.Rows;
            int rowOffset = source.IncludeHeaders ? 1 : 0;
            int totalRows = checked(rows.RowCount + rowOffset);
            int columnCount = rows.ColumnCount;
            if (totalRows > 65_536 || columnCount > 256) {
                throw new NotSupportedException("Native XLS saving supports the BIFF8 worksheet limit of 65,536 rows and 256 columns.");
            }

            var firstColumns = new ushort[totalRows];
            for (int row = 0; row < firstColumns.Length; row++) {
                firstColumns[row] = ushort.MaxValue;
            }
            var lastColumns = new ushort[totalRows];
            var sharedStrings = new LegacyXlsDirectSharedStringBuilder();
            int cellSlotCount = checked(totalRows * columnCount);
            byte[] cellKinds = cellSlotCount == 0
                ? Array.Empty<byte>()
                : ArrayPool<byte>.Shared.Rent(cellSlotCount);
            ulong[] cellPayloads = cellSlotCount == 0
                ? Array.Empty<ulong>()
                : ArrayPool<ulong>.Shared.Rent(cellSlotCount);
            bool returnCellCache = cellSlotCount != 0;
            object?[]? flatValues = rows.TryGetFlatValues(out object?[] candidateValues, out int flatColumnCount)
                && flatColumnCount == columnCount
                && candidateValues.Length == checked(rows.RowCount * columnCount)
                    ? candidateValues
                    : null;
            int nonEmptyCellCount = 0;
            uint? firstDataRow = null;
            uint? lastDataRow = null;

            try {
                if (source.IncludeHeaders) {
                    for (int column = 0; column < columnCount; column++) {
                        int cellSlot = column;
                        cellKinds[cellSlot] = (byte)ExcelDirectTabularValueKind.Text;
                        cellPayloads[cellSlot] = sharedStrings.Add(rows.GetColumnName(column));
                        IncludeDirectCell(firstColumns, lastColumns, 0, checked((ushort)column));
                        nonEmptyCellCount++;
                    }
                    if (columnCount != 0) {
                        firstDataRow = 0;
                        lastDataRow = 0;
                    }
                }

                bool canCancel = cancellationToken.CanBeCanceled;
                for (int row = 0; row < rows.RowCount; row++) {
                    if (canCancel && (row & 1023) == 0) {
                        cancellationToken.ThrowIfCancellationRequested();
                    }

                    int directRow = row + rowOffset;
                    object?[]? bufferedRow = flatValues == null
                        && rows.TryGetBufferedRow(row, out object?[]? candidateRow)
                        && candidateRow?.Length == columnCount
                            ? candidateRow
                            : null;
                    for (int column = 0; column < columnCount; column++) {
                        object? rawValue = flatValues != null
                            ? flatValues[checked((row * columnCount) + column)]
                            : bufferedRow != null
                                ? bufferedRow[column]
                                : rows.GetValue(row, column);
                        ExcelDirectTabularValue value = ExcelDirectTabularValue.Normalize(rawValue);
                        if (value.Kind == ExcelDirectTabularValueKind.Unsupported) {
                            plan = null!;
                            return false;
                        }

                        int cellSlot = checked((directRow * columnCount) + column);
                        cellKinds[cellSlot] = (byte)value.Kind;
                        cellPayloads[cellSlot] = value.Kind switch {
                            ExcelDirectTabularValueKind.Text => sharedStrings.Add(value.Text ?? string.Empty),
                            ExcelDirectTabularValueKind.Boolean => value.Boolean ? 1UL : 0UL,
                            ExcelDirectTabularValueKind.Number => unchecked((ulong)BitConverter.DoubleToInt64Bits(value.Number)),
                            _ => 0UL
                        };
                        if (value.Kind == ExcelDirectTabularValueKind.Empty) {
                            continue;
                        }

                        nonEmptyCellCount++;
                        IncludeDirectCell(firstColumns, lastColumns, directRow, checked((ushort)column));
                        uint dataRow = checked((uint)directRow);
                        firstDataRow = firstDataRow.HasValue ? Math.Min(firstDataRow.Value, dataRow) : dataRow;
                        lastDataRow = lastDataRow.HasValue ? Math.Max(lastDataRow.Value, dataRow) : dataRow;
                    }
                }

                LegacyXlsDimensions dimensions = totalRows == 0 || columnCount == 0
                    ? default
                    : new LegacyXlsDimensions(0, checked((uint)totalRows), 0, checked((ushort)columnCount));
                if (totalRows != 0 && columnCount != 0) {
                    IncludeDirectCell(firstColumns, lastColumns, 0, 0);
                    IncludeDirectCell(
                        firstColumns,
                        lastColumns,
                        totalRows - 1,
                        checked((ushort)(columnCount - 1)));
                }

                int populatedRowCount = 0;
                for (int row = 0; row < firstColumns.Length; row++) {
                    if (firstColumns[row] != ushort.MaxValue) populatedRowCount++;
                }
                long estimatedCapacity = checked(
                    16_384L
                    + ((long)(nonEmptyCellCount + 2) * 18L)
                    + ((long)populatedRowCount * 22L)
                    + ((long)dimensions.RowBlockCount * 8L)
                    + sharedStrings.EstimatedSerializedBytes);
                int workbookStreamCapacity = estimatedCapacity >= int.MaxValue
                    ? int.MaxValue
                    : checked((int)estimatedCapacity);

                plan = new DirectTabularPlan(
                    totalRows,
                    columnCount,
                    cellSlotCount,
                    cellKinds,
                    cellPayloads,
                    firstColumns,
                    lastColumns,
                    dimensions,
                    firstDataRow,
                    lastDataRow,
                    nonEmptyCellCount,
                    populatedRowCount,
                    workbookStreamCapacity,
                    sharedStrings.Build());
                returnCellCache = false;
                return true;
            } finally {
                if (returnCellCache) {
                    ArrayPool<byte>.Shared.Return(cellKinds);
                    ArrayPool<ulong>.Shared.Return(cellPayloads);
                }
            }
        }

        private static void IncludeDirectCell(
            ushort[] firstColumns,
            ushort[] lastColumns,
            int row,
            ushort column) {
            if (firstColumns[row] == ushort.MaxValue) {
                firstColumns[row] = column;
                lastColumns[row] = column;
                return;
            }

            if (column < firstColumns[row]) firstColumns[row] = column;
            if (column > lastColumns[row]) lastColumns[row] = column;
        }

        private static DirectTabularWorkbookStream BuildDirectTabularWorkbookStream(
            ExcelDocument document,
            ExcelSheet sheet,
            DirectTabularPlan plan) {
            Stopwatch? stageWatch = document.Execution.OnTiming == null ? null : Stopwatch.StartNew();
            LegacyXlsFontTable fontTable = LegacyXlsFontTable.Create(document);
            LegacyXlsStyleTable styleTable = LegacyXlsStyleTable.CreateDirectTabular(document, fontTable);
            LegacyXlsExternSheetTable externSheetTable = LegacyXlsExternSheetTable.CreateDirectTabular(sheet.Name);
            LegacyXlsSharedStringTable sharedStrings = plan.SharedStrings;
            ReportDirectWriteTiming(document, stageWatch, "Save.Xls.Direct.BuildWorkbookStream.CreateTables");

            var stream = new DirectTabularWorkbookStream(plan.WorkbookStreamCapacity);
            try {
                long boundSheetPosition = WriteDirectWorkbookGlobals(
                    stream,
                    document,
                    sheet,
                    fontTable,
                    styleTable,
                    externSheetTable,
                    sharedStrings);
                ReportDirectWriteTiming(document, stageWatch, "Save.Xls.Direct.BuildWorkbookStream.WriteGlobals");

                int worksheetOffset = checked((int)stream.Position);
                WriteDirectTabularWorksheet(stream, plan);
                ReportDirectWriteTiming(document, stageWatch, "Save.Xls.Direct.BuildWorkbookStream.WriteWorksheet");
                long endPosition = stream.Position;
                stream.Position = boundSheetPosition + 4;
                WriteUInt32(stream, unchecked((uint)worksheetOffset));
                stream.Position = endPosition;
                ReportDirectWriteTiming(document, stageWatch, "Save.Xls.Direct.BuildWorkbookStream.Materialize");
                return stream;
            } catch {
                stream.Dispose();
                throw;
            }
        }

        private static long WriteDirectWorkbookGlobals(
            Stream stream,
            ExcelDocument document,
            ExcelSheet sheet,
            LegacyXlsFontTable fontTable,
            LegacyXlsStyleTable styleTable,
            LegacyXlsExternSheetTable externSheetTable,
            LegacyXlsSharedStringTable sharedStrings) {
            WriteRecord(stream, 0x0809, WorkbookGlobalsBof);
            WriteRecord(stream, 0x00e1, BuildUInt16Payload(1200));
            WriteRecord(stream, 0x00c1, BuildUInt16Payload(0));
            WriteRecord(stream, 0x00e2, Array.Empty<byte>());
            WriteRecord(stream, 0x005c, BuildWriteAccessPayload("OfficeIMO"));
            WriteRecord(stream, 0x0042, BuildUInt16Payload(1200));
            WriteRecord(stream, 0x0161, BuildUInt16Payload(0));
            WriteRecord(stream, 0x013d, BuildSheetTabIdsPayload(document, [sheet]));
            WriteRecord(stream, 0x009c, BuildUInt16Payload(14));
            WriteRecord(stream, 0x0019, BuildUInt16Payload(0));
            WriteRecord(stream, 0x0012, BuildUInt16Payload(0));
            WriteRecord(stream, 0x0013, BuildUInt16Payload(0));
            WriteRecord(stream, 0x01af, BuildUInt16Payload(0));
            WriteRecord(stream, 0x01bc, BuildUInt16Payload(0));
            WriteRecord(stream, 0x003d, BuildDirectWindow1Payload());
            WriteRecord(stream, 0x0040, BuildUInt16Payload(0));
            WriteRecord(stream, 0x008d, BuildUInt16Payload(0));
            WriteRecord(stream, 0x0022, BuildUInt16Payload(
                document.DateSystem == ExcelDateSystem.NineteenFour ? (ushort)1 : (ushort)0));
            WriteRecord(stream, 0x000e, BuildUInt16Payload(1));
            WriteRecord(stream, 0x01b7, BuildUInt16Payload(0));
            WriteRecord(stream, 0x00da, BuildUInt16Payload(0));

            foreach (byte[] fontPayload in fontTable.FontRecords) {
                WriteRecord(stream, 0x0031, fontPayload);
            }
            foreach (byte[] formatPayload in styleTable.FormatRecords) {
                WriteRecord(stream, 0x041e, formatPayload);
            }
            foreach (byte[] cellFormatPayload in styleTable.CellFormatRecords) {
                WriteRecord(stream, 0x00e0, cellFormatPayload);
            }
            foreach (byte[] stylePayload in styleTable.StyleRecords) {
                WriteRecord(stream, 0x0293, stylePayload);
            }
            WriteRecord(stream, 0x0160, BuildUInt16Payload(0));

            long boundSheetPosition = stream.Position;
            WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, sheet));
            WriteRecord(stream, 0x008c, BuildCountryPayload());
            foreach (LegacyXlsExternSheetTable.SupportingLinkRecord supportingLinkRecord in externSheetTable.SupportingLinkRecords) {
                WriteRecord(stream, supportingLinkRecord.RecordType, supportingLinkRecord.Payload);
            }
            WriteRecord(stream, 0x0017, externSheetTable.Payload);
            sharedStrings.WriteRecords(stream);
            WriteRecord(stream, 0x000a, Array.Empty<byte>());
            return boundSheetPosition;
        }

        private static void WriteDirectTabularWorksheet(
            DirectTabularWorkbookStream stream,
            DirectTabularPlan plan) {
            LegacyXlsDimensions dimensions = plan.Dimensions;
            WriteDirectWorksheetPrefix(
                stream,
                dimensions,
                BuildDirectWorksheetIndexPayload(plan, dimensions.RowBlockCount),
                out long indexRecordPosition,
                out long defaultColumnWidthPosition);

            int rowBlockCount = dimensions.RowBlockCount;
            long cellTableUpperBound = checked(
                ((long)plan.PopulatedRowCount * 22L)
                + ((long)(plan.NonEmptyCellCount + 2) * 18L)
                + ((long)rowBlockCount * 8L));
            long requiredCapacity = checked(stream.Position + cellTableUpperBound + 128L);
            if (requiredCapacity > int.MaxValue) {
                throw new NotSupportedException("The direct XLS cell table exceeds the maximum in-memory workbook size.");
            }
            stream.EnsureCapacity(checked((int)requiredCapacity));

            byte[] output = stream.Buffer;
            int cellTableStart = checked((int)stream.Position);
            int writePosition = cellTableStart;
            var dbCellPositions = new uint[rowBlockCount];
            Span<int> firstCellPositions = stackalloc int[DirectRowsPerDbCellBlock];
            for (int blockIndex = 0; blockIndex < dimensions.RowBlockCount; blockIndex++) {
                int blockFirstRow = checked((int)dimensions.FirstRow + (blockIndex * DirectRowsPerDbCellBlock));
                int blockRowAfterLast = Math.Min(
                    checked(blockFirstRow + DirectRowsPerDbCellBlock),
                    checked((int)dimensions.RowAfterLast));
                int blockCount = 0;
                for (int row = blockFirstRow; row < blockRowAfterLast; row++) {
                    if (plan.FirstColumns[row] != ushort.MaxValue) blockCount++;
                }

                int firstRowPosition = writePosition;
                for (int row = blockFirstRow; row < blockRowAfterLast; row++) {
                    ushort firstColumn = plan.FirstColumns[row];
                    if (firstColumn == ushort.MaxValue) continue;
                    WriteDirectRowRecord(
                        output,
                        ref writePosition,
                        checked((ushort)row),
                        firstColumn,
                        plan.LastColumns[row]);
                }

                int cellRowIndex = 0;
                for (int row = blockFirstRow; row < blockRowAfterLast; row++) {
                    if (plan.FirstColumns[row] == ushort.MaxValue) continue;
                    firstCellPositions[cellRowIndex++] = writePosition;
                    WriteDirectRowCells(output, ref writePosition, plan, row);
                }

                int dbCellPosition = writePosition;
                dbCellPositions[blockIndex] = checked((uint)dbCellPosition);
                WriteDirectDbCellRecord(
                    output,
                    ref writePosition,
                    firstRowPosition,
                    firstCellPositions.Slice(0, blockCount),
                    dbCellPosition);
            }

            // Advance MemoryStream's logical length without clearing the direct
            // writes. Source and destination are the identical buffer range.
            stream.Write(output, cellTableStart, writePosition - cellTableStart);

            WriteRecord(stream, 0x023e, [
                0xb6, 0x06, 0x00, 0x00, 0x00, 0x00, 0x40, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
            ]);
            WriteRecord(stream, 0x001d, BuildDefaultSelectionPayload());
            WriteRecord(stream, 0x000a, Array.Empty<byte>());

            PatchIndexRecord(stream, indexRecordPosition, defaultColumnWidthPosition, dbCellPositions);
        }

        private static void WriteDirectWorksheetPrefix(
            Stream stream,
            LegacyXlsDimensions dimensions,
            byte[] indexPayload,
            out long indexRecordPosition,
            out long defaultColumnWidthPosition) {
            WriteRecord(stream, 0x0809, WorksheetBof);
            indexRecordPosition = stream.Position;
            WriteRecord(stream, 0x020b, indexPayload);
            WriteRecord(stream, 0x000d, BuildInt16Payload(1));
            WriteRecord(stream, 0x000c, BuildUInt16Payload(100));
            WriteRecord(stream, 0x000f, BuildUInt16Payload(1));
            WriteRecord(stream, 0x0011, BuildUInt16Payload(0));
            WriteRecord(stream, 0x0010, BuildDoublePayload(0.001d));
            WriteRecord(stream, 0x005f, BuildUInt16Payload(1));
            WriteRecord(stream, 0x002a, BuildUInt16Payload(0));
            WriteRecord(stream, 0x002b, BuildUInt16Payload(0));
            WriteRecord(stream, 0x0082, BuildUInt16Payload(1));
            WriteRecord(stream, 0x0080, new byte[8]);
            WriteRecord(stream, 0x0225, [0x00, 0x00, 0xff, 0x00]);
            WriteRecord(stream, 0x0081, BuildUInt16Payload(0x0104));
            WriteRecord(stream, 0x0014, Array.Empty<byte>());
            WriteRecord(stream, 0x0015, Array.Empty<byte>());
            WriteRecord(stream, 0x0083, BuildUInt16Payload(0));
            WriteRecord(stream, 0x0084, BuildUInt16Payload(0));
            WriteRecord(stream, 0x00a1, BuildDefaultDirectSetupPayload());
            defaultColumnWidthPosition = stream.Position;
            WriteRecord(stream, 0x0055, BuildUInt16Payload(8));
            WriteRecord(stream, 0x0200, BuildDimensionsPayload(dimensions));
        }

        private static byte[] BuildDirectWorksheetIndexPayload(DirectTabularPlan plan, int dbCellCount) {
            byte[] payload = new byte[checked(16 + (dbCellCount * 4))];
            if (plan.FirstDataRow.HasValue && plan.LastDataRow.HasValue) {
                WriteUInt32(payload, 4, plan.FirstDataRow.Value);
                WriteUInt32(payload, 8, checked(plan.LastDataRow.Value + 1U));
            }
            return payload;
        }

        private static void WriteDirectRowRecord(
            byte[] output,
            ref int position,
            ushort row,
            ushort firstColumn,
            ushort lastColumn) {
            WriteUInt16(output, position, 0x0208);
            WriteUInt16(output, position + 2, 16);
            WriteUInt16(output, position + 4, row);
            WriteUInt16(output, position + 6, firstColumn);
            WriteUInt16(output, position + 8, checked((ushort)(lastColumn + 1)));
            WriteUInt16(output, position + 10, 0x00ff);
            WriteUInt16(output, position + 12, 0);
            WriteUInt16(output, position + 14, 0);
            WriteUInt16(output, position + 16, 0x0100);
            WriteUInt16(output, position + 18, DirectDefaultCellStyleIndex);
            position += 20;
        }

        private static void WriteDirectRowCells(
            byte[] output,
            ref int position,
            DirectTabularPlan plan,
            int directRow) {
            for (int column = 0; column < plan.ColumnCount; column++) {
                int cellSlot = checked((directRow * plan.ColumnCount) + column);
                ExcelDirectTabularValueKind kind = (ExcelDirectTabularValueKind)plan.CellKinds[cellSlot];
                ulong payload = plan.CellPayloads[cellSlot];
                ushort legacyColumn = checked((ushort)column);
                switch (kind) {
                    case ExcelDirectTabularValueKind.Empty:
                        if ((directRow == 0 && column == 0)
                            || (directRow == plan.TotalRows - 1 && column == plan.ColumnCount - 1)) {
                            WriteDirectFixedCellHeader(output, ref position, 0x0201, 6, checked((ushort)directRow), legacyColumn);
                        }
                        break;
                    case ExcelDirectTabularValueKind.Text:
                        WriteDirectFixedCellHeader(output, ref position, 0x00fd, 10, checked((ushort)directRow), legacyColumn);
                        WriteUInt32(output, position, checked((uint)payload));
                        position += 4;
                        break;
                    case ExcelDirectTabularValueKind.Boolean:
                        WriteDirectFixedCellHeader(output, ref position, 0x0205, 8, checked((ushort)directRow), legacyColumn);
                        output[position++] = payload != 0 ? (byte)1 : (byte)0;
                        output[position++] = 0;
                        break;
                    case ExcelDirectTabularValueKind.Number:
                        WriteDirectFixedCellHeader(output, ref position, 0x0203, 14, checked((ushort)directRow), legacyColumn);
                        WriteUInt32(output, position, unchecked((uint)payload));
                        WriteUInt32(output, position + 4, unchecked((uint)(payload >> 32)));
                        position += 8;
                        break;
                    default:
                        throw new InvalidOperationException("The direct XLS tabular source changed after validation.");
                }
            }
        }

        private static void WriteDirectFixedCellHeader(
            byte[] output,
            ref int position,
            ushort type,
            ushort payloadLength,
            ushort row,
            ushort column) {
            WriteUInt16(output, position, type);
            WriteUInt16(output, position + 2, payloadLength);
            WriteUInt16(output, position + 4, row);
            WriteUInt16(output, position + 6, column);
            WriteUInt16(output, position + 8, DirectDefaultCellStyleIndex);
            position += 10;
        }

        private static void WriteDirectDbCellRecord(
            byte[] output,
            ref int position,
            int firstRowPosition,
            ReadOnlySpan<int> firstCellPositions,
            int dbCellPosition) {
            WriteUInt16(output, position, 0x00d7);
            WriteUInt16(output, position + 2, checked((ushort)(4 + (firstCellPositions.Length * 2))));
            WriteUInt32(output, position + 4, checked((uint)(dbCellPosition - firstRowPosition)));
            int priorPosition = firstRowPosition + 20;
            int offsetPosition = position + 8;
            for (int index = 0; index < firstCellPositions.Length; index++) {
                int cellPosition = firstCellPositions[index];
                WriteUInt16(output, offsetPosition, checked((ushort)(cellPosition - priorPosition)));
                priorPosition = cellPosition;
                offsetPosition += 2;
            }
            position = offsetPosition;
        }

        private static byte[] BuildDefaultDirectSetupPayload() => [
            0x01, 0x00, 0x64, 0x00, 0x01, 0x00, 0x01, 0x00, 0x01, 0x00,
            0x02, 0x00, 0x2c, 0x01, 0x2c, 0x01, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0xe0, 0x3f, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0xe0, 0x3f, 0x01, 0x00
        ];

        private static byte[] BuildDefaultSelectionPayload() => [
            0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x01,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
        ];

        private static byte[] BuildCountryPayload() => [0x01, 0x00, 0x01, 0x00];

        private static byte[] BuildDirectWindow1Payload() => [
            0x00, 0x00, 0x00, 0x00,
            0x00, 0x40, 0x00, 0x20,
            0x38, 0x00,
            0x00, 0x00, 0x00, 0x00,
            0x01, 0x00,
            0x58, 0x02
        ];

        private sealed class DirectTabularPlan : IDisposable {
            internal DirectTabularPlan(
                int totalRows,
                int columnCount,
                int cellSlotCount,
                byte[] cellKinds,
                ulong[] cellPayloads,
                ushort[] firstColumns,
                ushort[] lastColumns,
                LegacyXlsDimensions dimensions,
                uint? firstDataRow,
                uint? lastDataRow,
                int nonEmptyCellCount,
                int populatedRowCount,
                int workbookStreamCapacity,
                LegacyXlsSharedStringTable sharedStrings) {
                TotalRows = totalRows;
                ColumnCount = columnCount;
                CellSlotCount = cellSlotCount;
                CellKinds = cellKinds;
                CellPayloads = cellPayloads;
                FirstColumns = firstColumns;
                LastColumns = lastColumns;
                Dimensions = dimensions;
                FirstDataRow = firstDataRow;
                LastDataRow = lastDataRow;
                NonEmptyCellCount = nonEmptyCellCount;
                PopulatedRowCount = populatedRowCount;
                WorkbookStreamCapacity = workbookStreamCapacity;
                SharedStrings = sharedStrings;
            }

            internal int TotalRows { get; }
            internal int ColumnCount { get; }
            internal int CellSlotCount { get; }
            internal byte[] CellKinds { get; }
            internal ulong[] CellPayloads { get; }
            internal ushort[] FirstColumns { get; }
            internal ushort[] LastColumns { get; }
            internal LegacyXlsDimensions Dimensions { get; }
            internal uint? FirstDataRow { get; }
            internal uint? LastDataRow { get; }
            internal int NonEmptyCellCount { get; }
            internal int PopulatedRowCount { get; }
            internal int WorkbookStreamCapacity { get; }
            internal LegacyXlsSharedStringTable SharedStrings { get; }

            public void Dispose() {
                if (CellSlotCount == 0) return;
                ArrayPool<byte>.Shared.Return(CellKinds);
                ArrayPool<ulong>.Shared.Return(CellPayloads);
            }
        }

        private sealed class DirectTabularWorkbookStream : Stream {
            private byte[]? _buffer;
            private int _length;
            private int _position;

            internal DirectTabularWorkbookStream(int initialCapacity) {
                _buffer = ArrayPool<byte>.Shared.Rent(Math.Max(1, initialCapacity));
            }

            internal byte[] Buffer => _buffer ?? throw new ObjectDisposedException(nameof(DirectTabularWorkbookStream));

            public override bool CanRead => false;

            public override bool CanSeek => _buffer != null;

            public override bool CanWrite => _buffer != null;

            public override long Length {
                get {
                    ThrowIfDisposed();
                    return _length;
                }
            }

            public override long Position {
                get {
                    ThrowIfDisposed();
                    return _position;
                }
                set {
                    ThrowIfDisposed();
                    if (value < 0 || value > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(value));
                    _position = checked((int)value);
                }
            }

            internal void EnsureCapacity(int requiredCapacity) {
                byte[] buffer = Buffer;
                if (requiredCapacity <= buffer.Length) return;

                int doubledCapacity = buffer.Length <= int.MaxValue / 2
                    ? buffer.Length * 2
                    : int.MaxValue;
                int newCapacity = Math.Max(requiredCapacity, doubledCapacity);
                byte[] expanded = ArrayPool<byte>.Shared.Rent(newCapacity);
                System.Buffer.BlockCopy(buffer, 0, expanded, 0, _length);
                Array.Clear(buffer, 0, _length);
                _buffer = expanded;
                ArrayPool<byte>.Shared.Return(buffer);
            }

            internal byte[] ToArray() {
                var bytes = new byte[_length];
                System.Buffer.BlockCopy(Buffer, 0, bytes, 0, _length);
                return bytes;
            }

            public override void Flush() {
                ThrowIfDisposed();
            }

            public override int Read(byte[] buffer, int offset, int count) =>
                throw new NotSupportedException();

            public override long Seek(long offset, SeekOrigin origin) {
                ThrowIfDisposed();
                long target = origin switch {
                    SeekOrigin.Begin => offset,
                    SeekOrigin.Current => checked((long)_position + offset),
                    SeekOrigin.End => checked((long)_length + offset),
                    _ => throw new ArgumentOutOfRangeException(nameof(origin))
                };
                Position = target;
                return target;
            }

            public override void SetLength(long value) {
                ThrowIfDisposed();
                if (value < 0 || value > int.MaxValue) throw new ArgumentOutOfRangeException(nameof(value));
                int newLength = checked((int)value);
                EnsureCapacity(newLength);
                if (newLength > _length) {
                    Array.Clear(Buffer, _length, newLength - _length);
                }
                _length = newLength;
                if (_position > newLength) _position = newLength;
            }

            public override void Write(byte[] buffer, int offset, int count) {
                if (buffer == null) throw new ArgumentNullException(nameof(buffer));
                if (offset < 0 || count < 0 || buffer.Length - offset < count) {
                    throw new ArgumentOutOfRangeException(offset < 0 ? nameof(offset) : nameof(count));
                }

                int end = checked(_position + count);
                EnsureCapacity(end);
                System.Buffer.BlockCopy(buffer, offset, Buffer, _position, count);
                _position = end;
                if (end > _length) _length = end;
            }

            public override void WriteByte(byte value) {
                int end = checked(_position + 1);
                EnsureCapacity(end);
                Buffer[_position] = value;
                _position = end;
                if (end > _length) _length = end;
            }

            protected override void Dispose(bool disposing) {
                byte[]? buffer = _buffer;
                if (buffer != null) {
                    _buffer = null;
                    Array.Clear(buffer, 0, _length);
                    ArrayPool<byte>.Shared.Return(buffer);
                    _length = 0;
                    _position = 0;
                }
                base.Dispose(disposing);
            }

            private void ThrowIfDisposed() {
                if (_buffer == null) throw new ObjectDisposedException(nameof(DirectTabularWorkbookStream));
            }
        }
    }
}
