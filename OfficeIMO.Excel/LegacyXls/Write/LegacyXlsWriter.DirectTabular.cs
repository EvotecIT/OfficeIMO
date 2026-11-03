using OfficeIMO.Excel.LegacyXls.Biff;
using System.Diagnostics;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        private const ushort DirectDefaultCellStyleIndex = 15;
        private const int DirectRowsPerDbCellBlock = 32;
        private const int DirectMaximumCachedCellSlots = 1_048_576;

        private static DirectTabularWorkbookStream BuildDirectTabularWorkbookStream(
            ExcelDocument document,
            ExcelSheet sheet,
            bool useCellValueNumberFormats,
            DirectTabularPlan plan) {
            Stopwatch? stageWatch = document.Execution.OnTiming == null ? null : Stopwatch.StartNew();
            uint? dateStyle = plan.HasDates ? sheet.GetOrCreateDirectTabularDateStyle(useCellValueNumberFormats) : null;
            LegacyXlsFontTable fontTable = LegacyXlsFontTable.Create(document);
            LegacyXlsStyleTable styleTable = LegacyXlsStyleTable.CreateDirectTabular(document, fontTable, dateStyle);
            plan.DateStyleIndex = dateStyle.HasValue ? styleTable.GetBiffStyleIndex(dateStyle) : DirectDefaultCellStyleIndex;
            LegacyXlsExternSheetTable externSheetTable = LegacyXlsExternSheetTable.CreateDirectTabular(sheet.Name);
            LegacyXlsSharedStringTable sharedStrings = plan.SharedStrings;
            ReportDirectWriteTiming(document, stageWatch, "Save.Xls.Direct.BuildWorkbookStream.CreateTables");

            plan.CancellationToken.ThrowIfCancellationRequested();
            var stream = new DirectTabularWorkbookStream(plan.WorkbookStreamCapacity);
            try {
                long boundSheetPosition = WriteDirectWorkbookGlobals(
                    stream,
                    document,
                    sheet,
                    fontTable,
                    styleTable,
                    externSheetTable,
                    sharedStrings,
                    plan.DateSystem,
                    plan.CancellationToken);
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
            LegacyXlsSharedStringTable sharedStrings,
            ExcelDateSystem dateSystem,
            CancellationToken cancellationToken) {
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
                dateSystem == ExcelDateSystem.NineteenFour ? (ushort)1 : (ushort)0));
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
            sharedStrings.WriteRecords(stream, cancellationToken);
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
            stream.RegisterDirectWriteUpperBound(checked((int)requiredCapacity));

            byte[] output = stream.Buffer;
            int cellTableStart = checked((int)stream.Position);
            int writePosition = cellTableStart;
            var dbCellPositions = new uint[rowBlockCount];
            Span<int> firstCellPositions = stackalloc int[DirectRowsPerDbCellBlock];
            for (int blockIndex = 0; blockIndex < dimensions.RowBlockCount; blockIndex++) {
                plan.ThrowIfCancellationRequested(blockIndex * DirectRowsPerDbCellBlock);
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
                    if (plan.FirstColumns[row] == ushort.MaxValue) {
                        if (!plan.HasCachedCells) ValidateUncachedDirectRow(plan, row);
                        continue;
                    }
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
            if (!plan.HasCachedCells) {
                WriteUncachedDirectRowCells(output, ref position, plan, directRow);
                return;
            }

            int rowStart = checked(directRow * plan.ColumnCount);
            WriteDirectRowCells(output, ref position, plan, directRow,
                plan.CellKinds.AsSpan(rowStart, plan.ColumnCount),
                plan.CellPayloads.AsSpan(rowStart, plan.ColumnCount));
        }

        // Only uncached rows need the zero-initialized per-row stack buffers.
        private static void WriteUncachedDirectRowCells(
            byte[] output,
            ref int position,
            DirectTabularPlan plan,
            int directRow) {
            Span<byte> rowKinds = stackalloc byte[256];
            Span<ulong> rowPayloads = stackalloc ulong[256];
            PrepareUncachedDirectRow(plan, directRow, rowKinds, rowPayloads);
            WriteDirectRowCells(output, ref position, plan, directRow, rowKinds, rowPayloads);
        }

        private static void WriteDirectRowCells(
            byte[] output,
            ref int position,
            DirectTabularPlan plan,
            int directRow,
            ReadOnlySpan<byte> rowKinds,
            ReadOnlySpan<ulong> rowPayloads) {
            for (int column = 0; column < plan.ColumnCount; column++) {
                var kind = (ExcelDirectTabularValueKind)rowKinds[column];
                ulong payload = rowPayloads[column];
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
                    case ExcelDirectTabularValueKind.DateTime:
                        ushort numberStyle = kind == ExcelDirectTabularValueKind.DateTime ? plan.DateStyleIndex : DirectDefaultCellStyleIndex;
                        if (BiffRkNumberReader.TryEncodeRkNumber(BitConverter.Int64BitsToDouble(unchecked((long)payload)), out uint rk)) {
                            WriteDirectFixedCellHeader(output, ref position, 0x027e, 10, checked((ushort)directRow), legacyColumn, numberStyle);
                            WriteUInt32(output, position, rk);
                            position += 4;
                            break;
                        }
                        WriteDirectFixedCellHeader(output, ref position, 0x0203, 14, checked((ushort)directRow), legacyColumn, numberStyle);
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
            ushort column,
            ushort styleIndex = DirectDefaultCellStyleIndex) {
            WriteUInt16(output, position, type);
            WriteUInt16(output, position + 2, payloadLength);
            WriteUInt16(output, position + 4, row);
            WriteUInt16(output, position + 6, column);
            WriteUInt16(output, position + 8, styleIndex);
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

    }
}
