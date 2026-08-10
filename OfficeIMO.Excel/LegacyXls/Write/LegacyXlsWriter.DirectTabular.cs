namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        private const ushort DirectDefaultCellStyleIndex = 15;
        private const int DirectRowsPerDbCellBlock = 32;

        private static byte[] BuildDirectTabularWorkbookStream(
            ExcelDocument document,
            ExcelSheet sheet,
            List<LegacyXlsCell> cells) {
            LegacyXlsFontTable fontTable = LegacyXlsFontTable.Create(document);
            LegacyXlsStyleTable styleTable = LegacyXlsStyleTable.CreateDirectTabular(document, fontTable);
            LegacyXlsExternSheetTable externSheetTable = LegacyXlsExternSheetTable.CreateDirectTabular(sheet.Name);
            LegacyXlsSharedStringTable sharedStrings = LegacyXlsSharedStringTable.Create([cells]);

            using var stream = new MemoryStream();
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

            int worksheetOffset = checked((int)stream.Position);
            WriteDirectTabularWorksheet(stream, cells, sharedStrings);
            long endPosition = stream.Position;
            stream.Position = boundSheetPosition + 4;
            WriteUInt32(stream, unchecked((uint)worksheetOffset));
            stream.Position = endPosition;
            return stream.ToArray();
        }

        private static void WriteDirectTabularWorksheet(
            MemoryStream stream,
            IReadOnlyList<LegacyXlsCell> cells,
            LegacyXlsSharedStringTable sharedStrings) {
            WriteRecord(stream, 0x0809, WorksheetBof);

            DirectCellRow[] rows = cells
                .GroupBy(static cell => cell.Row)
                .Select(static group => new DirectCellRow(group.Key, group.ToArray()))
                .ToArray();
            LegacyXlsDimensions dimensions = GetWorksheetDimensions(cells);
            long indexRecordPosition = stream.Position;
            WriteRecord(stream, 0x020b, BuildWorksheetIndexPayload(cells, dimensions.RowBlockCount));

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
            long defaultColumnWidthPosition = stream.Position;
            WriteRecord(stream, 0x0055, BuildUInt16Payload(8));
            WriteRecord(stream, 0x0200, BuildDimensionsPayload(dimensions));

            var dbCellPositions = new List<uint>(dimensions.RowBlockCount);
            int rowIndex = 0;
            for (int blockIndex = 0; blockIndex < dimensions.RowBlockCount; blockIndex++) {
                uint blockFirstRow = checked(dimensions.FirstRow + ((uint)blockIndex * DirectRowsPerDbCellBlock));
                uint blockRowAfterLast = Math.Min(
                    checked(blockFirstRow + DirectRowsPerDbCellBlock),
                    dimensions.RowAfterLast);
                int blockStart = rowIndex;
                while (rowIndex < rows.Length && rows[rowIndex].Row < blockRowAfterLast) {
                    rowIndex++;
                }
                int blockCount = rowIndex - blockStart;
                long firstRowPosition = stream.Position;
                for (int index = 0; index < blockCount; index++) {
                    WriteRecord(stream, 0x0208, BuildDirectRowPayload(rows[blockStart + index]));
                }

                var firstCellPositions = new long[blockCount];
                for (int index = 0; index < blockCount; index++) {
                    firstCellPositions[index] = stream.Position;
                    WriteCellRecords(stream, 0, rows[blockStart + index].Cells, sharedStrings);
                }

                long dbCellPosition = stream.Position;
                dbCellPositions.Add(checked((uint)dbCellPosition));
                WriteRecord(stream, 0x00d7, BuildDbCellPayload(firstRowPosition, firstCellPositions, dbCellPosition));
            }

            WriteRecord(stream, 0x023e, [
                0xb6, 0x06, 0x00, 0x00, 0x00, 0x00, 0x40, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00
            ]);
            WriteRecord(stream, 0x001d, BuildDefaultSelectionPayload());
            WriteRecord(stream, 0x000a, Array.Empty<byte>());

            PatchIndexRecord(stream, indexRecordPosition, defaultColumnWidthPosition, dbCellPositions);
        }

        private static byte[] BuildDirectRowPayload(DirectCellRow row) {
            byte[] payload = new byte[16];
            WriteUInt16(payload, 0, row.Row);
            WriteUInt16(payload, 2, row.Cells[0].Column);
            WriteUInt16(payload, 4, checked((ushort)(row.Cells[row.Cells.Count - 1].Column + 1)));
            WriteUInt16(payload, 6, 0x00ff);
            WriteUInt16(payload, 12, 0x0100);
            WriteUInt16(payload, 14, DirectDefaultCellStyleIndex);
            return payload;
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

        private sealed class DirectCellRow {
            internal DirectCellRow(ushort row, IReadOnlyList<LegacyXlsCell> cells) {
                Row = row;
                Cells = cells;
            }

            internal ushort Row { get; }

            internal IReadOnlyList<LegacyXlsCell> Cells { get; }
        }
    }
}
