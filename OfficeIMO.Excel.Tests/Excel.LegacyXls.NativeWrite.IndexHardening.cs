using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Read;
using OfficeIMO.Excel.LegacyXls.Write;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(0, 1)]
        [InlineData(1, 2)]
        public void LegacyXls_IndexedDiscoveryFallsBackWhenDbCellCoverageIsNotContiguous(
            int pointerIndex,
            int replacementIndex) {
            byte[] xls = CreateDenseIndexedWorkbook(("Data", 1));
            byte[] workbookStream = ReadCompoundStream(xls, "Workbook");
            IndexedWorksheetLayout layout = Assert.Single(ReadIndexedWorksheetLayouts(workbookStream));
            Assert.True(layout.DbCellOffsets.Count > replacementIndex);

            WriteUInt32(
                workbookStream,
                layout.IndexPayloadOffset + 16 + (pointerIndex * sizeof(uint)),
                checked((uint)layout.DbCellOffsets[replacementIndex]));
            byte[] mutated = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            AssertIndexedFallbackReadsDenseValues(mutated, "Data", firstValue: 1);
        }

        [Fact]
        public void LegacyXls_IndexedDiscoveryRejectsDbCellPointersFromAnotherWorksheet() {
            byte[] xls = CreateDenseIndexedWorkbook(("First", 1), ("Second", 1001));
            AssertUsesIndexedDiscovery(xls, "First");
            byte[] workbookStream = ReadCompoundStream(xls, "Workbook");
            IReadOnlyList<IndexedWorksheetLayout> layouts = ReadIndexedWorksheetLayouts(workbookStream);
            Assert.Equal(2, layouts.Count);
            Assert.NotEmpty(layouts[0].DbCellOffsets);
            Assert.NotEmpty(layouts[1].DbCellOffsets);

            WriteUInt32(
                workbookStream,
                layouts[0].IndexPayloadOffset + 16,
                checked((uint)layouts[1].DbCellOffsets[0]));
            byte[] mutated = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            AssertIndexedFallbackReadsDenseValues(mutated, "First", firstValue: 1);
        }

        [Fact]
        public void LegacyXls_IndexedReaderStillObservesConfiguredCancellation() {
            byte[] xls = CreateDenseIndexedWorkbook(("Data", 1));
            using var cancellation = new CancellationTokenSource();
            var options = new ExcelReaderOptions { CancellationToken = cancellation.Token };
            using LegacyXlsTabularWorkbook workbook = LegacyXlsTabularWorkbook.Open(xls, options);
            using var reader = Assert.IsType<LegacyXlsTabularDataReader>(
                workbook.OpenTable("Data", hasHeaderRow: true, options, cancellation.Token));

            cancellation.Cancel();

            Assert.Throws<OperationCanceledException>(() => reader.Read());
        }

        [Theory]
        [InlineData(0)]
        [InlineData(1)]
        public void LegacyXls_IndexedReaderRejectsCellRecordsCrossingRows(int targetRow) {
            byte[] xls = CreateDenseIndexedWorkbook(("Data", 1));
            AssertUsesIndexedDiscovery(xls, "Data");
            byte[] workbookStream = ReadCompoundStream(xls, "Workbook");
            bool inWorksheet = false;
            int cellOffset = -1;
            int cellLength = 0;
            for (int offset = 0; offset + 4 <= workbookStream.Length;) {
                ushort type = ReadUInt16(workbookStream, offset);
                int length = ReadUInt16(workbookStream, offset + 2);
                int payloadOffset = offset + 4;
                Assert.True(length <= workbookStream.Length - payloadOffset);
                if (type == 0x0809 && length >= 4) {
                    inWorksheet = ReadUInt16(workbookStream, payloadOffset + 2) == 0x0010;
                }
                if (inWorksheet
                    && (type == 0x00fd || type == 0x027e || type == 0x0203)
                    && length >= 10
                    && ReadUInt16(workbookStream, payloadOffset) == targetRow) {
                    cellOffset = offset;
                    cellLength = length;
                    break;
                }
                offset = payloadOffset + length;
            }
            Assert.True(cellOffset >= 0, "The fixture cell record was not found.");
            WriteUInt16(workbookStream, cellOffset + 2, checked((ushort)(cellLength + 1)));
            byte[] mutated = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            Assert.Throws<InvalidDataException>(() => {
                using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(mutated);
                while (reader.Read()) {
                    _ = reader.GetInt32(0);
                }
            });
        }

        [Theory]
        [InlineData(1)]
        [InlineData(4)]
        public void LegacyXls_IndexedHeaderFallsBackFromMisalignedCellPointers(int headerGap) {
            byte[] xls = CreateWorkbookWithMisindexedHeader(headerGap);
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(xls);
            Assert.Equal("Value", reader.GetName(0));
            int count = 0;
            while (reader.Read()) {
                Assert.Equal(count + 1, reader.GetInt32(0));
                count++;
            }
            Assert.Equal(255, count);
        }

        private static byte[] CreateWorkbookWithMisindexedHeader(int headerGap) {
            byte[] workbook = ReadCompoundStream(CreateDenseIndexedWorkbook(("Data", 1)), "Workbook");
            IndexedWorksheetLayout layout = Assert.Single(ReadIndexedWorksheetLayouts(workbook));
            int dbCellOffset = layout.DbCellOffsets[0];
            int dbCellPayload = dbCellOffset + 4;
            int firstRowOffset = checked(dbCellOffset - (int)ReadUInt32(workbook, dbCellPayload));
            int firstRowEnd = firstRowOffset + 4 + ReadUInt16(workbook, firstRowOffset + 2);
            int firstCell = firstRowEnd + ReadUInt16(workbook, dbCellPayload + 4);
            int thirdCell = firstCell
                + ReadUInt16(workbook, dbCellPayload + 6)
                + ReadUInt16(workbook, dbCellPayload + 8);
            Assert.Equal(0x00fd, ReadUInt16(workbook, firstCell));

            const int insertedBytes = 132;
            byte[] changed = new byte[workbook.Length + insertedBytes];
            Array.Copy(workbook, 0, changed, 0, firstCell);
            Array.Copy(workbook, firstCell, changed, firstCell + insertedBytes, workbook.Length - firstCell);
            WriteUInt16(changed, firstCell, 0x7777);
            WriteUInt16(changed, firstCell + 2, 128);

            // The actual cells stay intact. Only the INDEX points into this opaque record.
            int fakeHeader = firstCell + 12;
            int fakeData = fakeHeader + 14 + headerGap;
            WriteUInt16(changed, fakeHeader, 0x00fd);
            WriteUInt16(changed, fakeHeader + 2, 10);
            WriteUInt32(changed, fakeHeader + 10, ReadUInt32(workbook, firstCell + 10));
            WriteUInt16(changed, fakeData, 0x027e);
            WriteUInt16(changed, fakeData + 2, 64);
            WriteUInt16(changed, fakeData + 4, 1);
            WriteUInt32(changed, fakeData + 10, 6);
            for (int index = 0; index < headerGap; index++) changed[fakeHeader + 14 + index] = 0xff;

            for (int index = 0; index < layout.DbCellOffsets.Count; index++) {
                WriteUInt32(changed, layout.IndexPayloadOffset + 16 + index * sizeof(uint),
                    checked((uint)(layout.DbCellOffsets[index] + insertedBytes)));
            }
            int newDbCellPayload = dbCellPayload + insertedBytes;
            WriteUInt32(changed, newDbCellPayload, checked(ReadUInt32(workbook, dbCellPayload) + insertedBytes));
            WriteUInt16(changed, newDbCellPayload + 4, checked((ushort)(fakeHeader - firstRowEnd)));
            WriteUInt16(changed, newDbCellPayload + 6, checked((ushort)(fakeData - fakeHeader)));
            WriteUInt16(changed, newDbCellPayload + 8, checked((ushort)(thirdCell + insertedBytes - fakeData)));
            return LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(changed);
        }

        private static byte[] CreateDenseIndexedWorkbook(
            params (string Name, int FirstValue)[] sheets) {
            using ExcelDocument document = ExcelDocument.Create();
            foreach ((string name, int firstValue) in sheets) {
                ExcelSheet sheet = document.AddSheet(name);
                sheet.Cell(1, 1).SetValue("Value");
                for (int row = 2; row <= 256; row++) {
                    sheet.Cell(row, 1).SetValue(firstValue + row - 2);
                }
            }

            byte[] workbook = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xls });
            if (sheets.Length == 1) {
                Assert.Equal(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
            }
            return workbook;
        }

        private static void AssertUsesIndexedDiscovery(byte[] workbookBytes, string sheetName) {
            var options = new ExcelReaderOptions();
            using LegacyXlsTabularWorkbook workbook = LegacyXlsTabularWorkbook.Open(workbookBytes, options);
            using var reader = Assert.IsType<LegacyXlsTabularDataReader>(
                workbook.OpenTable(sheetName, hasHeaderRow: true, options));
            Assert.True(reader.UsedIndexedDiscovery);
        }

        private static void AssertIndexedFallbackReadsDenseValues(
            byte[] workbookBytes,
            string sheetName,
            int firstValue) {
            var options = new ExcelReaderOptions();
            using LegacyXlsTabularWorkbook workbook = LegacyXlsTabularWorkbook.Open(workbookBytes, options);
            using var reader = Assert.IsType<LegacyXlsTabularDataReader>(
                workbook.OpenTable(sheetName, hasHeaderRow: true, options));
            Assert.False(reader.UsedIndexedDiscovery);

            int rowCount = 0;
            while (reader.Read()) {
                Assert.Equal(firstValue + rowCount, reader.GetInt32(0));
                rowCount++;
            }
            Assert.Equal(255, rowCount);
        }

        private static IReadOnlyList<IndexedWorksheetLayout> ReadIndexedWorksheetLayouts(
            byte[] workbookStream) {
            var layouts = new List<IndexedWorksheetLayout>();
            IndexedWorksheetLayout? current = null;
            int offset = 0;
            while (offset + 4 <= workbookStream.Length) {
                ushort type = ReadUInt16(workbookStream, offset);
                ushort length = ReadUInt16(workbookStream, offset + 2);
                int payloadOffset = offset + 4;
                if (payloadOffset + length > workbookStream.Length) break;

                if (type == 0x0809
                    && length >= 4
                    && ReadUInt16(workbookStream, payloadOffset + 2) == 0x0010) {
                    current = new IndexedWorksheetLayout();
                    layouts.Add(current);
                } else if (current != null) {
                    if (type == 0x020b) {
                        current.IndexPayloadOffset = payloadOffset;
                    } else if (type == 0x00d7) {
                        current.DbCellOffsets.Add(offset);
                    } else if (type == 0x000a) {
                        current = null;
                    }
                }

                offset = payloadOffset + length;
            }

            foreach (IndexedWorksheetLayout layout in layouts) {
                Assert.True(layout.IndexPayloadOffset >= 0, "The worksheet INDEX record was not found.");
                Assert.NotEmpty(layout.DbCellOffsets);
            }
            return layouts;
        }

        private sealed class IndexedWorksheetLayout {
            internal int IndexPayloadOffset { get; set; } = -1;
            internal List<int> DbCellOffsets { get; } = new();
        }
    }
}
