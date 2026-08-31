using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls.Read;
using OfficeIMO.Excel.LegacyXls.Write;
using System.Threading;
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
            var options = new ExcelReadOptions { CancellationToken = cancellation.Token };
            using LegacyXlsTabularWorkbook workbook = LegacyXlsTabularWorkbook.Open(xls, options);
            using var reader = Assert.IsType<LegacyXlsTabularDataReader>(
                workbook.OpenTable("Data", hasHeaderRow: true, options, cancellation.Token));

            cancellation.Cancel();

            Assert.Throws<OperationCanceledException>(() => reader.Read());
        }

        private static byte[] CreateDenseIndexedWorkbook(
            params (string Name, int FirstValue)[] sheets) {
            using ExcelDocument document = ExcelDocument.Create();
            foreach ((string name, int firstValue) in sheets) {
                ExcelSheet sheet = document.AddWorksheet(name);
                sheet.CellValue(1, 1, "Value");
                for (int row = 2; row <= 256; row++) {
                    sheet.CellValue(row, 1, firstValue + row - 2);
                }
            }

            byte[] workbook = document.ToBytes(ExcelFileFormat.Xls);
            if (sheets.Length == 1) {
                Assert.Equal(ExcelSavePackageWriter.NativeBinaryDirectPackage, document.LastSaveDiagnostics.Writer);
            }
            return workbook;
        }

        private static void AssertUsesIndexedDiscovery(byte[] workbookBytes, string sheetName) {
            var options = new ExcelReadOptions();
            using LegacyXlsTabularWorkbook workbook = LegacyXlsTabularWorkbook.Open(workbookBytes, options);
            using var reader = Assert.IsType<LegacyXlsTabularDataReader>(
                workbook.OpenTable(sheetName, hasHeaderRow: true, options));
            Assert.True(reader.UsedIndexedDiscovery);
        }

        private static void AssertIndexedFallbackReadsDenseValues(
            byte[] workbookBytes,
            string sheetName,
            int firstValue) {
            var options = new ExcelReadOptions();
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
