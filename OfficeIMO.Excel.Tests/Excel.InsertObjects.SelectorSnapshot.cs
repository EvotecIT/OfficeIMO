using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void WriteData_EnumeratesBeforeSelectorsAndCapturesRowsAcrossChunkBoundaries() {
            const int count = 4_097;
            int enumerated = 0;
            int selectorCalls = 0;
            bool disposed = false;
            SelectorSnapshotRow? first = null;

            IEnumerable<SelectorSnapshotRow> Rows() {
                try {
                    for (int index = 0; index < count; index++) {
                        var row = new SelectorSnapshotRow { Id = index };
                        if (index == 0) first = row;
                        enumerated++;
                        yield return row;
                    }
                } finally {
                    disposed = true;
                }
            }

            object Select(SelectorSnapshotRow row, int column) {
                Assert.True(disposed);
                Assert.Equal(count, enumerated);
                Assert.Equal(row.Id * 2 + column, selectorCalls++);
                return row.Id + column;
            }

            using var document = ExcelDocument.Create();
            document.AddSheet("Data").WriteData(Rows(),
                ("Id", row => Select(row, 0)), ("Next", row => Select(row, 1)));
            Assert.Equal(count * 2, selectorCalls);
            first!.Id = -100;

            for (int save = 0; save < 2; save++) {
                byte[] bytes = document.ToBytes(ExcelFileFormat.Xlsb);
                Assert.Equal(count * 2, selectorCalls);
                using var reader = ExcelDocument.OpenReader(bytes);
                Assert.Equal("Id", reader.GetName(0));
                Assert.Equal("Next", reader.GetName(1));
                for (int index = 0; index < count; index++) {
                    Assert.True(reader.Read());
                    Assert.Equal(index, reader.GetInt32(0));
                    Assert.Equal(index + 1, reader.GetInt32(1));
                }
                Assert.False(reader.Read());
            }
        }

        [Theory]
        [InlineData(ExcelFileFormat.Xlsx)]
        [InlineData(ExcelFileFormat.Xlsb)]
        public void WriteData_OnSeveralSheetsKeepsTheSingleSheetValuesAndDateFormat(ExcelFileFormat format) {
            var start = new DateTime(2024, 1, 2, 13, 45, 0);
            var rows = Enumerable.Range(0, 300).ToArray();
            (string, Func<int, object?>)[] columns = [
                ("Name", index => "Name" + index), ("Id", index => index),
                ("CreatedOn", index => start.AddDays(index)), ("Amount", index => index * 1.25)
            ];

            string expectedFormat;
            using (var single = ExcelDocument.Create()) {
                single.AddSheet("Data").WriteData(rows, columns);
                using var loaded = ExcelDocument.Load(new MemoryStream(single.ToBytes(format)));
                expectedFormat = loaded.Sheets[0].Cell(2, 3).GetStyle().NumberFormatCode!;
            }

            using var document = ExcelDocument.Create();
            for (int sheet = 1; sheet <= 3; sheet++) document.AddSheet("Sheet" + sheet).WriteData(rows, columns);
            using var reloaded = ExcelDocument.Load(new MemoryStream(document.ToBytes(format)));
            Assert.Equal(3, reloaded.Sheets.Count);
            foreach (var sheet in reloaded.Sheets) {
                foreach (int row in new[] { 2, 301 }) {
                    int index = row - 2;
                    Assert.Equal(expectedFormat, sheet.Cell(row, 3).GetStyle().NumberFormatCode);
                    Assert.True(sheet.TryGetCellText(row, 1, out string name));
                    Assert.Equal("Name" + index, name);
                    Assert.True(sheet.TryGetCellText(row, 2, out string id));
                    Assert.Equal(index.ToString(System.Globalization.CultureInfo.InvariantCulture), id);
                    Assert.True(sheet.TryGetCellText(row, 3, out string serial));
                    Assert.Equal(start.AddDays(index).ToOADate(), double.Parse(serial, System.Globalization.CultureInfo.InvariantCulture), 9);
                }
            }
        }

        private sealed class SelectorSnapshotRow {
            internal int Id { get; set; }
        }
    }
}
