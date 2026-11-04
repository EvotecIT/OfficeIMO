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

        private sealed class SelectorSnapshotRow {
            internal int Id { get; set; }
        }
    }
}
