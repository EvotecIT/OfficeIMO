using OfficeIMO.Word;
using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SetCellText_UpdatesGeneratedTableAndRoundTrips() {
            byte[] bytes;
            using (WordDocument document = WordDocument.Create()) {
                WordTable table = document.AddTable(2, 2);

                Assert.Same(table, table.SetCellText(0, 0, "Record", bold: true));
                table.SetCellText(0, 1, "Owner", bold: true)
                    .SetCellText(1, 0, "001")
                    .SetCellText(1, 1, "Operations");

                Assert.Equal("Record", table.Rows[0].Cells[0].Paragraphs[0].Text);
                Assert.True(table.Rows[0].Cells[0].Paragraphs[0].Bold);
                Assert.Equal("Operations", table.Rows[1].Cells[1].Paragraphs[0].Text);
                Assert.False(table.Rows[1].Cells[1].Paragraphs[0].Bold);

                bytes = document.ToBytes();
            }

            using WordDocument reloaded = WordDocument.Load(new MemoryStream(bytes, writable: false));
            WordTable savedTable = reloaded.Tables[0];
            Assert.Equal("Record", savedTable.Rows[0].Cells[0].Paragraphs[0].Text);
            Assert.True(savedTable.Rows[0].Cells[0].Paragraphs[0].Bold);
            Assert.Equal("Operations", savedTable.Rows[1].Cells[1].Paragraphs[0].Text);
        }

        [Fact]
        public void SetCellText_TracksRowChangesAndChecksBounds() {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(2, 2);

            table.AddRow(2);
            table.SetCellText(2, 1, "Added");
            Assert.Equal("Added", table.Rows[2].Cells[1].Paragraphs[0].Text);

            table.Rows[0].Remove();
            table.SetCellText(0, 0, "Now first");
            Assert.Equal("Now first", table.Rows[0].Cells[0].Paragraphs[0].Text);

            Assert.Throws<ArgumentOutOfRangeException>(() => table.SetCellText(-1, 0, "x"));
            Assert.Throws<ArgumentOutOfRangeException>(() => table.SetCellText(0, -1, "x"));
            Assert.Throws<ArgumentOutOfRangeException>(() => table.SetCellText(2, 0, "x"));
            Assert.Throws<ArgumentOutOfRangeException>(() => table.SetCellText(0, 2, "x"));
        }
    }
}
