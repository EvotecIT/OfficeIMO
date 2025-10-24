using System;
using System.Linq;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests {
    public class ParagraphParentMutationTests {
        [Fact]
        public void ParagraphParent_DoesNotMutateDom() {
            using var doc = WordDocument.Create();

            // Arrange: create a table with a paragraph in its single cell
            var table = doc.AddTable(1, 1, WordTableStyle.TableGrid);
            var cell = table.Rows[0].Cells[0];

            // Remove tcPr so we can detect accidental mutations caused by a read
            cell.RemoveTableCellProperties();
            Assert.Null(cell._tableCell.TableCellProperties);

            var para = cell.Paragraphs.First();
            string before = cell._tableCell.OuterXml;

            // Act: resolve parent
            var parent = (para.Parent as WordTableCell);

            // Assert: parent resolved, but DOM unchanged
            Assert.NotNull(parent);
            Assert.Null(cell._tableCell.TableCellProperties);
            string after = cell._tableCell.OuterXml;
            Assert.Equal(before, after);
        }
    }
}

