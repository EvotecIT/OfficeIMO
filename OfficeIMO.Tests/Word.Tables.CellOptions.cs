using System;
using System.IO;
using OfficeIMO.Word;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests cell-level options such as margins, wrapping, and fit text.
    /// </summary>
    public partial class Word {
        /// <summary>
        /// Verifies that individual table cell margins can be set and persisted.
        /// </summary>
        [Fact]
        public void Test_TableCellMargins() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithCellMargins.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(2, 2);

                // Set a default right margin for the table
                table.StyleDetails!.MarginDefaultRightWidth = 100;

                // Override margins for a specific cell
                var cell = table.Rows[0].Cells[1];
                cell.MarginRightWidth = 200;
                cell.MarginTopCentimeters = 0.3;

                // Verify values before saving
                Assert.True(cell.MarginRightWidth == 200);
                Assert.True(Math.Abs(cell.MarginTopCentimeters.GetValueOrDefault() - 0.3) < 0.01);
                Assert.True(table.Rows[0].Cells[0].MarginRightWidth == null);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithCellMargins.docx"))) {
                var cell = document.Tables[0].Rows[0].Cells[1];
                Assert.True(cell.MarginRightWidth == 200);
                Assert.True(Math.Abs(cell.MarginTopCentimeters.GetValueOrDefault() - 0.3) < 0.01);
                Assert.True(document.Tables[0].Rows[0].Cells[0].MarginRightWidth == null);
                document.Save();
            }
        }

        /// <summary>
        /// Verifies WrapText and FitText persistence on table cells.
        /// </summary>
        [Fact]
        public void Test_TableCellWrapAndFitText() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithCellOptions.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(1, 2);

                var cell1 = table.Rows[0].Cells[0];
                cell1.WrapText = false;

                var cell2 = table.Rows[0].Cells[1];
                cell2.FitText = true;

                Assert.False(cell1.WrapText);
                Assert.True(cell2.FitText);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var cell1 = document.Tables[0].Rows[0].Cells[0];
                var cell2 = document.Tables[0].Rows[0].Cells[1];

                Assert.False(cell1.WrapText);
                Assert.True(cell2.FitText);

                document.Save();
            }
        }

        /// <summary>
        /// Ensures that the paragraph parent navigation returns the owning table cell instance.
        /// </summary>
        [Fact]
        public void Test_ParagraphParentWithinCell() {
            string filePath = Path.Combine(_directoryWithFiles, "ParagraphParentInCell.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(1, 1);
                var cell = table.Rows[0].Cells[0];
                var paragraph = cell.Paragraphs.First();
                paragraph.Text = "Cell paragraph";

                var parentCell = Assert.IsType<WordTableCell>(paragraph.Parent);
                Assert.Equal("Cell paragraph", parentCell.Paragraphs.First().Text);
                Assert.Equal(table.RowsCount, parentCell.ParentTable.RowsCount);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var table = document.Tables[0];
                var cell = table.Rows[0].Cells[0];
                var paragraph = cell.Paragraphs.First();

                var parentCell = Assert.IsType<WordTableCell>(paragraph.Parent);
                Assert.Equal(paragraph.Text, parentCell.Paragraphs.First().Text);
                Assert.Equal(table.RowsCount, parentCell.ParentTable.RowsCount);

                document.Save();
            }
        }

        /// <summary>
        /// Ensures that evaluating paragraph parent does not inject missing table cell properties.
        /// </summary>
        [Fact]
        public void Test_ParagraphParentDoesNotCreateTableCellProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "ParagraphParentNoCellProps.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(1, 1);
                var cell = table.Rows[0].Cells[0];
                cell.RemoveTableCellProperties();

                Assert.Null(cell._tableCell.TableCellProperties);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var table = document.Tables[0];
                var cell = table.Rows[0].Cells[0];

                cell.RemoveTableCellProperties();
                Assert.Null(cell._tableCell.TableCellProperties);

                var paragraph = cell.Paragraphs.First();
                var parentCell = Assert.IsType<WordTableCell>(paragraph.Parent);

                Assert.Same(cell._tableCell, parentCell._tableCell);
                Assert.Null(cell._tableCell.TableCellProperties);
                Assert.Null(parentCell._tableCell.TableCellProperties);

                document.Save();
            }
        }
    }
}
