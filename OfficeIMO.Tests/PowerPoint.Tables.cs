        public void CanManipulateTableCellsAndPreserveStyle() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTable table = slide.AddTable(2, 2);
            File.Delete(filePath);
        }

        [Theory]
        [InlineData(-1)]
        [InlineData(2)]
        public void GetCellInvalidRowThrowsArgumentOutOfRangeException(int row) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(2, 2);

                    ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() => table.GetCell(row, 0));
                    Assert.StartsWith($"Row index {row} is out of range.", exception.Message);
                    Assert.EndsWith("(Parameter 'row')", exception.Message);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Theory]
        [InlineData(-1)]
        [InlineData(2)]
        public void GetCellInvalidColumnThrowsArgumentOutOfRangeException(int column) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(2, 2);

                    ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() => table.GetCell(0, column));
                    Assert.StartsWith($"Column index {column} is out of range.", exception.Message);
                    Assert.EndsWith("(Parameter 'column')", exception.Message);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Theory]
        [InlineData(-1)]
        [InlineData(2)]
        public void RemoveRowInvalidIndexThrowsArgumentOutOfRangeException(int rowIndex) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(2, 2);

                    ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() => table.RemoveRow(rowIndex));
                    Assert.StartsWith($"Row index {rowIndex} is out of range.", exception.Message);
                    Assert.EndsWith("(Parameter 'index')", exception.Message);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Theory]
        [InlineData(-1)]
        [InlineData(2)]
        public void RemoveColumnInvalidIndexThrowsArgumentOutOfRangeException(int columnIndex) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(2, 2);

                    ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() => table.RemoveColumn(columnIndex));
                    Assert.StartsWith($"Column index {columnIndex} is out of range.", exception.Message);
                    Assert.EndsWith("(Parameter 'index')", exception.Message);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointTables {
        [Fact]
        public void CanManipulateTableCellsAndPreserveStyle() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTable table = slide.AddTable(2, 2);
                PowerPointTableCell cell = table.GetCell(0, 0);
                cell.Text = "Test";
                cell.FillColor = "FF0000";
                cell.Merge = (1, 2);
                table.AddRow();
                table.AddColumn();
                table.RemoveRow(2);
                table.RemoveColumn(2);
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointTable table = presentation.Slides[0].Tables.First();
                Assert.Equal(2, table.Rows);
                Assert.Equal(2, table.Columns);
                PowerPointTableCell cell = table.GetCell(0, 0);
                Assert.Equal("Test", cell.Text);
                Assert.Equal((1, 2), cell.Merge);
            }

            using (PresentationDocument doc = PresentationDocument.Open(filePath, false)) {
                A.Table table = doc.PresentationPart!.SlideParts.First().Slide.Descendants<A.Table>().First();
                string? styleId = table.TableProperties?.GetFirstChild<A.TableStyleId>()?.Text;
                Assert.Equal("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", styleId);
            }

            File.Delete(filePath);
        }
    }
}
