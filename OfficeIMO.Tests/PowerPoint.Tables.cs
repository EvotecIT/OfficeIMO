        public void CanManipulateTableCellsAndPreserveStyle() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();

            File.Delete(filePath);
        }

        [Theory]
        [InlineData(-1)]
        [InlineData(2)]
        public void GetCellThrowsForInvalidRow(int invalidRow) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointTable table = presentation.AddSlide().AddTable(2, 2);

                ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() => table.GetCell(invalidRow, 0));

                Assert.Equal("row", exception.ParamName);
                Assert.Equal(invalidRow, Assert.IsType<int>(exception.ActualValue));
                Assert.Contains("Row index", exception.Message);
                Assert.Contains("Valid range", exception.Message);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Theory]
        [InlineData(-1)]
        [InlineData(2)]
        public void GetCellThrowsForInvalidColumn(int invalidColumn) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointTable table = presentation.AddSlide().AddTable(2, 2);

                ArgumentOutOfRangeException exception = Assert.Throws<ArgumentOutOfRangeException>(() => table.GetCell(0, invalidColumn));

                Assert.Equal("column", exception.ParamName);
                Assert.Equal(invalidColumn, Assert.IsType<int>(exception.ActualValue));
                Assert.Contains("Column index", exception.Message);
                Assert.Contains("Valid range", exception.Message);
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
