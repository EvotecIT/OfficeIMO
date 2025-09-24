        public void CanManipulateTableCellsAndPreserveStyle() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();

            File.Delete(filePath);
        }

        [Fact]
        public void AddRowThrowsWhenIndexIsOutOfRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(2, 2);

                    ArgumentOutOfRangeException negativeIndex = Assert.Throws<ArgumentOutOfRangeException>(() => table.AddRow(-1));
                    Assert.Equal("index", negativeIndex.ParamName);
                    Assert.Equal(-1, (int)negativeIndex.ActualValue!);
                    Assert.Contains("Row index must be between 0", negativeIndex.Message, StringComparison.Ordinal);

                    int tooLargeIndex = table.Rows + 1;
                    ArgumentOutOfRangeException tooLarge = Assert.Throws<ArgumentOutOfRangeException>(() => table.AddRow(tooLargeIndex));
                    Assert.Equal("index", tooLarge.ParamName);
                    Assert.Equal(tooLargeIndex, (int)tooLarge.ActualValue!);
                    Assert.Contains("Row index must be between 0", tooLarge.Message, StringComparison.Ordinal);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void AddColumnThrowsWhenIndexIsOutOfRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(2, 2);

                    ArgumentOutOfRangeException negativeIndex = Assert.Throws<ArgumentOutOfRangeException>(() => table.AddColumn(-1));
                    Assert.Equal("index", negativeIndex.ParamName);
                    Assert.Equal(-1, (int)negativeIndex.ActualValue!);
                    Assert.Contains("Column index must be between 0", negativeIndex.Message, StringComparison.Ordinal);

                    int tooLargeIndex = table.Columns + 1;
                    ArgumentOutOfRangeException tooLarge = Assert.Throws<ArgumentOutOfRangeException>(() => table.AddColumn(tooLargeIndex));
                    Assert.Equal("index", tooLarge.ParamName);
                    Assert.Equal(tooLargeIndex, (int)tooLarge.ActualValue!);
                    Assert.Contains("Column index must be between 0", tooLarge.Message, StringComparison.Ordinal);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void RemoveRowThrowsWhenIndexIsOutOfRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(2, 2);

                    ArgumentOutOfRangeException negativeIndex = Assert.Throws<ArgumentOutOfRangeException>(() => table.RemoveRow(-1));
                    Assert.Equal("index", negativeIndex.ParamName);
                    Assert.Equal(-1, (int)negativeIndex.ActualValue!);
                    Assert.Contains("Row index must be between 0", negativeIndex.Message, StringComparison.Ordinal);

                    int tooLargeIndex = table.Rows;
                    ArgumentOutOfRangeException tooLarge = Assert.Throws<ArgumentOutOfRangeException>(() => table.RemoveRow(tooLargeIndex));
                    Assert.Equal("index", tooLarge.ParamName);
                    Assert.Equal(tooLargeIndex, (int)tooLarge.ActualValue!);
                    Assert.Contains("Row index must be between 0", tooLarge.Message, StringComparison.Ordinal);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void RemoveColumnThrowsWhenIndexIsOutOfRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(2, 2);

                    ArgumentOutOfRangeException negativeIndex = Assert.Throws<ArgumentOutOfRangeException>(() => table.RemoveColumn(-1));
                    Assert.Equal("index", negativeIndex.ParamName);
                    Assert.Equal(-1, (int)negativeIndex.ActualValue!);
                    Assert.Contains("Column index must be between 0", negativeIndex.Message, StringComparison.Ordinal);

                    int tooLargeIndex = table.Columns;
                    ArgumentOutOfRangeException tooLarge = Assert.Throws<ArgumentOutOfRangeException>(() => table.RemoveColumn(tooLargeIndex));
                    Assert.Equal("index", tooLarge.ParamName);
                    Assert.Equal(tooLargeIndex, (int)tooLarge.ActualValue!);
                    Assert.Contains("Column index must be between 0", tooLarge.Message, StringComparison.Ordinal);
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
