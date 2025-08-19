using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointTables {
        [Fact]
        public void CanManipulateTableCellsAndPreserveStyle() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PPTable table = slide.AddTable(2, 2);
                PPTableCell cell = table.GetCell(0, 0);
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
                PPTable table = presentation.Slides[0].Tables.First();
                Assert.Equal(2, table.Rows);
                Assert.Equal(2, table.Columns);
                PPTableCell cell = table.GetCell(0, 0);
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
