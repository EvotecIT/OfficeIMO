using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointTableCellFormatting {
        [Fact]
        public void CanSetTableCellBordersAndPadding() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTable table = slide.AddTable(2, 2);
                PowerPointTableCell cell = table.GetCell(0, 0);
                cell.Text = "Header";
                cell.SetBorders(TableCellBorders.All, "FF0000", 1.5);
                cell.PaddingLeftPoints = 4;
                cell.PaddingRightPoints = 4;
                cell.PaddingTopPoints = 3;
                cell.PaddingBottomPoints = 3;
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointTableCell cell = presentation.Slides.Single().Tables.First().GetCell(0, 0);
                Assert.Equal("FF0000", cell.BorderColor);
                Assert.Equal(4, cell.PaddingLeftPoints ?? 0, 2);
                Assert.Equal(4, cell.PaddingRightPoints ?? 0, 2);
                Assert.Equal(3, cell.PaddingTopPoints ?? 0, 2);
                Assert.Equal(3, cell.PaddingBottomPoints ?? 0, 2);
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                A.TableCell cell = document.PresentationPart!.SlideParts.First().Slide
                    .Descendants<A.TableCell>()
                    .First();
                int? width = cell.TableCellProperties?.LeftBorderLineProperties?.Width?.Value;
                Assert.Equal((int)Math.Round(1.5 * 12700), width);
            }

            File.Delete(filePath);
        }
    }
}
