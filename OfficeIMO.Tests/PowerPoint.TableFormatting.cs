using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointTableFormatting {
        private const int EmusPerPoint = 12700;

        [Fact]
        public void CanSizeTableAndFormatHeader() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            const double columnWidthPoints = 120;
            const double rowHeightPoints = 30;

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(2, 2, 914400L, 1828800L, 5000000L, 2000000L);
                    table.SetColumnWidthPoints(0, columnWidthPoints);
                    table.SetRowHeightPoints(0, rowHeightPoints);

                    PowerPointTableCell header = table.GetCell(0, 0);
                    header.Text = "Header";
                    header.Bold = true;
                    header.FontSize = 14;
                    header.Color = "FFFFFF";

                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    A.Table table = document.PresentationPart!.SlideParts.First()
                        .Slide.Descendants<A.Table>().First();

                    long expectedColumnWidth = (long)Math.Round(columnWidthPoints * EmusPerPoint);
                    long expectedRowHeight = (long)Math.Round(rowHeightPoints * EmusPerPoint);

                    long actualColumnWidth = table.TableGrid!.Elements<A.GridColumn>().First().Width!.Value;
                    Assert.Equal(expectedColumnWidth, actualColumnWidth);

                    long actualRowHeight = table.Elements<A.TableRow>().First().Height!.Value;
                    Assert.Equal(expectedRowHeight, actualRowHeight);

                    A.RunProperties props = table.Elements<A.TableRow>().First()
                        .Elements<A.TableCell>().First()
                        .TextBody!.Elements<A.Paragraph>().First()
                        .Elements<A.Run>().First()
                        .RunProperties!;

                    Assert.True(props.Bold!.Value);
                    Assert.Equal(1400, props.FontSize!.Value);
                    Assert.Equal("FFFFFF", props.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
