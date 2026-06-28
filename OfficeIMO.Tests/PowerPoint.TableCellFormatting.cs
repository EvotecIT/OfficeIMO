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

        [Fact]
        public void CanSetTableCellParagraphsThroughFriendlyApi() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(1, 1);
                    PowerPointTableCell cell = table.GetCell(0, 0);
                    IReadOnlyList<PowerPointParagraph> paragraphs = cell.SetParagraphs(new[] { "First", "Second" }, paragraph => {
                        paragraph.SetLineSpacingPoints(18);
                    });
                    paragraphs[0].AddRun(" bold", run => run.Bold = true);
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointTableCell cell = presentation.Slides.Single().Tables.First().GetCell(0, 0);
                    Assert.Equal(new[] { "First bold", "Second" }, cell.Paragraphs.Select(paragraph => paragraph.Text).ToArray());
                    Assert.Equal(18D, cell.Paragraphs[0].LineSpacingPoints);
                    Assert.True(cell.Paragraphs[0].Runs.Last().Bold);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanSetTableCellListsThroughFriendlyApi() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(1, 1);
                    PowerPointTableCell cell = table.GetCell(0, 0);
                    cell.SetBullets(new[] { "First bullet", "Second bullet" }, configure: paragraph => {
                        paragraph.SetBulletSizePoints(11);
                    });
                    cell.AddNumberedList(new[] { "First number", "Second number" }, A.TextAutoNumberSchemeValues.AlphaLowerCharacterPeriod, startAt: 2);
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointTableCell cell = presentation.Slides.Single().Tables.First().GetCell(0, 0);
                    PowerPointParagraph[] paragraphs = cell.Paragraphs.ToArray();

                    Assert.Equal(new[] { "First bullet", "Second bullet", "First number", "Second number" }, paragraphs.Select(paragraph => paragraph.Text).ToArray());
                    Assert.Equal("\u2022", paragraphs[0].BulletCharacter);
                    Assert.Equal(11, paragraphs[0].BulletSizePoints);
                    Assert.Equal(18D, paragraphs[0].LeftMarginPoints);
                    Assert.Equal(-18D, paragraphs[0].IndentPoints);
                    Assert.True(paragraphs[2].IsNumbered);
                    Assert.Equal(2, paragraphs[2].NumberingStartAt);
                    Assert.True(paragraphs[3].IsNumbered);
                    Assert.Null(paragraphs[3].NumberingStartAt);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
