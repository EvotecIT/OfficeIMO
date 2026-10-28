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
        public void Hidden_borders_can_be_restored_without_conflicting_fills_or_changing_other_edges() {
            using var presentation = PowerPointDocument.Create();
            var cell = presentation.AddSlide().AddTable(1,1).GetCell(0, 0);
            cell.Text = "Preserved"; cell.FillColor = "ABCDEF";
            cell.SetBorders(PowerPointTableCellBorders.All, "123456", 1);
            cell.HideBorders(PowerPointTableCellBorders.Right | PowerPointTableCellBorders.Bottom);
            cell.SetBorders(PowerPointTableCellBorders.Bottom, "FF0000", 3);
            Assert.Throws<ArgumentOutOfRangeException>(() => cell.SetBorders(PowerPointTableCellBorders.Top, "00FF00", 1584.25));
            Assert.Throws<ArgumentOutOfRangeException>(() => cell.HideBorders((PowerPointTableCellBorders)128));
            using var saved = new MemoryStream(); presentation.Save(saved); saved.Position = 0;
            using (var reopened = PowerPointDocument.Load(saved)) {
                Assert.Empty(reopened.ValidateDocument());
                var output = reopened.Slides.Single().Tables.Single().GetCell(0, 0);
                Assert.Equal("Preserved", output.Text); Assert.Equal("ABCDEF", output.FillColor);
            }
            saved.Position = 0; using var package = PresentationDocument.Open(saved, false);
            var properties = package.PresentationPart!.SlideParts.Single().Slide.Descendants<A.TableCell>().Single().TableCellProperties!;
            Assert.NotNull(properties.RightBorderLineProperties!.GetFirstChild<A.NoFill>());
            Assert.Null(properties.BottomBorderLineProperties!.GetFirstChild<A.NoFill>());
            Assert.Equal("FF0000", properties.BottomBorderLineProperties.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Equal(38100, properties.BottomBorderLineProperties.Width!.Value);
            Assert.Equal("123456", properties.TopBorderLineProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val!.Value);
            Assert.Equal(12700, properties.TopBorderLineProperties.Width!.Value);
        }

        [Fact]
        public void CanSetTableCellBordersAndPadding() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTable table = slide.AddTable(2,2);
                PowerPointTableCell cell = table.GetCell(0, 0);
                cell.Text = "Header";
                cell.SetBorders(PowerPointTableCellBorders.All, "FF0000", 1.5);
                cell.PaddingLeftPoints = 4;
                cell.PaddingRightPoints = 4;
                cell.PaddingTopPoints = 3;
                cell.PaddingBottomPoints = 3;
                presentation.Save();
            }

            using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
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
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(1,1);
                    PowerPointTableCell cell = table.GetCell(0, 0);
                    IReadOnlyList<PowerPointParagraph> paragraphs = cell.SetParagraphs(new[] { "First", "Second" }, paragraph => {
                        paragraph.SetLineSpacingPoints(18);
                    });
                    paragraphs[0].AddRun(" bold", run => run.Bold = true);
                    presentation.Save();
                }

                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
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
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(1,1);
                    PowerPointTableCell cell = table.GetCell(0, 0);
                    cell.SetBullets(new[] { "First bullet", "Second bullet" }, configure: paragraph => {
                        paragraph.SetBulletSizePoints(11);
                    });
                    cell.AddNumberedList(new[] { "First number", "Second number" }, PowerPointNumberingScheme.AlphaLowerCharacterPeriod, startAt: 2);
                    presentation.Save();
                }

                using (PowerPointDocument presentation = PowerPointDocument.Load(filePath)) {
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
