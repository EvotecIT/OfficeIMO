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


        [Fact]
        public void CanToggleHeaderAndBandedRows() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTable table = slide.AddTable(2, 2);
                table.HeaderRow = false;
                table.BandedRows = false;
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointTable table = presentation.Slides.Single().Tables.First();
                Assert.False(table.HeaderRow);
                Assert.False(table.BandedRows);
            }

            using (PresentationDocument doc = PresentationDocument.Open(filePath, false)) {
                A.Table table = doc.PresentationPart!.SlideParts.First().Slide.Descendants<A.Table>().First();
                A.TableProperties? properties = table.TableProperties;
                Assert.NotNull(properties);
                Assert.False(properties!.FirstRow?.Value ?? false);
                Assert.False(properties.BandRow?.Value ?? false);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanAddTableWithStyleName() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string? styleName = null;
            string? styleId = null;

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointTableStyleInfo style = presentation.TableStyles
                    .FirstOrDefault(s => !string.IsNullOrWhiteSpace(s.StyleId));
                Assert.False(string.IsNullOrWhiteSpace(style.StyleId));

                styleId = style.StyleId;
                styleName = string.IsNullOrWhiteSpace(style.Name) ? style.StyleId : style.Name;

                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTable(rows: 2, columns: 2, styleName: styleName,
                    left: 0L, top: 0L, width: 6000000L, height: 2000000L,
                    firstRow: true, bandedRows: true);
                presentation.Save();
            }

            using (PresentationDocument doc = PresentationDocument.Open(filePath, false)) {
                A.Table table = doc.PresentationPart!.SlideParts.First().Slide.Descendants<A.Table>().First();
                string? appliedStyle = table.TableProperties?.GetFirstChild<A.TableStyleId>()?.Text;
                Assert.Equal(styleId, appliedStyle);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanSetTableCellAutoFit() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTable table = slide.AddTable(1, 1);
                PowerPointTableCell cell = table.GetCell(0, 0);
                cell.SetTextAutoFit(PowerPointTextAutoFit.Normal,
                    new PowerPointTextAutoFitOptions(fontScalePercent: 80, lineSpaceReductionPercent: 10));
                presentation.Save();
            }

            using (PresentationDocument doc = PresentationDocument.Open(filePath, false)) {
                A.TableCell cell = doc.PresentationPart!.SlideParts.First().Slide
                    .Descendants<A.TableCell>().First();
                A.BodyProperties? body = cell.TextBody?.GetFirstChild<A.BodyProperties>();
                A.NormalAutoFit? normal = body?.GetFirstChild<A.NormalAutoFit>();
                Assert.NotNull(normal);
                Assert.Equal(80000, normal!.FontScale!.Value);
                Assert.Equal(10000, normal.LineSpaceReduction!.Value);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanBindTableWithColumnDefinitions() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            var data = new[] {
                new SalesRow("Alpha", 12, 15),
                new SalesRow("Beta", 9, 11)
            };

            var columns = new[] {
                PowerPointTableColumn<SalesRow>.Create("Product", row => row.Product).WithWidthCm(4.0),
                PowerPointTableColumn<SalesRow>.Create("Q1", row => row.Q1),
                PowerPointTableColumn<SalesRow>.Create("Q2", row => row.Q2)
            };

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTable(data, columns, includeHeaders: true,
                    left: 0L, top: 0L, width: 6000000L, height: 2000000L);
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointTable table = presentation.Slides[0].Tables.First();
                Assert.Equal("Product", table.GetCell(0, 0).Text);
                Assert.Equal("Alpha", table.GetCell(1, 0).Text);
                Assert.Equal("15", table.GetCell(1, 2).Text);
                Assert.True(table.GetColumnWidth(0) > 0);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void AutoBoundTableUsesAvailableWidth() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            var data = new[] {
                new SalesRow("Alpha", 12, 15),
                new SalesRow("Beta", 9, 11)
            };

            const long tableWidth = 6000000L;

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTable(data, includeHeaders: true,
                    left: 0L, top: 0L, width: tableWidth, height: 2000000L);
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointTable table = presentation.Slides[0].Tables.First();
                long sum = 0;
                for (int i = 0; i < table.Columns; i++) {
                    long width = table.GetColumnWidth(i);
                    Assert.True(width > 0);
                    sum += width;
                }
                Assert.Equal(tableWidth, sum);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void ExplicitColumnWidthsCannotExceedTableWidth() {
            var data = new[] {
                new SalesRow("Alpha", 12, 15)
            };

            var columns = new[] {
                PowerPointTableColumn<SalesRow>.Create("Product", row => row.Product).WithWidth(1500000L),
                PowerPointTableColumn<SalesRow>.Create("Q1", row => row.Q1).WithWidth(1500000L)
            };

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                Assert.Throws<ArgumentException>(() =>
                    slide.AddTable(data, columns, includeHeaders: true, left: 0L, top: 0L, width: 2000000L, height: 1000000L));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanApplyColumnWidthRatios() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTable table = slide.AddTable(1, 2, left: 0L, top: 0L, width: 6000000L, height: 1500000L);
                table.SetColumnWidthsByRatio(2, 1);
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointTable table = presentation.Slides[0].Tables.First();
                Assert.Equal(4000000L, table.GetColumnWidth(0));
                Assert.Equal(2000000L, table.GetColumnWidth(1));
            }

            File.Delete(filePath);
        }

        private sealed class SalesRow {
            public SalesRow(string product, int q1, int q2) {
                Product = product;
                Q1 = q1;
                Q2 = q2;
            }

            public string Product { get; }
            public int Q1 { get; }
            public int Q2 { get; }
        }
    }

}

