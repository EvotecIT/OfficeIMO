using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointTablePaginationTests {
        [Fact]
        public void AddTableSlides_PreservesEveryRowAndRepeatsHeaders() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                TableRow[] source = Enumerable.Range(1, 13)
                    .Select(index => new TableRow(index, "Item " + index))
                    .ToArray();
                var columns = new[] {
                    PowerPointTableColumn<TableRow>.Create("Id", row => row.Id),
                    PowerPointTableColumn<TableRow>.Create("Name", row => row.Name)
                };
                var configuredPages = new List<int>();

                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(400),OfficeIMO.OfficeLength.Points(225));
                    PowerPointPaginatedTableResult result = presentation.AddTableSlides(source, columns,
                        new PowerPointTablePaginationOptions {
                            TableBounds = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Points(20),OfficeIMO.OfficeLength.Points(40),OfficeIMO.OfficeLength.Points(360),OfficeIMO.OfficeLength.Points(100)),
                            MinimumRowHeight = global::OfficeIMO.OfficeLength.Points(20),
                            ConfigureSlide = (slide, context) => {
                                configuredPages.Add(context.PageIndex);
                                slide.AddTextBoxPoints(
                                    "Inventory " + (context.PageIndex + 1) + "/" + context.PageCount,
                                    20, 10, 250, 24);
                            }
                        });

                    Assert.Equal(4, result.PageCount);
                    Assert.Equal(new[] { 0, 1, 2, 3 }, configuredPages);
                    Assert.All(result.Tables, table => {
                        Assert.True(table.HeaderRow);
                        Assert.Equal("Id", table.GetCell(0, 0).Text);
                        Assert.Equal("Name", table.GetCell(0, 1).Text);
                    });
                    string[] representedNames = result.Tables
                        .SelectMany(table => Enumerable.Range(1, table.Rows - 1)
                            .Select(row => table.GetCell(row, 1).Text))
                        .ToArray();
                    Assert.Equal(source.Select(row => row.Name), representedNames);
                    Assert.Equal(13, result.SourceRowCount);
                    presentation.Save();
                }

                using PowerPointDocument reopened = PowerPointDocument.Load(filePath);
                Assert.Equal(4, reopened.Slides.Count);
                Assert.All(reopened.Slides, slide => Assert.Single(slide.Tables));
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void AddTableSlides_RejectsBoundsThatCannotFitHeaderAndData() {
            using var stream = new MemoryStream();
            using PowerPointDocument presentation = PowerPointDocument.Create(stream);
            var rows = new[] { new TableRow(1, "Only") };
            var columns = new[] { PowerPointTableColumn<TableRow>.Create("Name", row => row.Name) };

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                presentation.AddTableSlides(rows, columns, new PowerPointTablePaginationOptions {
                    TableBounds = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Points(10),OfficeIMO.OfficeLength.Points(10),OfficeIMO.OfficeLength.Points(200),OfficeIMO.OfficeLength.Points(20)),
                    MinimumRowHeight = global::OfficeIMO.OfficeLength.Points(20)
                }));

            Assert.Contains("at least one data row", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void AddTableSlides_PreservesWholeEmuRowHeightsAfterSaving() {
            using var stream = new MemoryStream();
            using (PowerPointDocument document = PowerPointDocument.Create(stream)) {
                var rows = Enumerable.Range(1, 5).Select(index => new TableRow(index, "Row " + index));
                var columns = new[] { PowerPointTableColumn<TableRow>.Create("Id", row => row.Id) };
                OfficeLength rowHeight = OfficeLength.Emus(304801);
                var options = new PowerPointTablePaginationOptions {
                    TableBounds = new PowerPointLayoutBox(OfficeLength.Points(20), OfficeLength.Points(30),
                        OfficeLength.Inches(4), OfficeLength.Emus(914403)),
                    MinimumRowHeight = rowHeight
                };
                PowerPointPaginatedTableResult result = document.AddTableSlides(rows, columns, options);
                Assert.Equal(3, result.PageCount);
                Assert.Equal(new[] { 3, 3, 2 }, result.Tables.Select(table => table.Rows));
                Assert.All(result.Tables, table => {
                    Assert.Equal(OfficeLength.Points(20), table.Left);
                    Assert.Equal(OfficeLength.Points(30), table.Top);
                    Assert.All(table.RowItems, row => Assert.Equal(rowHeight, row.Height));
                });
                Assert.Throws<ArgumentOutOfRangeException>(() => options.MinimumRowHeight = OfficeLength.Zero);
                Assert.Equal(rowHeight, options.MinimumRowHeight);
                document.Save();
            }
            stream.Position = 0;
            using PowerPointDocument reopened = PowerPointDocument.Load(stream);
            Assert.Equal(new long[] { 914403, 914403, 609602 },
                reopened.Slides.Select(slide => slide.Tables.Single().Height.ToEmus()));
            Assert.All(reopened.Slides, slide => Assert.All(slide.Tables.Single().RowItems,
                row => Assert.Equal(OfficeLength.Emus(304801), row.Height)));
        }

        private sealed class TableRow {
            internal TableRow(int id, string name) {
                Id = id;
                Name = name;
            }

            internal int Id { get; }
            internal string Name { get; }
        }
    }
}
