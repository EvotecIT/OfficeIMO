using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.SlideSize.SetSizePoints(400, 225);
                    PowerPointPaginatedTableResult result = presentation.AddTableSlides(source, columns,
                        new PowerPointTablePaginationOptions {
                            TableBounds = PowerPointLayoutBox.FromPoints(20, 40, 360, 100),
                            MinimumRowHeightPoints = 20,
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

                using PowerPointPresentation reopened = PowerPointPresentation.Load(filePath);
                Assert.Equal(4, reopened.Slides.Count);
                Assert.All(reopened.Slides, slide => Assert.Single(slide.Tables));
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void AddTableSlides_RejectsBoundsThatCannotFitHeaderAndData() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            var rows = new[] { new TableRow(1, "Only") };
            var columns = new[] { PowerPointTableColumn<TableRow>.Create("Name", row => row.Name) };

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                presentation.AddTableSlides(rows, columns, new PowerPointTablePaginationOptions {
                    TableBounds = PowerPointLayoutBox.FromPoints(10, 10, 200, 20),
                    MinimumRowHeightPoints = 20
                }));

            Assert.Contains("at least one data row", exception.Message, StringComparison.Ordinal);
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
