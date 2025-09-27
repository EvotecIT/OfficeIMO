        [Fact]
        public void CanManipulateTableCellsAndPreserveStyle() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            const long tableWidth = 5_000_001L;
            const long tableHeight = 3_000_001L;

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTable table = slide.AddTable(2, 2, width: tableWidth, height: tableHeight);
                PowerPointTableCell cell = table.GetCell(0, 0);
                cell.Text = "Test";
                cell.FillColor = "FF0000";
                cell.Merge = (1, 2);
                table.AddRow();
            using (PresentationDocument doc = PresentationDocument.Open(filePath, false)) {
                A.Table table = doc.PresentationPart!.SlideParts.First().Slide.Descendants<A.Table>().First();
                string? styleId = table.TableProperties?.GetFirstChild<A.TableStyleId>()?.Text;
                Assert.Equal("{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}", styleId);

                long[] columnWidths = table.TableGrid!.Elements<A.GridColumn>().Select(column => column.Width?.Value ?? 0L)
                    .ToArray();
                Assert.Equal(new[] { tableWidth / 2 + tableWidth % 2, tableWidth / 2 }, columnWidths);
                Assert.Equal(tableWidth, columnWidths.Sum());

                long[] rowHeights = table.Elements<A.TableRow>().Select(row => row.Height?.Value ?? 0L).ToArray();
                Assert.Equal(new[] { tableHeight / 2 + tableHeight % 2, tableHeight / 2 }, rowHeights);
                Assert.Equal(tableHeight, rowHeights.Sum());
            }

            File.Delete(filePath);
        }

        [Fact]
        public void AddTable_WithSingleRowAndColumn_UsesProvidedDimensions() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            const long tableWidth = 1_234_567L;
            const long tableHeight = 765_432L;

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTable(1, 1, width: tableWidth, height: tableHeight);
                presentation.Save();
            }

            using (PresentationDocument doc = PresentationDocument.Open(filePath, false)) {
                A.Table table = doc.PresentationPart!.SlideParts.First().Slide.Descendants<A.Table>().First();

                long[] columnWidths = table.TableGrid!.Elements<A.GridColumn>().Select(column => column.Width?.Value ?? 0L)
                    .ToArray();
                Assert.Single(columnWidths, tableWidth);

                long[] rowHeights = table.Elements<A.TableRow>().Select(row => row.Height?.Value ?? 0L).ToArray();
                Assert.Single(rowHeights, tableHeight);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void AddTable_ExposesDistributedDimensionsViaProperties() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            const long tableWidth = 4_000_003L;
            const long tableHeight = 1_500_002L;
            const int columns = 4;
            const int rows = 3;

            long[] expectedColumnWidths = DistributeEvenly(tableWidth, columns);
            long[] expectedRowHeights = DistributeEvenly(tableHeight, rows);

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointTable table = presentation.AddSlide().AddTable(rows, columns, width: tableWidth, height: tableHeight);

                Assert.Equal(expectedColumnWidths, table.ColumnWidths);
                Assert.Equal(tableWidth, table.ColumnWidths.Sum());
                Assert.Equal(expectedRowHeights, table.RowHeights);
                Assert.Equal(tableHeight, table.RowHeights.Sum());

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointTable table = presentation.Slides[0].Tables.First();

                Assert.Equal(expectedColumnWidths, table.ColumnWidths);
                Assert.Equal(expectedRowHeights, table.RowHeights);
            }

            File.Delete(filePath);
        }

        private static long[] DistributeEvenly(long total, int parts) {
            long[] result = new long[parts];
            long baseValue = total / parts;
            long remainder = total % parts;

            for (int i = 0; i < parts; i++) {
                long adjustment = 0;
                if (remainder > 0) {
                    adjustment = 1;
                    remainder--;
                } else if (remainder < 0) {
                    adjustment = -1;
                    remainder++;
                }

                result[i] = baseValue + adjustment;
            }

            return result;
        }
    }
}
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
