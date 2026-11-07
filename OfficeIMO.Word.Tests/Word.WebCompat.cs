using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData("valid", "1800", "2400")]
        [InlineData("missing", "2100", "2100")]
        [InlineData("invalid", "2100", "2100")]
        [InlineData("missing-before", "2100", "2100")]
        [InlineData("missing-after", "2100", "2100")]
        public void NormalizeForOnline_PreservesCompletePhysicalGridAndRepairsMissingGeometry(string geometry, string first, string second) {
            using var document = WordDocument.Create();
            WordTable table = document.AddTable(1, 2);
            Table raw = table._table;
            raw.GetFirstChild<TableProperties>().TableWidth = new TableWidth { Type = TableWidthUnitValues.Dxa, Width = "7200" };
            TableCell[] cells = raw.Elements<TableRow>().Single().Elements<TableCell>().ToArray();
            cells[0].TableCellProperties.GridSpan = new GridSpan { Val = 2 };
            cells[0].TableCellProperties.TableCellWidth = new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "4200" };
            cells[1].TableCellProperties.TableCellWidth = new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "3000" };
            table.GridColumnWidth = new[] { 1800, 2400, 3000 }
                .Select(width => global::OfficeIMO.OfficeLength.Emus((long)width * 635)).ToList();
            if (geometry == "missing") raw.GetFirstChild<TableGrid>().Remove();
            else if (geometry == "invalid") raw.GetFirstChild<TableGrid>().GetFirstChild<GridColumn>().Width = "-1";
            else if (geometry == "missing-before") raw.Elements<TableRow>().Single().TableRowProperties = new TableRowProperties(new GridBefore());
            else if (geometry == "missing-after") raw.Elements<TableRow>().Single().TableRowProperties = new TableRowProperties(new GridAfter());

            document.NormalizeTablesForOnline();
            using var stream = new MemoryStream(document.ToBytes());
            using var saved = WordprocessingDocument.Open(stream, false);
            Table output = saved.MainDocumentPart.Document.Body.Elements<Table>().Single();
            Assert.Equal(new[] { first, second, "3000" }, output.GetFirstChild<TableGrid>().Elements<GridColumn>().Select(column => column.Width.Value));
            Assert.Equal(2, output.Elements<TableRow>().Single().Elements<TableCell>().First().TableCellProperties.GridSpan.Val.Value);
        }

        [Fact]
        public void Test_NormalizeTablesForOnline_UpdatesBodyHeaderAndFooterTables() {
            using var document = WordDocument.Create();
            var bodyTable = document.AddTable(2, 2);
            bodyTable.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(5000) * 635)));

            var header = document.Sections[0].PageSetup.Header;
            var headerTable = header.AddTable(1, 2);
            headerTable.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(4000) * 635)));

            var footer = document.Sections[0].PageSetup.Footer;
            var footerTable = footer.AddTable(1, 2);
            footerTable.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(4000) * 635)));

            document.NormalizeTablesForOnline();
            using var saved = new MemoryStream();
            document.Save(saved);

            var bodyGrid = document._wordprocessingDocument.MainDocumentPart!.Document.Body!.Descendants<TableGrid>().FirstOrDefault();
            Assert.NotNull(bodyGrid);
            Assert.Equal(2, bodyGrid!.Elements<GridColumn>().Count());

            var headerGrid = document._wordprocessingDocument.MainDocumentPart.HeaderParts
                .SelectMany(part => part.Header?.Descendants<TableGrid>() ?? Enumerable.Empty<TableGrid>())
                .FirstOrDefault();
            Assert.NotNull(headerGrid);
            Assert.Equal(2, headerGrid!.Elements<GridColumn>().Count());

            var footerGrid = document._wordprocessingDocument.MainDocumentPart.FooterParts
                .SelectMany(part => part.Footer?.Descendants<TableGrid>() ?? Enumerable.Empty<TableGrid>())
                .FirstOrDefault();
            Assert.NotNull(footerGrid);
            Assert.Equal(2, footerGrid!.Elements<GridColumn>().Count());
        }

        [Fact]
        public void Test_NormalizeTablesForOnline_ConvertsMergedCellsAndNormalizesNestedTables() {
            using var document = WordDocument.Create();

            var bodyTable = document.AddTable(2, 3);
            bodyTable.Width = WordTableWidth.Percent((5000) / 50d);
            bodyTable.Rows[0].Cells[0].Paragraphs[0].Text = "Merged";
            bodyTable.Rows[0].Cells[1].Paragraphs[0].Text = "Removed after merge";
            bodyTable.Rows[0].Cells[0].MergeHorizontally(1, copyParagraphs: true);

            var nestedTable = bodyTable.Rows[1].Cells[0].AddTable(1, 3);
            nestedTable.Width = WordTableWidth.Percent((5000) / 50d);
            nestedTable.Rows[0].Cells[0].Paragraphs[0].Text = "Nested merged";
            nestedTable.Rows[0].Cells[1].Paragraphs[0].Text = "Nested hidden";
            nestedTable.Rows[0].Cells[0].MergeHorizontally(1, copyParagraphs: true);

            var header = document.Sections[0].PageSetup.Header;
            var headerTable = header.AddTable(1, 3);
            headerTable.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(4500) * 635)));
            headerTable.Rows[0].Cells[0].Paragraphs[0].Text = "Header merged";
            headerTable.Rows[0].Cells[1].Paragraphs[0].Text = "Header hidden";
            headerTable.Rows[0].Cells[0].MergeHorizontally(1, copyParagraphs: true);

            document.NormalizeTablesForOnline();
            using var saved = new MemoryStream();
            document.Save(saved);

            var bodyMergedTable = document._wordprocessingDocument.MainDocumentPart!.Document.Body!
                .Descendants<Table>()
                .First(table => !table.Ancestors<TableCell>().Any());
            AssertMergedTable(bodyMergedTable, expectedGridColumns: 3);

            var nestedTableXml = document._wordprocessingDocument.MainDocumentPart!.Document.Body!
                .Descendants<Table>()
                .First(table => table.Ancestors<TableCell>().Any());
            AssertMergedTable(nestedTableXml, expectedGridColumns: 3);

            var hostCellMargins = nestedTableXml.Ancestors<TableCell>().First().TableCellProperties?.TableCellMargin;
            Assert.Equal("72", hostCellMargins?.TopMargin?.Width?.Value);
            Assert.Equal("72", hostCellMargins?.BottomMargin?.Width?.Value);

            var headerMergedTable = document._wordprocessingDocument.MainDocumentPart.HeaderParts
                .SelectMany(part => part.Header?.Descendants<Table>() ?? Enumerable.Empty<Table>())
                .First();
            AssertMergedTable(headerMergedTable, expectedGridColumns: 3);
        }

        [Theory]
        [InlineData(0, 240, "0", "240")]
        [InlineData(120, 0, "120", "0")]
        [InlineData(0, null, "0", "72")]
        [InlineData(null, 240, "72", "240")]
        public void Test_NormalizeTablesForOnline_PreservesExplicitNestedHostMargins(
            int? top, int? bottom, string expectedTop, string expectedBottom) {
            using var document = WordDocument.Create();
            WordTableCell host = document.AddTable(1, 1).Rows[0].Cells[0];
            WordTable nested = host.AddTable(1, 1);
            nested.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(2400) * 635)));
            host.MarginTop= top.HasValue ? global::OfficeIMO.OfficeLength.Points((checked((short)top.Value)) / 20D) : null;
            host.MarginBottom= bottom.HasValue ? global::OfficeIMO.OfficeLength.Points((checked((short)bottom.Value)) / 20D) : null;

            document.NormalizeTablesForOnline();
            using var saved = new MemoryStream();
            document.Save(saved);
            saved.Position = 0;
            using var package = WordprocessingDocument.Open(saved, false);
            Table nestedXml = Assert.Single(package.MainDocumentPart!.Document.Body!.Descendants<Table>(),
                table => table.Ancestors<TableCell>().Any());
            TableCellMargin? margins = nestedXml.Ancestors<TableCell>().First().TableCellProperties?.TableCellMargin;

            Assert.Equal(expectedTop, margins?.TopMargin?.Width?.Value);
            Assert.Equal(expectedBottom, margins?.BottomMargin?.Width?.Value);
            Assert.Equal(TableWidthUnitValues.Dxa, margins?.TopMargin?.Type?.Value);
            Assert.Equal(TableWidthUnitValues.Dxa, margins?.BottomMargin?.Type?.Value);
        }

        [Fact]
        public void Test_NormalizeTablesForOnline_PreservesExplicitNestedTableTopBorderSuppression() {
            using var document = WordDocument.Create();
            WordTable nested = document.AddTable(1, 1).Rows[0].Cells[0].AddTable(1, 1);
            nested.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(2400) * 635)));
            nested.StyleDetails!.SetCustomBorders(topStyle: WordBorderStyle.Nil);

            document.NormalizeTablesForOnline();
            using var saved = new MemoryStream();
            document.Save(saved);
            saved.Position = 0;
            using var package = WordprocessingDocument.Open(saved, false);
            Table nestedXml = Assert.Single(package.MainDocumentPart!.Document.Body!.Descendants<Table>(),
                table => table.Ancestors<TableCell>().Any());
            TableCell cell = Assert.Single(Assert.Single(nestedXml.Elements<TableRow>()).Elements<TableCell>());

            Assert.Equal(BorderValues.Nil, nestedXml.GetFirstChild<TableProperties>()?.TableBorders?.TopBorder?.Val?.Value);
            Assert.Null(cell.TableCellProperties?.TableCellBorders?.TopBorder);
        }

        [Fact]
        public void Test_NormalizeTablesForOnline_PreservesExplicitNestedCellTopBorderSuppression() {
            using var document = WordDocument.Create();
            WordTable nested = document.AddTable(1, 1).Rows[0].Cells[0].AddTable(1, 2);
            nested.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(2400) * 635)));
            nested.Rows[0].Cells[0].Borders.TopStyle = WordBorderStyle.Nil;

            document.NormalizeTablesForOnline();
            using var saved = new MemoryStream();
            document.Save(saved);
            saved.Position = 0;
            using var package = WordprocessingDocument.Open(saved, false);
            Table nestedXml = Assert.Single(package.MainDocumentPart!.Document.Body!.Descendants<Table>(),
                table => table.Ancestors<TableCell>().Any());
            TableCell[] cells = Assert.Single(nestedXml.Elements<TableRow>()).Elements<TableCell>().ToArray();

            Assert.Equal(2, cells.Length);
            Assert.Equal(BorderValues.Nil, cells[0].TableCellProperties?.TableCellBorders?.TopBorder?.Val?.Value);
            Assert.Equal(BorderValues.Single, cells[1].TableCellProperties?.TableCellBorders?.TopBorder?.Val?.Value);
            Assert.Equal(4U, cells[1].TableCellProperties?.TableCellBorders?.TopBorder?.Size?.Value);
        }

        private static void AssertMergedTable(Table table, int expectedGridColumns) {
            var firstRowCells = table.Elements<TableRow>().First().Elements<TableCell>().ToList();
            Assert.Equal(expectedGridColumns - 1, firstRowCells.Count);
            Assert.Equal(expectedGridColumns, table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count());

            var mergedCellProperties = firstRowCells[0].TableCellProperties;
            Assert.Equal(2, (int?)mergedCellProperties?.GetFirstChild<GridSpan>()?.Val?.Value);
            Assert.DoesNotContain(table.Descendants<HorizontalMerge>(), merge => merge.Val?.Value == MergedCellValues.Continue);
        }
    }
}
