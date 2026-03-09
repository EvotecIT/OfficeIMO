using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_NormalizeTablesForOnline_UpdatesBodyHeaderAndFooterTables() {
            using var document = WordDocument.Create();
            var bodyTable = document.AddTable(2, 2);
            bodyTable.WidthType = TableWidthUnitValues.Dxa;
            bodyTable.Width = 5000;

            var header = document.Sections[0].GetOrCreateHeader(HeaderFooterValues.Default);
            var headerTable = header.AddTable(1, 2);
            headerTable.WidthType = TableWidthUnitValues.Dxa;
            headerTable.Width = 4000;

            var footer = document.Sections[0].GetOrCreateFooter(HeaderFooterValues.Default);
            var footerTable = footer.AddTable(1, 2);
            footerTable.WidthType = TableWidthUnitValues.Dxa;
            footerTable.Width = 4000;

            document.NormalizeTablesForOnline();
            document.Save();

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
            bodyTable.WidthType = TableWidthUnitValues.Pct;
            bodyTable.Width = 5000;
            bodyTable.Rows[0].Cells[0].Paragraphs[0].Text = "Merged";
            bodyTable.Rows[0].Cells[1].Paragraphs[0].Text = "Removed after merge";
            bodyTable.Rows[0].Cells[0].MergeHorizontally(1, copyParagraphs: true);

            var nestedTable = bodyTable.Rows[1].Cells[0].AddTable(1, 3);
            nestedTable.WidthType = TableWidthUnitValues.Pct;
            nestedTable.Width = 5000;
            nestedTable.Rows[0].Cells[0].Paragraphs[0].Text = "Nested merged";
            nestedTable.Rows[0].Cells[1].Paragraphs[0].Text = "Nested hidden";
            nestedTable.Rows[0].Cells[0].MergeHorizontally(1, copyParagraphs: true);

            var header = document.Sections[0].GetOrCreateHeader(HeaderFooterValues.Default);
            var headerTable = header.AddTable(1, 3);
            headerTable.WidthType = TableWidthUnitValues.Dxa;
            headerTable.Width = 4500;
            headerTable.Rows[0].Cells[0].Paragraphs[0].Text = "Header merged";
            headerTable.Rows[0].Cells[1].Paragraphs[0].Text = "Header hidden";
            headerTable.Rows[0].Cells[0].MergeHorizontally(1, copyParagraphs: true);

            document.NormalizeTablesForOnline();
            document.Save();

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
