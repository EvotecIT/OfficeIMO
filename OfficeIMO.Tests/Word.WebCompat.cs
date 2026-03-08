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
    }
}
