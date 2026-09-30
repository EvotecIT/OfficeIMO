using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("row", true)]
        [InlineData("column", true)]
        [InlineData("page", true)]
        [InlineData("row", false)]
        public void PivotAxes_DefaultItemsReferenceCacheKeys(string axis, bool subtotal) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            WritePivotSource(sheet, "Region", "Product", "Sales", 3);
            sheet.AddPivotTable("A1:C4", "E2", "SalesPivot",
                rowFields: axis == "row" ? new[] { "Region" } : null,
                columnFields: axis == "column" ? new[] { "Region" } : null,
                pageFields: axis == "page" ? new[] { "Region" } : null,
                dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) },
                fieldOptions: subtotal ? null : new[] { new ExcelPivotFieldOptions("Region", defaultSubtotal: false) });
            var part = Assert.Single(sheet.WorksheetPart.PivotTableParts);
            var fields = part.PivotTableDefinition!.PivotFields!.Elements<PivotField>().ToArray();
            var cache = part.PivotTableCacheDefinitionPart!.PivotCacheDefinition!.CacheFields!.Elements<CacheField>().First();
            var items = fields[0].Items!.Elements<Item>().ToArray();
            Assert.Equal(cache.SharedItems!.ChildElements.Count, items.Count(item => item.Index != null));
            Assert.Equal(new uint[] { 0, 1 }, items.Where(item => item.Index != null).Select(item => item.Index!.Value));
            Assert.Equal(subtotal, items.Any(item => item.ItemType?.Value == ItemValues.Default));
            Assert.Null(fields[2].Items);
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
