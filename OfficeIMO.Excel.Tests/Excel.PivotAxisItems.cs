using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests
{
    public partial class Excel
    {
        [Theory]
        [InlineData("row", true)]
        [InlineData("column", true)]
        [InlineData("page", true)]
        [InlineData("row", false)]
        public void PivotAxes_DefaultItemsReferenceCacheKeys(string axis, bool subtotal)
        {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Data");
            WritePivotSource(sheet, "Region", "Product", "Sales", 3);
            sheet.AddPivotTable("A1:C4", "E2", options: new ExcelPivotTableOptions { Name = "SalesPivot", RowFields = (axis == "row" ? new[] { "Region" } : null)is { } authoredRowFields1 ? System.Linq.Enumerable.ToList(authoredRowFields1) : System.Array.Empty<string>(), ColumnFields = (axis == "column" ? new[] { "Region" } : null)is { } authoredColumnFields2 ? System.Linq.Enumerable.ToList(authoredColumnFields2) : System.Array.Empty<string>(), PageFields = (axis == "page" ? new[] { "Region" } : null)is { } authoredPageFields3 ? System.Linq.Enumerable.ToList(authoredPageFields3) : System.Array.Empty<string>(), DataFields = new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum) }, FieldOptions = (subtotal ? null : new[] { new ExcelPivotFieldOptions("Region", defaultSubtotal: false) })is { } authoredFieldOptions4 ? System.Linq.Enumerable.ToList(authoredFieldOptions4) : System.Array.Empty<ExcelPivotFieldOptions>() });
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