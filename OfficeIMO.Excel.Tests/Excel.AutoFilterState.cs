using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_AutoFilterState_ReadsValuesCustomBlankAndTopCriteria() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "Amount");
            sheet.CellValue(1, 3, "Rank");
            sheet.CellValue(2, 1, "EU");
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(2, 3, 1);
            sheet.CellValue(3, 1, "US");
            sheet.CellValue(3, 2, 20);
            sheet.CellValue(3, 3, 2);
            sheet.AutoFilterAdd("A1:C3");
            sheet.AutoFilterByHeaderEquals("Region", new[] { "EU", "US" });
            sheet.AutoFilterByHeaderBetween("Amount", 10, 20);
            sheet.AutoFilterTopBottom("A1:C3", 2, 1, top: true, percent: false);

            ExcelAutoFilterInfo state = Assert.Single(sheet.GetAutoFilters());
            Assert.Equal("A1:C3", state.Range);
            Assert.False(state.IsTableFilter);
            Assert.Equal(3, state.Columns.Count);

            ExcelAutoFilterColumnInfo values = state.Columns[0];
            Assert.Equal(ExcelAutoFilterCriteriaKind.Values, values.Kind);
            Assert.Equal(new[] { "EU", "US" }, values.Values);

            ExcelAutoFilterColumnInfo custom = state.Columns[1];
            Assert.Equal(ExcelAutoFilterCriteriaKind.Custom, custom.Kind);
            Assert.True(custom.MatchAll);
            Assert.Equal(new[] { "10", "20" }, custom.Conditions.Select(condition => condition.Value).ToArray());

            ExcelAutoFilterColumnInfo top = state.Columns[2];
            Assert.Equal(ExcelAutoFilterCriteriaKind.TopBottom, top.Kind);
            Assert.True(top.Top);
            Assert.False(top.Percent);
            Assert.Equal(1d, top.TopValue);
            Assert.True(sheet.ClearAutoFilterColumn(2));
            Assert.Equal(2, Assert.Single(sheet.GetAutoFilters()).Columns.Count);
        }
    }
}
