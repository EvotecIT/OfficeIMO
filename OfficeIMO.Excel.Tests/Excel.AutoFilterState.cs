using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_AutoFilterState_ReadsValuesCustomBlankAndTopCriteria() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Region");
            sheet.Cell(1, 2).SetValue("Amount");
            sheet.Cell(1, 3).SetValue("Rank");
            sheet.Cell(2, 1).SetValue("EU");
            sheet.Cell(2, 2).SetValue(10);
            sheet.Cell(2, 3).SetValue(1);
            sheet.Cell(3, 1).SetValue("US");
            sheet.Cell(3, 2).SetValue(20);
            sheet.Cell(3, 3).SetValue(2);
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
