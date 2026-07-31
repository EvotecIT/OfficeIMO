using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_SecurityRegions_ManageAllowedEditAndIgnoredErrorsWithoutOpenXmlTypes() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.Protect();

            sheet.SetAllowedEditRange("Inputs", new[] { "A2:A5", "C2" }, password: "edit");
            sheet.AddIgnoredErrorRegion(
                new[] { "B2:B5" },
                ExcelIgnoredErrorKind.NumberStoredAsText | ExcelIgnoredErrorKind.FormulaRange);

            ExcelAllowedEditRangeInfo allowed = Assert.Single(sheet.GetAllowedEditRanges());
            Assert.Equal("Inputs", allowed.Name);
            Assert.Equal(new[] { "A2:A5", "C2" }, allowed.Ranges);
            Assert.True(allowed.IsPasswordProtected);

            ExcelIgnoredErrorRegionInfo ignored = Assert.Single(sheet.GetIgnoredErrorRegions());
            Assert.Equal(new[] { "B2:B5" }, ignored.Ranges);
            Assert.True(ignored.Errors.HasFlag(ExcelIgnoredErrorKind.NumberStoredAsText));
            Assert.True(ignored.Errors.HasFlag(ExcelIgnoredErrorKind.FormulaRange));

            sheet.InsertRows(3, 2);

            Assert.Equal(new[] { "A2:A7", "C2" }, Assert.Single(sheet.GetAllowedEditRanges()).Ranges);
            Assert.Equal(new[] { "B2:B7" }, Assert.Single(sheet.GetIgnoredErrorRegions()).Ranges);
            Assert.True(sheet.RemoveAllowedEditRange("inputs"));
            Assert.Equal(1, sheet.RemoveIgnoredErrorRegions("B4"));
            Assert.Empty(sheet.GetAllowedEditRanges());
            Assert.Empty(sheet.GetIgnoredErrorRegions());
        }

        [Fact]
        public void Test_SecurityRegions_RejectAllowedEditRangeBeforeProtection() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");

            var exception = Assert.Throws<System.InvalidOperationException>(() =>
                sheet.SetAllowedEditRange("Inputs", new[] { "A1" }));

            Assert.Contains("Protect", exception.Message);
            Assert.Empty(sheet.GetAllowedEditRanges());
        }

        [Fact]
        public void Test_SecurityRegions_CellDeletionRemovesFullyDeletedRegions() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.Protect();
            sheet.SetAllowedEditRange("OnlyCell", new[] { "A1" });
            sheet.AddIgnoredErrorRegion(new[] { "A1" }, ExcelIgnoredErrorKind.NumberStoredAsText);

            ExcelMutationResult result = sheet.DeleteCells("A1", ExcelCellShiftDirection.Left);

            Assert.True(result.PackageIsValid, string.Join(" | ", result.Diagnostics.Select(item => item.Message)));
            Assert.Empty(sheet.GetAllowedEditRanges());
            Assert.Empty(sheet.GetIgnoredErrorRegions());
        }
    }
}
