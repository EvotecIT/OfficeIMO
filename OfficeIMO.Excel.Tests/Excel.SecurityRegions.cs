using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
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

        [Fact]
        public void Test_SecurityRegions_RejectOtherWorksheetQualifiers() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.AddWorksheet("Other");
            sheet.Protect();

            Assert.Throws<System.ArgumentException>(() => sheet.SetAllowedEditRange("Inputs", new[] { "Other!A1" }));
            Assert.Throws<System.ArgumentException>(() => sheet.AddIgnoredErrorRegion(new[] { "Other!A1" }, ExcelIgnoredErrorKind.NumberStoredAsText));
            sheet.SetAllowedEditRange("Inputs", new[] { "'Data'!A1" });
            Assert.Equal("A1", Assert.Single(Assert.Single(sheet.GetAllowedEditRanges()).Ranges));
        }

        [Fact]
        public void Test_SecurityRegions_ReplacementClearsModernPasswordHash() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.Protect();
            sheet.SetAllowedEditRange("Inputs", new[] { "A1" });
            ProtectedRange range = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<ProtectedRange>());
            range.AlgorithmName = "SHA-512";
            range.HashValue = "AQID";
            range.SaltValue = "BAUG";
            range.SpinCount = 1000U;

            sheet.SetAllowedEditRange("Inputs", new[] { "B2" });

            range = Assert.Single(sheet.WorksheetPart.Worksheet.Descendants<ProtectedRange>());
            Assert.Null(range.AlgorithmName);
            Assert.Null(range.HashValue);
            Assert.Null(range.SaltValue);
            Assert.Null(range.SpinCount);
            Assert.False(Assert.Single(sheet.GetAllowedEditRanges()).IsPasswordProtected);
        }
    }
}
