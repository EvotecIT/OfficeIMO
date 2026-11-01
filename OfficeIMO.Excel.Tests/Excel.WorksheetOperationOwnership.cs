using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("reorder", true)]
        [InlineData("reorder", false)]
        [InlineData("merge-source", true)]
        [InlineData("merge-source", false)]
        public void WorksheetOperationArgumentsMustRemainInTheirWorkbook(string operation, bool cacheSheets) {
            using var document = ExcelDocument.Create();
            document.SheetCachingEnabled = cacheSheets;
            ExcelSheet original = document.AddSheet("Data");
            original.CellValue(1, 1, "Region");
            original.CellValue(2, 1, "Original");
            ExcelSheet other = document.AddSheet("Other");
            other.CellValue(1, 1, "Region");
            other.CellValue(2, 1, "Keep");
            using var saved = new MemoryStream();
            document.Save(saved);
            int destination = 1;
            Apply(original);

            document.RemoveWorksheet("Data");
            ExcelSheet replacement = document.AddSheet("Data");
            replacement.CellValue(1, 1, "Region");
            replacement.CellValue(2, 1, "Replacement");
            string[] names = document.GetSheetNames().ToArray();
            destination = 0;
            Assert.Throws<ArgumentException>(() => Apply(original));
            Assert.Equal(names, document.GetSheetNames());
            Assert.Equal("Replacement", replacement.CellAt(2, 1).GetValue<string>());
            Assert.Null(replacement.CellAt(3, 1).GetValue<string>());
            Assert.Equal("Keep", other.CellAt(2, 1).GetValue<string>());
            if (operation == "merge-source") Assert.Equal("Original", other.CellAt(3, 1).GetValue<string>());
            Assert.Null(other.CellAt(4, 1).GetValue<string>());

            void Apply(ExcelSheet sheet) {
                switch (operation) {
                    case "reorder": document.ReorderWorksheet(sheet, destination); break;
                    case "merge-source": document.MergeWorksheets(other, sheet); break;
                    default: throw new InvalidOperationException("Unknown worksheet operation fixture.");
                }
            }
        }
    }
}
