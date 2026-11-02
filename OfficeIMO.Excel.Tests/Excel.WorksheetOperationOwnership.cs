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
            original.Cell(1, 1).SetValue("Region");
            original.Cell(2, 1).SetValue("Original");
            ExcelSheet other = document.AddSheet("Other");
            other.Cell(1, 1).SetValue("Region");
            other.Cell(2, 1).SetValue("Keep");
            using var saved = new MemoryStream();
            document.Save(saved);
            int destination = 1;
            Apply(original);

            document.RemoveWorksheet("Data");
            ExcelSheet replacement = document.AddSheet("Data");
            replacement.Cell(1, 1).SetValue("Region");
            replacement.Cell(2, 1).SetValue("Replacement");
            string[] names = document.GetSheetNames().ToArray();
            destination = 0;
            Assert.Throws<ArgumentException>(() => Apply(original));
            Assert.Equal(names, document.GetSheetNames());
            Assert.Equal("Replacement", replacement.Cell(2, 1).GetValue<string>());
            Assert.Null(replacement.Cell(3, 1).GetValue<string>());
            Assert.Equal("Keep", other.Cell(2, 1).GetValue<string>());
            if (operation == "merge-source") Assert.Equal("Original", other.Cell(3, 1).GetValue<string>());
            Assert.Null(other.Cell(4, 1).GetValue<string>());

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
