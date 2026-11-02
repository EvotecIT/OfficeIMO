using System;
using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void ActiveWorksheetObjectMustRemainInItsWorkbook(bool cacheSheets) {
            using var document = ExcelDocument.Create();
            document.SheetCachingEnabled = cacheSheets;
            var original = document.AddSheet("Data");
            var other = document.AddSheet("Other");
            document.SetActiveWorksheet(other);
            using var stream = new MemoryStream();
            document.Save(stream);
            document.SetActiveWorksheet(original);
            Assert.Equal("Data", document.ActiveWorksheet!.Name);

            using var foreign = ExcelDocument.Create();
            var foreignSheet = foreign.AddSheet("Other");
            Assert.Throws<ArgumentException>(() => document.SetActiveWorksheet(foreignSheet));
            Assert.Equal("Data", document.ActiveWorksheet!.Name);

            document.RemoveWorksheet("Data");
            document.AddSheet("Data");
            document.SetActiveWorksheet("Other");
            Assert.Throws<ArgumentException>(() => document.SetActiveWorksheet(original));
            Assert.Equal("Other", document.ActiveWorksheet!.Name);
            using var saved = new MemoryStream();
            document.Save(saved);
            saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal("Other", reopened.ActiveWorksheet!.Name);
        }
    }
}
